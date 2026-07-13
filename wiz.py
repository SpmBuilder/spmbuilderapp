import streamlit as st
import pandas as pd
import numpy as np
import json
import re
from pathlib import Path
import io
import base64
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Multi-Sheet Data Processing Wizard",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .reference-box {
        background-color: #e7f3ff;
        padding: 15px;
        border-radius: 8px;
        border-left: 4px solid #2196F3;
        margin: 10px 0;
    }
    .operation-active {
        border: 2px solid #1f77b4;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
        background-color: #f8f9fa;
    }
    .automation-box {
        border: 2px solid #6f42c1;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
        background-color: #f5f0ff;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Session state initialization
# ---------------------------------------------------------------------------
defaults = {
    'data_sources': {},          # {name: df}
    'source_metadata': {},       # {name: {file, sheet, rows, columns}}
    'operations': [],            # recorded steps for the active source
    'active_source': None,
    'relationships': [],
    'current_operation': None,
    'loaded_recipe': None,       # recipe parsed from an uploaded json
    'recipe_mapping': {},        # {old_source_name: new_source_name}
    'template_headers': [],      # ordered list of output column headers
    'source_snapshots': {},      # {name: baseline df the current operations list replays from}
    'operations_undo_stack': [], # list of previous 'operations' lists, for Undo
    'batch_results': {},         # {result_name: df} from the last Batch Automation run
    'filter_conditions': [],     # For multiple filter operations
    'filter_logic': 'AND',       # How multiple filter conditions are combined
}
for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ---------------------------------------------------------------------------
# Formula engine helpers (safer + actually correct)
# ---------------------------------------------------------------------------
def _f_upper(x):
    return x.str.upper() if hasattr(x, "str") else str(x).upper()


def _f_lower(x):
    return x.str.lower() if hasattr(x, "str") else str(x).lower()


def _f_len(x):
    if hasattr(x, "str"):
        return x.str.len()
    return pd.Series([len(str(v)) for v in x]) if hasattr(x, "__iter__") else len(str(x))


SAFE_NAMESPACE_FUNCS = {
    "abs": abs,
    "round": round,
    "upper": _f_upper,
    "lower": _f_lower,
    "len": _f_len,
    "np": np,
    "pd": pd,
}


def _replace_columns(expr, df):
    """Replace [Column Name] with df['Column Name']."""
    def repl(match):
        col_name = match.group(1)
        return f"df['{col_name}']"
    return re.sub(r"\[([^\]]+)\]", repl, expr)


def _xlookup(lookup_value, ref_source_name, lookup_column, return_column, if_not_found=None, sources=None):
    """Excel-XLOOKUP-style lookup against another loaded data source.
    Works on a whole Series (vectorized) or a single scalar value.
    """
    sources = sources or {}
    ref_df = sources.get(ref_source_name)
    if ref_df is None:
        raise Exception(f"Reference source '{ref_source_name}' is not loaded")
    if lookup_column not in ref_df.columns:
        raise Exception(f"'{lookup_column}' not found in reference source '{ref_source_name}'")
    if return_column not in ref_df.columns:
        raise Exception(f"'{return_column}' not found in reference source '{ref_source_name}'")

    mapping = dict(zip(ref_df[lookup_column], ref_df[return_column]))

    if hasattr(lookup_value, "map"):  # pandas Series -> vectorized lookup
        result = lookup_value.map(mapping)
        if if_not_found is not None:
            result = result.fillna(if_not_found)
        return result
    return mapping.get(lookup_value, if_not_found)


def detect_referenced_source_names(text, known_source_names):
    """Best-effort scan of a formula/text for quoted names of currently loaded sources,
    so recipes know which reference files a formula depends on (for automation remapping).
    """
    found = []
    for name in known_source_names:
        if f"'{name}'" in text or f'"{name}"' in text:
            found.append(name)
    return found


def try_convert_datetime_column(series):
    """Return a converted datetime Series if `series` looks genuinely date-like,
    otherwise return None. Much stricter than a bare pd.to_datetime(errors='coerce')
    check: requires the column to be text (not already numeric), and requires a
    strong majority of non-null values to actually parse as dates. This avoids the
    classic false positive where a single stray value in a text/ID/category column
    gets misread as a date, silently turning the rest of the column into NaT.
    """
    # Never treat an already-numeric column as a date candidate (pandas will
    # happily "parse" plain numbers as timestamps, which is almost never what
    # the user means).
    if pd.api.types.is_numeric_dtype(series):
        return None
    if pd.api.types.is_datetime64_any_dtype(series):
        return series

    non_null = series.dropna()
    if len(non_null) == 0:
        return None

    # If everything looks like a plain number (possibly with commas/decimals),
    # this is not a date column even though to_datetime could coerce it.
    numeric_like = pd.to_numeric(non_null.astype(str).str.replace(',', '', regex=False), errors='coerce')
    if numeric_like.notna().mean() >= 0.9:
        return None

    try:
        parsed = pd.to_datetime(non_null, errors='coerce')
    except Exception:
        return None

    parse_rate = parsed.notna().mean()
    if parse_rate >= 0.9:
        return pd.to_datetime(series, errors='coerce')
    return None


def get_unique_values_sorted(df, column, max_values=1000):
    """Get unique values from a column, sorted, with limit for performance."""
    try:
        values = df[column].dropna().unique()
        if len(values) > max_values:
            # For large datasets, show a sample
            values = sorted(values)[:max_values]
            return values, True
        return sorted(values), False
    except:
        return [], False


class MultiSheetProcessor:
    """Handles multiple sheets and files with references"""

    @staticmethod
    def load_multi_sheet_file(uploaded_file):
        try:
            excel_file = pd.ExcelFile(uploaded_file)
            sheets = {}

            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                df.columns = [str(c).strip() for c in df.columns]

                # Automatically detect genuinely date-like columns (strict check,
                # see try_convert_datetime_column - avoids false positives on
                # text/ID/category columns).
                for col in df.columns:
                    converted = try_convert_datetime_column(df[col])
                    if converted is not None:
                        df[col] = converted

                sheets[sheet_name] = df

            return sheets

        except Exception as e:
            st.error(f"Error loading Excel file: {e}")
            return None


    @staticmethod
    def load_single_sheet(uploaded_file):
        try:
            ext = Path(uploaded_file.name).suffix.lower()

            if ext == ".csv":
                df = pd.read_csv(uploaded_file)
                df.columns = [str(c).strip() for c in df.columns]

                for col in df.columns:
                    converted = try_convert_datetime_column(df[col])
                    if converted is not None:
                        df[col] = converted

                return {"Sheet1": df}

            elif ext in [".xlsx", ".xls"]:
                return MultiSheetProcessor.load_multi_sheet_file(uploaded_file)

            st.error("Unsupported file type")
            return None

        except Exception as e:
            st.error(f"Error loading file: {e}")
            return None

    @staticmethod
    def get_filter_mask(df, column, operation, value):
        """Compute a boolean mask (aligned to df.index) for a single condition,
        without narrowing the dataframe. Used both by filter_rows (single
        condition) and filter_multiple_rows (so multiple conditions can be
        combined with AND/OR against the same original data, instead of
        narrowing the data step by step - which makes OR-style combinations,
        e.g. "Status == A OR Status == B", impossible).
        Returns None if the condition can't be evaluated (column missing,
        bad value, etc.) - callers should treat None as "no filtering".
        """
        if column not in df.columns:
            st.error(f"Column '{column}' not found!")
            return None

        series = df[column]

        # Strict date-detection: only treat as a date column if a strong
        # majority of values genuinely parse as dates (see try_convert_datetime_column).
        converted_dates = try_convert_datetime_column(series)
        is_date = converted_dates is not None
        if is_date:
            series = converted_dates

        if is_date:
            try:
                compare = pd.to_datetime(value, errors='coerce')
                if pd.isna(compare):
                    st.error("Invalid date format. Use YYYY-MM-DD or similar.")
                    return None
            except Exception:
                st.error("Invalid date format. Use YYYY-MM-DD or similar.")
                return None

            series_normalized = series.dt.normalize()
            compare_normalized = compare.normalize()

            if operation == "==":
                return series_normalized == compare_normalized
            elif operation == "!=":
                return series_normalized != compare_normalized
            elif operation == ">":
                return series > compare
            elif operation == "<":
                return series < compare
            elif operation == ">=":
                return series >= compare
            elif operation == "<=":
                return series <= compare
            elif operation == "contains":
                return series.dt.strftime("%Y-%m-%d").str.contains(str(value), case=False, na=False)
            return None

        if pd.api.types.is_numeric_dtype(series):
            try:
                num_value = float(value)
            except ValueError:
                st.error("Please enter a numeric value for this column")
                return None

            if operation == "==":
                return series == num_value
            elif operation == "!=":
                return series != num_value
            elif operation == ">":
                return series > num_value
            elif operation == "<":
                return series < num_value
            elif operation == ">=":
                return series >= num_value
            elif operation == "<=":
                return series <= num_value
            return None

        # Text columns - .strip() guards against stray Excel whitespace.
        series = series.astype(str).str.strip()
        value_str = str(value).strip()

        if operation == "contains":
            return series.str.contains(value_str, case=False, na=False)
        elif operation == "==":
            return series.str.lower() == value_str.lower()
        elif operation == "!=":
            return series.str.lower() != value_str.lower()
        elif operation == ">":
            return series > value_str
        elif operation == "<":
            return series < value_str
        elif operation == ">=":
            return series >= value_str
        elif operation == "<=":
            return series <= value_str
        return None

    @staticmethod
    def filter_rows(df, column, operation, value):
        """Filter rows with improved date handling and error checking"""
        try:
            mask = MultiSheetProcessor.get_filter_mask(df, column, operation, value)
            if mask is None:
                return df
            return df[mask]
        except Exception as e:
            st.error(f"Error filtering: {e}")
            return df

    @staticmethod
    def filter_multiple_rows(df, conditions, logic='AND'):
        """Apply multiple filter conditions, combined with AND or OR logic.

        Each condition's mask is computed independently against the ORIGINAL
        (unfiltered) data and the masks are combined afterward. This makes the
        result independent of condition order, and - critically - makes OR
        combinations possible: e.g. "Status == Pending" OR "Status == Review"
        can never both be true on one row under sequential AND-narrowing
        (each condition would filter out the other's matches), so OR support
        requires evaluating conditions against the same base data.
        """
        try:
            masks = []
            for condition in conditions:
                column = condition.get('column')
                operation = condition.get('operation')
                value = condition.get('value')
                if column and operation and value is not None:
                    mask = MultiSheetProcessor.get_filter_mask(df, column, operation, value)
                    if mask is not None:
                        masks.append(mask)

            if not masks:
                return df

            combined = masks[0]
            for m in masks[1:]:
                combined = (combined | m) if logic == 'OR' else (combined & m)

            return df[combined]
        except Exception as e:
            st.error(f"Error applying multiple filters: {e}")
            return df

    @staticmethod
    def add_calculated_column(df, new_column, formula, all_sources=None):
        try:
            result = MultiSheetProcessor.evaluate_formula(df, formula, all_sources)
            df = df.copy()
            df[new_column] = result
            return df
        except Exception as e:
            st.error(f"Error adding column: {e}")
            return df

    @staticmethod
    def evaluate_formula(df, formula, all_sources=None):
        """Evaluate a formula with [Column] references. Supports:
        - Arithmetic / string ops on columns
        - abs(), round(), upper(), lower(), len()
        - IF <cond> THEN <expr> ELSE <expr> END  (vectorized)
        - XLOOKUP(lookup_value, 'RefSourceName', 'RefKeyColumn', 'ReturnColumn', if_not_found)
          looks up values from any other loaded data source, like Excel's XLOOKUP.
        """
        try:
            namespace = dict(SAFE_NAMESPACE_FUNCS)
            namespace["df"] = df
            namespace["XLOOKUP"] = lambda lookup_value, ref_name, lookup_col, return_col, if_not_found=None: \
                _xlookup(lookup_value, ref_name, lookup_col, return_col, if_not_found, sources=all_sources)
            namespace["VLOOKUP"] = namespace["XLOOKUP"]  # familiar alias, same behavior

            if_match = re.search(
                r'IF\s+(.+?)\s+THEN\s+(.+?)\s+ELSE\s+(.+?)\s+END',
                formula, re.IGNORECASE | re.DOTALL
            )
            if if_match:
                condition = _replace_columns(if_match.group(1).strip(), df)
                true_val = _replace_columns(if_match.group(2).strip(), df)
                false_val = _replace_columns(if_match.group(3).strip(), df)

                cond_result = eval(condition, {"__builtins__": {}}, namespace)
                true_result = eval(true_val, {"__builtins__": {}}, namespace)
                false_result = eval(false_val, {"__builtins__": {}}, namespace)

                return pd.Series(
                    np.where(cond_result, true_result, false_result),
                    index=df.index
                )

            expr = _replace_columns(formula, df)
            result = eval(expr, {"__builtins__": {}}, namespace)
            return result
        except Exception as e:
            raise Exception(f"Error evaluating formula: {e}")

    @staticmethod
    def group_by(df, group_column, agg_column, agg_function):
        try:
            return df.groupby(group_column)[agg_column].agg(agg_function).reset_index()
        except Exception as e:
            st.error(f"Error grouping: {e}")
            return df

    @staticmethod
    def sort_data(df, column, ascending=True):
        try:
            return df.sort_values(by=column, ascending=ascending)
        except Exception as e:
            st.error(f"Error sorting: {e}")
            return df

    @staticmethod
    def rename_column(df, old_name, new_name):
        try:
            return df.rename(columns={old_name: new_name})
        except Exception as e:
            st.error(f"Error renaming: {e}")
            return df

    @staticmethod
    def drop_column(df, column):
        try:
            return df.drop(columns=[column])
        except Exception as e:
            st.error(f"Error dropping column: {e}")
            return df

    @staticmethod
    def pivot_table(df, index, columns, values, aggfunc):
        try:
            result = pd.pivot_table(df, values=values, index=index,
                                     columns=columns, aggfunc=aggfunc)
            return result.reset_index()
        except Exception as e:
            st.error(f"Error creating pivot: {e}")
            return df

    @staticmethod
    def create_relationship(df1, df2, on, how='inner'):
        try:
            return pd.merge(df1, df2, on=on, how=how)
        except Exception as e:
            st.error(f"Error creating relationship: {e}")
            return None

    @staticmethod
    def vlookup(df, lookup_col, lookup_df, lookup_on, return_col):
        try:
            mapping = dict(zip(lookup_df[lookup_on], lookup_df[return_col]))
            return df[lookup_col].map(mapping)
        except Exception as e:
            st.error(f"Error in VLOOKUP: {e}")
            return None

    @staticmethod
    def filter_by_reference(df, column, reference_df, ref_column, operation='in'):
        try:
            ref_values = reference_df[ref_column].unique()
            if operation == 'in':
                return df[df[column].isin(ref_values)]
            elif operation == 'not in':
                return df[~df[column].isin(ref_values)]
            return df
        except Exception as e:
            st.error(f"Error filtering by reference: {e}")
            return df

    @staticmethod
    def append_data(df1, df2):
        try:
            return pd.concat([df1, df2], ignore_index=True)
        except Exception as e:
            st.error(f"Error appending data: {e}")
            return df1

    @staticmethod
    def summarize_by_reference(df, group_col, agg_col, ref_df, ref_col, agg_func='sum'):
        try:
            agg = df.groupby(group_col)[agg_col].agg(agg_func).reset_index()
            return pd.merge(agg, ref_df, left_on=group_col, right_on=ref_col, how='left')
        except Exception as e:
            st.error(f"Error summarizing by reference: {e}")
            return df

    @staticmethod
    def build_template_output(df, template_headers, column_configs, all_sources=None):
        """Build a new table whose columns exactly match template_headers, where each
        column is populated according to its own rule: direct copy, XLOOKUP-style
        reference lookup, formula, or a static value.
        """
        if len(df) == 0:
            st.warning("⚠️ Building template output, but the input data already has 0 rows "
                       "(an earlier step - likely a filter - removed everything). "
                       "The output will also have 0 rows.")
        result = pd.DataFrame(index=df.index)
        for header in template_headers:
            cfg = column_configs.get(header, {'mode': 'static', 'value': ''})
            mode = cfg.get('mode')
            try:
                if mode == 'direct':
                    col = cfg.get('source_column')
                    result[header] = df[col] if col in df.columns else np.nan

                elif mode == 'lookup':
                    ref_name = cfg.get('ref_source')
                    ref_df = (all_sources or {}).get(ref_name)
                    if ref_df is None:
                        st.error(f"Reference source '{ref_name}' not found for column '{header}'")
                        result[header] = np.nan
                        continue
                    key_col = cfg.get('key_column')
                    ref_key_col = cfg.get('ref_key_column')
                    ref_return_col = cfg.get('ref_return_column')
                    mapping = dict(zip(ref_df[ref_key_col], ref_df[ref_return_col]))
                    mapped = df[key_col].map(mapping)
                    default = cfg.get('default')
                    if default not in (None, ''):
                        mapped = mapped.fillna(default)
                    result[header] = mapped

                elif mode == 'formula':
                    result[header] = MultiSheetProcessor.evaluate_formula(df, cfg.get('formula', ''), all_sources)

                elif mode == 'static':
                    result[header] = cfg.get('value', '')

                else:
                    result[header] = np.nan
            except Exception as e:
                st.error(f"Error building column '{header}': {e}")
                result[header] = np.nan

        return result.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Operation execution (used both for "Execute All" and Automation replay)
# ---------------------------------------------------------------------------
SOURCE_REF_KEYS = ['source', 'lookup_source', 'ref_source', 'append_source']


def _warn_if_wiped(before_len, result, op_desc):
    """Surface a clear warning the moment a step zeroes out rows that
    previously had data, instead of letting it fail silently and only
    showing up as an empty export several steps later."""
    if before_len > 0 and result is not None and len(result) == 0:
        st.warning(f"⚠️ Step wiped all rows to 0: {op_desc}. Check the column name, "
                   f"value, and data type used in this step.")
    return result


def apply_operation(df, operation, all_sources=None):
    op_type = operation['type']
    before_len = len(df)

    if op_type == 'filter':
        result = MultiSheetProcessor.filter_rows(df, operation['column'], operation['operation'], operation['value'])
        return _warn_if_wiped(before_len, result, describe_operation(operation))

    elif op_type == 'filter_multiple':
        result = MultiSheetProcessor.filter_multiple_rows(df, operation['conditions'], operation.get('logic', 'AND'))
        return _warn_if_wiped(before_len, result, describe_operation(operation))

    elif op_type == 'calc_column':
        return MultiSheetProcessor.add_calculated_column(df, operation['new_column'], operation['formula'], all_sources)

    elif op_type == 'build_template':
        return MultiSheetProcessor.build_template_output(
            df, operation['template_headers'], operation['column_configs'], all_sources
        )

    elif op_type == 'group_by':
        return MultiSheetProcessor.group_by(df, operation['group_column'], operation['aggregate_column'], operation['function'])

    elif op_type == 'sort':
        return MultiSheetProcessor.sort_data(df, operation['column'], operation['ascending'])

    elif op_type == 'rename':
        return MultiSheetProcessor.rename_column(df, operation['old_name'], operation['new_name'])

    elif op_type == 'drop':
        return MultiSheetProcessor.drop_column(df, operation['column'])

    elif op_type == 'merge':
        if all_sources and operation['source'] in all_sources:
            return MultiSheetProcessor.create_relationship(
                df, all_sources[operation['source']], operation['on'], operation['how']
            )
        return df

    elif op_type == 'vlookup':
        if all_sources and operation['lookup_source'] in all_sources:
            result = MultiSheetProcessor.vlookup(
                df, operation['lookup_col'], all_sources[operation['lookup_source']],
                operation['lookup_on'], operation['return_col']
            )
            if result is not None:
                df = df.copy()
                df[operation['new_column']] = result
            return df
        return df

    elif op_type == 'filter_by_ref':
        if all_sources and operation['ref_source'] in all_sources:
            return MultiSheetProcessor.filter_by_reference(
                df, operation['column'], all_sources[operation['ref_source']],
                operation['ref_column'], operation.get('operation', 'in')
            )
        return df

    elif op_type == 'append':
        if all_sources and operation['append_source'] in all_sources:
            return MultiSheetProcessor.append_data(df, all_sources[operation['append_source']])
        return df

    elif op_type == 'summarize_by_ref':
        if all_sources and operation['ref_source'] in all_sources:
            return MultiSheetProcessor.summarize_by_reference(
                df, operation['group_col'], operation['agg_col'],
                all_sources[operation['ref_source']], operation['ref_col'],
                operation.get('agg_func', 'sum')
            )
        return df

    else:
        return df


def execute_all_operations(df, operations, all_sources):
    for op in operations:
        df = apply_operation(df, op, all_sources)
    return df


def extract_referenced_sources(recipe):
    """All source names (primary + any referenced in steps) a recipe needs."""
    refs = set()
    if recipe.get('primary_source'):
        refs.add(recipe['primary_source'])
    for op in recipe.get('operations', []):
        for key in SOURCE_REF_KEYS:
            if key in op and op[key]:
                refs.add(op[key])
        if op['type'] == 'calc_column' and op.get('referenced_sources'):
            refs.update(op['referenced_sources'])
        if op['type'] == 'build_template':
            for cfg in op.get('column_configs', {}).values():
                if cfg.get('mode') == 'lookup' and cfg.get('ref_source'):
                    refs.add(cfg['ref_source'])
                if cfg.get('mode') == 'formula' and cfg.get('referenced_sources'):
                    refs.update(cfg['referenced_sources'])
    return sorted(refs)


def _remap_formula_text(formula, referenced_sources, mapping):
    """Swap quoted old source names for new ones inside a formula string."""
    for old_name in referenced_sources:
        new_name = mapping.get(old_name)
        if new_name:
            formula = formula.replace(f"'{old_name}'", f"'{new_name}'").replace(f'"{old_name}"', f'"{new_name}"')
    return formula


def remap_operations(operations, mapping):
    """Swap old source names in a recipe's steps for the newly chosen ones."""
    new_ops = []
    for op in operations:
        new_op = dict(op)

        for key in SOURCE_REF_KEYS:
            if key in new_op and new_op[key] in mapping:
                new_op[key] = mapping[new_op[key]]

        if new_op['type'] == 'calc_column' and new_op.get('referenced_sources'):
            new_op['formula'] = _remap_formula_text(new_op['formula'], new_op['referenced_sources'], mapping)
            new_op['referenced_sources'] = [mapping.get(s, s) for s in new_op['referenced_sources']]

        if new_op['type'] == 'build_template':
            new_configs = {}
            for col, cfg in new_op.get('column_configs', {}).items():
                cfg = dict(cfg)
                if cfg.get('mode') == 'lookup' and cfg.get('ref_source') in mapping:
                    cfg['ref_source'] = mapping[cfg['ref_source']]
                if cfg.get('mode') == 'formula' and cfg.get('referenced_sources'):
                    cfg['formula'] = _remap_formula_text(cfg['formula'], cfg['referenced_sources'], mapping)
                    cfg['referenced_sources'] = [mapping.get(s, s) for s in cfg['referenced_sources']]
                new_configs[col] = cfg
            new_op['column_configs'] = new_configs

        new_ops.append(new_op)
    return new_ops


def build_recipe():
    return {
        'primary_source': st.session_state.active_source,
        'operations': st.session_state.operations,
        'timestamp': datetime.now().isoformat(),
    }


# ---------------------------------------------------------------------------
# Editable / reorderable recipe steps + undo
#
# The active source's data is always recomputed from its baseline snapshot by
# replaying st.session_state.operations in order. This makes reordering,
# editing, and deleting steps (and undoing changes to the step list) safe and
# consistent, instead of mutating the dataframe in place one click at a time.
# ---------------------------------------------------------------------------
def recompute_active_source():
    """Rebuild the active source's dataframe from its baseline snapshot by
    replaying the current operations list, in order."""
    name = st.session_state.active_source
    if not name:
        return
    snapshot = st.session_state.source_snapshots.get(name)
    if snapshot is None:
        return  # no baseline to replay from (shouldn't normally happen)
    result = snapshot.copy()
    for op in st.session_state.operations:
        result = apply_operation(result, op, st.session_state.data_sources)
    st.session_state.data_sources[name] = result


def push_undo_snapshot():
    """Call before any change to the operations list, so it can be undone."""
    st.session_state.operations_undo_stack.append(list(st.session_state.operations))
    if len(st.session_state.operations_undo_stack) > 25:
        st.session_state.operations_undo_stack.pop(0)


def undo_last_operation():
    if st.session_state.operations_undo_stack:
        st.session_state.operations = st.session_state.operations_undo_stack.pop()
        recompute_active_source()


def describe_operation(op):
    """Short, human-readable summary of a recipe step, for the editor list."""
    t = op.get('type')
    try:
        if t == 'filter':
            return f"Filter: {op['column']} {op['operation']} {op['value']}"
        if t == 'filter_multiple':
            cond_count = len(op.get('conditions', []))
            logic = op.get('logic', 'AND')
            return f"Multiple Filters ({logic}): {cond_count} condition(s)"
        if t == 'calc_column':
            return f"Add column '{op['new_column']}' = {op['formula']}"
        if t == 'group_by':
            return f"Group by {op['group_column']}, {op['function']}({op['aggregate_column']})"
        if t == 'sort':
            return f"Sort by {op['column']} ({'asc' if op['ascending'] else 'desc'})"
        if t == 'rename':
            return f"Rename '{op['old_name']}' → '{op['new_name']}'"
        if t == 'drop':
            return f"Drop column '{op['column']}'"
        if t == 'merge':
            return f"Merge with '{op['source']}' on {op['on']} ({op['how']})"
        if t == 'vlookup':
            return f"VLOOKUP '{op['lookup_col']}' from '{op['lookup_source']}' → '{op['new_column']}'"
        if t == 'filter_by_ref':
            return f"Filter '{op['column']}' {op['operation']} ref '{op['ref_source']}'.{op['ref_column']}"
        if t == 'append':
            return f"Append rows from '{op['append_source']}'"
        if t == 'summarize_by_ref':
            return f"Summarize {op.get('agg_func', 'sum')}({op['agg_col']}) by {op['group_col']}, joined with '{op['ref_source']}'"
        if t == 'build_template':
            return f"Build output from template ({len(op.get('template_headers', []))} column(s))"
        if t == 'pivot':
            return f"Pivot: index={op['index']}, columns={op['columns']}, values={op['values']} ({op['aggfunc']})"
    except Exception:
        pass
    return t


def render_operation_edit_form(op, idx, baseline_df, all_sources):
    """Render inputs pre-filled with an existing step's values and return an
    updated operation dict. baseline_df is used to populate dropdowns; it may
    not exactly match this step's actual input columns if earlier steps
    changed them, so a free-text fallback is offered where relevant.
    """
    t = op.get('type')
    new_op = dict(op)
    cols = list(baseline_df.columns) if baseline_df is not None else []
    src_names = list(all_sources.keys())

    if t == 'filter':
        c1, c2 = st.columns(2)
        with c1:
            new_op['column'] = st.text_input("Column", op['column'], key=f"edit_col_{idx}")
        with c2:
            ops_list = ['==', '!=', '>', '<', '>=', '<=', 'contains']
            new_op['operation'] = st.selectbox("Operation", ops_list,
                                                index=ops_list.index(op['operation']) if op['operation'] in ops_list else 0,
                                                key=f"edit_op_{idx}")
        new_op['value'] = st.text_input("Value", op['value'], key=f"edit_val_{idx}")

    elif t == 'filter_multiple':
        st.info("Edit each condition individually")
        logic_opts = ['AND', 'OR']
        cur_logic = op.get('logic', 'AND')
        new_op['logic'] = st.radio(
            "Combine conditions using", logic_opts,
            index=logic_opts.index(cur_logic) if cur_logic in logic_opts else 0,
            horizontal=True, key=f"edit_mf_logic_{idx}",
            help="AND: rows must match every condition. OR: rows must match at least one condition."
        )
        new_conditions = []
        for i, cond in enumerate(op.get('conditions', [])):
            st.write(f"Condition {i+1}:")
            col1, col2, col3 = st.columns(3)
            with col1:
                cond_col = st.text_input("Column", cond.get('column', ''), key=f"edit_mf_col_{idx}_{i}")
            with col2:
                ops_list = ['==', '!=', '>', '<', '>=', '<=', 'contains']
                cond_op = st.selectbox("Operation", ops_list,
                                       index=ops_list.index(cond.get('operation', '==')) if cond.get('operation') in ops_list else 0,
                                       key=f"edit_mf_op_{idx}_{i}")
            with col3:
                cond_val = st.text_input("Value", cond.get('value', ''), key=f"edit_mf_val_{idx}_{i}")
            new_conditions.append({'column': cond_col, 'operation': cond_op, 'value': cond_val})
        new_op['conditions'] = new_conditions

    elif t == 'calc_column':
        new_op['new_column'] = st.text_input("New column name", op['new_column'], key=f"edit_newcol_{idx}")
        new_op['formula'] = st.text_area("Formula", op['formula'], key=f"edit_formula_{idx}")
        new_op['referenced_sources'] = detect_referenced_source_names(new_op['formula'], all_sources.keys())

    elif t == 'sort':
        new_op['column'] = st.text_input("Sort column", op['column'], key=f"edit_sortcol_{idx}")
        order = st.radio("Order", ["Ascending", "Descending"], index=0 if op['ascending'] else 1,
                          horizontal=True, key=f"edit_sortorder_{idx}")
        new_op['ascending'] = order == "Ascending"

    elif t == 'rename':
        new_op['old_name'] = st.text_input("Old name", op['old_name'], key=f"edit_oldname_{idx}")
        new_op['new_name'] = st.text_input("New name", op['new_name'], key=f"edit_newname_{idx}")

    elif t == 'drop':
        new_op['column'] = st.text_input("Column to drop", op['column'], key=f"edit_dropcol_{idx}")

    elif t == 'group_by':
        new_op['group_column'] = st.text_input("Group column", op['group_column'], key=f"edit_groupcol_{idx}")
        new_op['aggregate_column'] = st.text_input("Aggregate column", op['aggregate_column'], key=f"edit_aggcol_{idx}")
        funcs = ['sum', 'mean', 'median', 'min', 'max', 'count', 'std']
        new_op['function'] = st.selectbox("Function", funcs,
                                           index=funcs.index(op['function']) if op['function'] in funcs else 0,
                                           key=f"edit_aggfunc_{idx}")

    elif t == 'vlookup':
        new_op['lookup_source'] = st.selectbox("Reference source", src_names,
                                                index=src_names.index(op['lookup_source']) if op['lookup_source'] in src_names else 0,
                                                key=f"edit_vlsrc_{idx}")
        new_op['lookup_col'] = st.text_input("Lookup column (current)", op['lookup_col'], key=f"edit_vlcol_{idx}")
        new_op['lookup_on'] = st.text_input("Lookup column (reference)", op['lookup_on'], key=f"edit_vlon_{idx}")
        new_op['return_col'] = st.text_input("Return column", op['return_col'], key=f"edit_vlret_{idx}")
        new_op['new_column'] = st.text_input("New column name", op['new_column'], key=f"edit_vlnew_{idx}")

    elif t == 'filter_by_ref':
        new_op['ref_source'] = st.selectbox("Reference source", src_names,
                                             index=src_names.index(op['ref_source']) if op['ref_source'] in src_names else 0,
                                             key=f"edit_fbrsrc_{idx}")
        new_op['column'] = st.text_input("Column to filter", op['column'], key=f"edit_fbrcol_{idx}")
        new_op['ref_column'] = st.text_input("Reference column", op['ref_column'], key=f"edit_fbrrefcol_{idx}")
        new_op['operation'] = st.selectbox("Operation", ['in', 'not in'],
                                            index=0 if op['operation'] == 'in' else 1, key=f"edit_fbrop_{idx}")

    elif t == 'append':
        new_op['append_source'] = st.selectbox("Source to append", src_names,
                                                index=src_names.index(op['append_source']) if op['append_source'] in src_names else 0,
                                                key=f"edit_appendsrc_{idx}")

    elif t == 'summarize_by_ref':
        new_op['ref_source'] = st.selectbox("Reference source", src_names,
                                             index=src_names.index(op['ref_source']) if op['ref_source'] in src_names else 0,
                                             key=f"edit_sumref_{idx}")
        new_op['group_col'] = st.text_input("Group column", op['group_col'], key=f"edit_sumgroup_{idx}")
        new_op['agg_col'] = st.text_input("Aggregate column", op['agg_col'], key=f"edit_sumagg_{idx}")
        new_op['ref_col'] = st.text_input("Reference join column", op['ref_col'], key=f"edit_sumrefcol_{idx}")
        funcs = ['sum', 'mean', 'count', 'min', 'max']
        cur_func = op.get('agg_func', 'sum')
        new_op['agg_func'] = st.selectbox("Function", funcs,
                                           index=funcs.index(cur_func) if cur_func in funcs else 0,
                                           key=f"edit_sumfunc_{idx}")

    elif t == 'pivot':
        new_op['index'] = st.text_input("Index column", op['index'], key=f"edit_pvindex_{idx}")
        new_op['columns'] = st.text_input("Columns", op['columns'], key=f"edit_pvcols_{idx}")
        new_op['values'] = st.text_input("Values", op['values'], key=f"edit_pvvals_{idx}")
        funcs = ['sum', 'mean', 'count', 'min', 'max']
        new_op['aggfunc'] = st.selectbox("Aggregation", funcs,
                                          index=funcs.index(op['aggfunc']) if op['aggfunc'] in funcs else 0,
                                          key=f"edit_pvagg_{idx}")

    elif t == 'merge':
        new_op['source'] = st.selectbox("Merge with", src_names,
                                         index=src_names.index(op['source']) if op['source'] in src_names else 0,
                                         key=f"edit_mergesrc_{idx}")
        new_op['on'] = st.text_input("Join column", op['on'], key=f"edit_mergeon_{idx}")
        how_opts = ['inner', 'left', 'right', 'outer']
        new_op['how'] = st.selectbox("Join type", how_opts,
                                      index=how_opts.index(op['how']) if op['how'] in how_opts else 0,
                                      key=f"edit_mergehow_{idx}")

    else:
        st.info("This step type doesn't support inline field editing yet — you can still reorder or delete it.")

    return new_op


def to_excel_bytes(dfs_by_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, data in dfs_by_name.items():
            data.to_excel(writer, index=False, sheet_name=name[:31])
    return output.getvalue()


# ========================= MAIN APP =========================

st.markdown('<h1 class="main-header">📚 Multi-Sheet Data Processing Wizard</h1>', unsafe_allow_html=True)
st.markdown("### Build your process once, save it, and reuse it on every new report")

# ---------------------------------------------------------------------------
# Sidebar - Data Source Management
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("📁 Data Sources")

    uploaded_files = st.file_uploader(
        "Upload your files (CSV or Excel)",
        type=['csv', 'xlsx', 'xls'],
        accept_multiple_files=True,
        help="Upload multiple files at once. Excel files will load all sheets.",
        key="main_uploader"
    )

    if uploaded_files:
        if st.button("🔄 Load All Files", use_container_width=True, key="load_all_files_btn"):
            for file in uploaded_files:
                file_name = Path(file.name).stem
                sheets = MultiSheetProcessor.load_single_sheet(file)
                if sheets:
                    for sheet_name, df in sheets.items():
                        source_name = f"{file_name}_{sheet_name}" if len(sheets) > 1 else file_name
                        st.session_state.data_sources[source_name] = df
                        st.session_state.source_snapshots[source_name] = df.copy()
                        st.session_state.source_metadata[source_name] = {
                            'file': file.name,
                            'sheet': sheet_name,
                            'rows': len(df),
                            'columns': len(df.columns)
                        }
            st.success(f"✅ Loaded {len(st.session_state.data_sources)} data sources")
            st.rerun()

    if st.session_state.data_sources:
        st.divider()
        st.header("📊 Loaded Data Sources")

        for source_name, df in list(st.session_state.data_sources.items()):
            col1, col2 = st.columns([3, 1])
            with col1:
                is_active = st.session_state.active_source == source_name
                btn_label = f"📊 {source_name}" + (" ✅" if is_active else "")
                if st.button(btn_label, key=f"activate_{source_name}", use_container_width=True):
                    if st.session_state.active_source != source_name:
                        # Switching sources starts a fresh recipe for that source.
                        st.session_state.operations = []
                        st.session_state.operations_undo_stack = []
                        st.session_state.filter_conditions = []
                    st.session_state.active_source = source_name
                    st.session_state.current_operation = None
                    if source_name not in st.session_state.source_snapshots:
                        st.session_state.source_snapshots[source_name] = st.session_state.data_sources[source_name].copy()
                    st.rerun()
                meta = st.session_state.source_metadata.get(source_name, {})
                st.caption(f"Rows: {meta.get('rows', 0)} | Cols: {meta.get('columns', 0)}")
            with col2:
                if st.button("❌", key=f"remove_{source_name}", help="Remove this source"):
                    del st.session_state.data_sources[source_name]
                    st.session_state.source_metadata.pop(source_name, None)
                    if st.session_state.active_source == source_name:
                        st.session_state.active_source = None
                    st.rerun()

        st.info(f"📊 Total: {len(st.session_state.data_sources)} data sources loaded")

    # Relationship Management
    if len(st.session_state.data_sources) >= 2:
        st.divider()
        st.header("🔗 Relationships")

        with st.expander("Define Relationships"):
            col1, col2 = st.columns(2)
            with col1:
                source1 = st.selectbox("Source 1", list(st.session_state.data_sources.keys()), key="rel_source1")
            with col2:
                source2 = st.selectbox("Source 2", list(st.session_state.data_sources.keys()), key="rel_source2")

            if source1 != source2:
                df1 = st.session_state.data_sources[source1]
                df2 = st.session_state.data_sources[source2]
                common_cols = [c for c in df1.columns if c in df2.columns]

                if common_cols:
                    on_col = st.selectbox("Join on column", common_cols, key="rel_on_col")
                    how = st.selectbox("Join type", ['inner', 'left', 'right', 'outer'], key="rel_how")

                    if st.button("✅ Create Relationship", use_container_width=True, key="create_rel_btn"):
                        result = MultiSheetProcessor.create_relationship(df1, df2, on_col, how)
                        if result is not None:
                            rel_name = f"{source1}_{source2}_joined"
                            st.session_state.data_sources[rel_name] = result
                            st.session_state.source_metadata[rel_name] = {
                                'file': 'Relationship',
                                'sheet': 'Joined',
                                'rows': len(result),
                                'columns': len(result.columns),
                                'relationship': f'{source1} ↔ {source2} on {on_col}'
                            }
                            st.session_state.relationships.append({
                                'source1': source1, 'source2': source2, 'on': on_col,
                                'how': how, 'result': rel_name
                            })
                            st.success(f"✅ Created relationship: {rel_name}")
                            st.caption(
                                "Note: to include this join in a saved recipe, make the joined "
                                f"source ('{rel_name}') your active source and continue building from there."
                            )
                            st.rerun()
                else:
                    st.warning("No common columns found between these sources")

    # ---------------------- Recipe (save the process you built) ----------------------
    if st.session_state.operations:
        st.divider()
        st.header("📝 Recipe")
        st.caption("Reorder, edit, or delete any step. Switching to a different source starts a new recipe.")

        if st.button("↩️ Undo Last Change", use_container_width=True, key="undo_btn",
                      disabled=(len(st.session_state.operations_undo_stack) == 0)):
            undo_last_operation()
            st.rerun()

        ops = st.session_state.operations
        baseline_df = st.session_state.source_snapshots.get(st.session_state.active_source)

        for i, op in enumerate(ops):
            s1, s2, s3, s4, s5 = st.columns([5, 1, 1, 1, 1])
            with s1:
                st.write(f"**{i+1}.** `{op['type']}` — {describe_operation(op)}")
            with s2:
                if st.button("⬆️", key=f"up_{i}", disabled=(i == 0), help="Move step up"):
                    push_undo_snapshot()
                    ops[i-1], ops[i] = ops[i], ops[i-1]
                    recompute_active_source()
                    st.rerun()
            with s3:
                if st.button("⬇️", key=f"down_{i}", disabled=(i == len(ops) - 1), help="Move step down"):
                    push_undo_snapshot()
                    ops[i+1], ops[i] = ops[i], ops[i+1]
                    recompute_active_source()
                    st.rerun()
            with s4:
                if st.button("✏️", key=f"edit_toggle_{i}", help="Edit this step"):
                    st.session_state[f"editing_step_{i}"] = not st.session_state.get(f"editing_step_{i}", False)
                    st.rerun()
            with s5:
                if st.button("🗑️", key=f"delete_{i}", help="Delete this step"):
                    push_undo_snapshot()
                    ops.pop(i)
                    recompute_active_source()
                    st.rerun()

            if st.session_state.get(f"editing_step_{i}", False):
                with st.expander(f"Editing step {i+1}: {op['type']}", expanded=True):
                    edited_op = render_operation_edit_form(op, i, baseline_df, st.session_state.data_sources)
                    if st.button("✅ Save Changes", key=f"save_edit_{i}", use_container_width=True):
                        push_undo_snapshot()
                        ops[i] = edited_op
                        recompute_active_source()
                        st.session_state[f"editing_step_{i}"] = False
                        st.success(f"Step {i+1} updated")
                        st.rerun()

        recipe = build_recipe()
        recipe_json = json.dumps(recipe, indent=2)

        st.download_button(
            "💾 Download Recipe (.json)",
            data=recipe_json,
            file_name=f"recipe_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json",
            use_container_width=True,
            key="download_recipe_btn"
        )

        with st.expander("View recipe JSON"):
            st.code(recipe_json, language='json')

        if st.button("🔄 Reset All", use_container_width=True, key="reset_all_btn"):
            st.session_state.data_sources = {}
            st.session_state.operations = []
            st.session_state.active_source = None
            st.session_state.relationships = []
            st.session_state.current_operation = None
            st.session_state.loaded_recipe = None
            st.session_state.recipe_mapping = {}
            st.session_state.source_snapshots = {}
            st.session_state.operations_undo_stack = []
            st.session_state.batch_results = {}
            st.session_state.filter_conditions = []
            st.success("✅ Reset complete!")
            st.rerun()

    # ---------------------- Automation (replay a saved recipe) ----------------------
    st.divider()
    st.header("🤖 Automation")
    st.caption("Upload a recipe you saved earlier and run it against new data — no rebuilding required.")

    recipe_file = st.file_uploader("Upload a recipe (.json)", type=['json'], key="recipe_uploader")
    if recipe_file is not None:
        try:
            st.session_state.loaded_recipe = json.load(recipe_file)
            st.session_state.recipe_mapping = {}
        except Exception as e:
            st.error(f"Couldn't read recipe file: {e}")
            st.session_state.loaded_recipe = None

    if st.session_state.loaded_recipe and st.session_state.data_sources:
        recipe = st.session_state.loaded_recipe
        needed_sources = extract_referenced_sources(recipe)
        st.write(f"This recipe needs **{len(needed_sources)}** source(s) and has "
                 f"**{len(recipe.get('operations', []))}** step(s).")

        loaded_names = list(st.session_state.data_sources.keys())
        for needed in needed_sources:
            default_index = loaded_names.index(needed) if needed in loaded_names else 0
            chosen = st.selectbox(
                f"Map '{needed}' to:", loaded_names,
                index=default_index, key=f"map_{needed}"
            )
            st.session_state.recipe_mapping[needed] = chosen

        if st.button("▶️ Run Automation", use_container_width=True, key="run_automation_btn"):
            mapping = st.session_state.recipe_mapping
            primary_new = mapping.get(recipe.get('primary_source'))
            if not primary_new:
                st.error("Please map the primary source before running.")
            else:
                remapped_ops = remap_operations(recipe.get('operations', []), mapping)
                primary_df = st.session_state.data_sources[primary_new]
                result_df = execute_all_operations(primary_df, remapped_ops, st.session_state.data_sources)
                result_name = f"{primary_new}_automated_{datetime.now().strftime('%H%M%S')}"
                st.session_state.data_sources[result_name] = result_df
                st.session_state.source_metadata[result_name] = {
                    'file': 'Automation', 'sheet': 'Result',
                    'rows': len(result_df), 'columns': len(result_df.columns)
                }
                st.session_state.source_snapshots[result_name] = primary_df.copy()
                st.session_state.active_source = result_name
                st.session_state.operations = list(remapped_ops)
                st.session_state.operations_undo_stack = []
                st.success(f"✅ Automation complete! New source: {result_name}")
                st.rerun()
    elif st.session_state.loaded_recipe and not st.session_state.data_sources:
        st.info("Upload the data file(s) this recipe should run on first.")

    # ---------------------- Batch Automation (one recipe, many files) ----------------------
    st.divider()
    st.header("📦 Batch Automation")
    st.caption("Run a saved recipe against many files in one go — each uploaded file becomes one output.")

    batch_recipe_file = st.file_uploader("Upload a recipe (.json)", type=['json'], key="batch_recipe_uploader")
    batch_files = st.file_uploader(
        "Upload files to process — one run of the recipe per file",
        type=['csv', 'xlsx', 'xls'], accept_multiple_files=True, key="batch_files_uploader"
    )

    if batch_recipe_file is not None and batch_files:
        try:
            batch_recipe = json.load(batch_recipe_file)
        except Exception as e:
            st.error(f"Couldn't read recipe file: {e}")
            batch_recipe = None

        if batch_recipe:
            primary_name = batch_recipe.get('primary_source')
            needed_sources = extract_referenced_sources(batch_recipe)
            other_needed = [s for s in needed_sources if s != primary_name]

            st.write(f"Each uploaded file will be used as **'{primary_name}'**. "
                     f"This recipe has **{len(batch_recipe.get('operations', []))}** step(s).")

            ref_mapping = {}
            loaded_names = list(st.session_state.data_sources.keys())
            if other_needed:
                st.caption("This recipe also needs reference source(s) — map them to data you already have loaded:")
                if not loaded_names:
                    st.warning("Load the needed reference file(s) first (via the uploader at the top of the sidebar).")
                for needed in other_needed:
                    default_index = loaded_names.index(needed) if needed in loaded_names else 0
                    if loaded_names:
                        ref_mapping[needed] = st.selectbox(
                            f"Map '{needed}' to:", loaded_names, index=default_index, key=f"batch_map_{needed}"
                        )

            if st.button("▶️ Run Batch", use_container_width=True, key="run_batch_btn"):
                results, errors = {}, {}
                remapped_ops = remap_operations(batch_recipe.get('operations', []), ref_mapping)
                for f in batch_files:
                    try:
                        file_stem = Path(f.name).stem
                        sheets = MultiSheetProcessor.load_single_sheet(f)
                        if not sheets:
                            errors[f.name] = "Could not read file"
                            continue
                        # Uses the first sheet as the primary source for each file.
                        primary_df = list(sheets.values())[0]
                        result_df = execute_all_operations(primary_df, remapped_ops, st.session_state.data_sources)
                        results[f"{file_stem}_result"] = result_df
                    except Exception as e:
                        errors[f.name] = str(e)

                st.session_state.batch_results = results
                if results:
                    st.success(f"✅ Batch complete: {len(results)} of {len(batch_files)} file(s) processed.")
                if errors:
                    st.error(f"⚠️ {len(errors)} file(s) failed:")
                    for fname, err in errors.items():
                        st.write(f"- **{fname}**: {err}")
                st.rerun()

    if st.session_state.batch_results:
        st.divider()
        st.subheader("📦 Batch Results")
        for name, rdf in st.session_state.batch_results.items():
            with st.expander(f"📊 {name} ({len(rdf)} rows, {len(rdf.columns)} cols)"):
                st.dataframe(rdf.head(20), use_container_width=True)

        combined_excel = to_excel_bytes(st.session_state.batch_results)
        st.download_button(
            "📥 Download All Batch Results (one Excel workbook)",
            data=combined_excel,
            file_name=f"batch_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="download_batch_all_btn"
        )

        if st.button("➕ Load batch results as data sources", use_container_width=True, key="load_batch_as_sources_btn"):
            for name, rdf in st.session_state.batch_results.items():
                st.session_state.data_sources[name] = rdf
                st.session_state.source_snapshots[name] = rdf.copy()
                st.session_state.source_metadata[name] = {
                    'file': 'Batch Automation', 'sheet': 'Result',
                    'rows': len(rdf), 'columns': len(rdf.columns)
                }
            st.success("Loaded batch results into your data sources.")
            st.rerun()


# ---------------------------------------------------------------------------
# Main content area - Active Data Source
# ---------------------------------------------------------------------------
if st.session_state.active_source:
    source_name = st.session_state.active_source
    df = st.session_state.data_sources[source_name]

    st.divider()
    meta = st.session_state.source_metadata.get(source_name, {})

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Active Source", source_name)
    with col2:
        st.metric("Rows", len(df))
    with col3:
        st.metric("Columns", len(df.columns))
    with col4:
        st.metric("Operations", len(st.session_state.operations))

    with st.expander("🔍 Data Preview", expanded=False):
        tab1, tab2, tab3 = st.tabs(["📊 Data", "📈 Statistics", "📋 Column Info"])
        with tab1:
            st.dataframe(df.head(100), use_container_width=True)
        with tab2:
            numeric_cols = df.select_dtypes(include=['number']).columns
            if len(numeric_cols) > 0:
                st.dataframe(df[numeric_cols].describe(), use_container_width=True)
            else:
                st.info("No numeric columns for statistics")
        with tab3:
            col_info = pd.DataFrame({
                'Column': df.columns,
                'Type': df.dtypes.astype(str),
                'Null Count': df.isnull().sum().values,
                'Unique Values': df.nunique().values
            })
            st.dataframe(col_info, use_container_width=True)

    # Reference Operations
    if len(st.session_state.data_sources) > 1:
        st.divider()
        st.header("🔗 Reference Operations")
        st.info("Use data from other sheets to enhance your current data")

        ref_col1, ref_col2 = st.columns(2)

        with ref_col1:
            with st.expander("🔍 VLOOKUP from Reference"):
                ref_source = st.selectbox(
                    "Reference source",
                    [s for s in st.session_state.data_sources.keys() if s != source_name],
                    key="vlookup_source"
                )
                if ref_source:
                    ref_df = st.session_state.data_sources[ref_source]
                    col1, col2 = st.columns(2)
                    with col1:
                        lookup_col = st.selectbox("Lookup column in current data", df.columns, key="vlookup_col")
                    with col2:
                        lookup_on = st.selectbox("Lookup column in reference", ref_df.columns, key="vlookup_on")
                    return_col = st.selectbox("Column to return from reference", ref_df.columns, key="vlookup_return")
                    new_col_name = st.text_input("New column name", f"{lookup_col}_from_{ref_source}", key="vlookup_newcol")

                    if st.button("✅ Apply VLOOKUP", use_container_width=True, key="apply_vlookup_btn"):
                        result = MultiSheetProcessor.vlookup(df, lookup_col, ref_df, lookup_on, return_col)
                        if result is not None:
                            push_undo_snapshot()
                            st.session_state.operations.append({
                                'type': 'vlookup', 'lookup_source': ref_source,
                                'lookup_col': lookup_col, 'lookup_on': lookup_on,
                                'return_col': return_col, 'new_column': new_col_name
                            })
                            recompute_active_source()
                            st.success(f"✅ Added column: {new_col_name}")
                            st.rerun()

        with ref_col2:
            with st.expander("🎯 Filter by Reference"):
                ref_source2 = st.selectbox(
                    "Reference source",
                    [s for s in st.session_state.data_sources.keys() if s != source_name],
                    key="filter_ref_source"
                )
                if ref_source2:
                    ref_df2 = st.session_state.data_sources[ref_source2]
                    col1, col2 = st.columns(2)
                    with col1:
                        filter_col = st.selectbox("Column to filter", df.columns, key="filter_col")
                    with col2:
                        ref_col = st.selectbox("Reference column", ref_df2.columns, key="filter_ref_col")
                    filter_op = st.selectbox("Operation", ['in', 'not in'], key="filter_ref_op")

                    if st.button("✅ Apply Filter", use_container_width=True, key="apply_filter_ref_btn"):
                        result = MultiSheetProcessor.filter_by_reference(df, filter_col, ref_df2, ref_col, filter_op)
                        if len(result) < len(df):
                            push_undo_snapshot()
                            st.session_state.operations.append({
                                'type': 'filter_by_ref', 'ref_source': ref_source2,
                                'column': filter_col, 'ref_column': ref_col, 'operation': filter_op
                            })
                            recompute_active_source()
                            st.success(f"✅ Filtered from {len(df)} to {len(result)} rows")
                            st.rerun()
                        else:
                            st.info("No rows filtered")

    # ---------------------------------------------------------------------
    # Template Output Builder — define your output headers, fill each one
    # via direct copy, an XLOOKUP-style reference lookup, a formula, or a
    # static value, then generate a table that matches your template exactly.
    # ---------------------------------------------------------------------
    st.divider()
    with st.expander("🧩 Build Output From Template", expanded=False):
        st.markdown(
            "Define the exact headers your output report needs, then tell each "
            "column where its data comes from — including looking values up from "
            "a reference file, just like `XLOOKUP` in Excel."
        )

        header_source = st.radio(
            "Where do the template headers come from?",
            ["Type them in", "Use headers from another loaded source", "Upload a template file"],
            horizontal=True, key="tpl_header_source"
        )

        if header_source == "Type them in":
            headers_text = st.text_input(
                "Comma-separated headers", ", ".join(st.session_state.template_headers),
                key="tpl_headers_text"
            )
            if st.button("Set Headers", key="tpl_set_headers_btn"):
                st.session_state.template_headers = [h.strip() for h in headers_text.split(",") if h.strip()]
                st.rerun()

        elif header_source == "Use headers from another loaded source":
            other_sources = [s for s in st.session_state.data_sources.keys() if s != source_name]
            if other_sources:
                pick_source = st.selectbox("Copy headers from", other_sources, key="tpl_copy_headers_source")
                if st.button("Set Headers", key="tpl_set_headers_from_source_btn"):
                    st.session_state.template_headers = list(st.session_state.data_sources[pick_source].columns)
                    st.rerun()
            else:
                st.info("Load another source first to copy its headers.")

        else:
            tpl_file = st.file_uploader(
                "Upload a file — only its header row is used", type=['csv', 'xlsx', 'xls'],
                key="tpl_file_uploader"
            )
            if tpl_file is not None:
                try:
                    if Path(tpl_file.name).suffix.lower() == '.csv':
                        header_cols = list(pd.read_csv(tpl_file, nrows=0).columns)
                    else:
                        header_cols = list(pd.read_excel(tpl_file, nrows=0).columns)
                    if st.button("Set Headers", key="tpl_set_headers_from_file_btn"):
                        st.session_state.template_headers = header_cols
                        st.rerun()
                except Exception as e:
                    st.error(f"Couldn't read headers from that file: {e}")

        if st.session_state.template_headers:
            st.success(f"Template has {len(st.session_state.template_headers)} column(s): "
                       f"{', '.join(st.session_state.template_headers)}")
            st.divider()
            st.markdown("**Configure each output column:**")

            other_sources = [s for s in st.session_state.data_sources.keys() if s != source_name]
            mode_options = ["Direct copy from current data", "Lookup from reference file (XLOOKUP)", "Formula", "Static value"]

            for header in st.session_state.template_headers:
                with st.expander(f"🔧 {header}"):
                    mode = st.selectbox(
                        "How should this column be filled?", mode_options,
                        key=f"tpl_mode_{header}"
                    )

                    if mode == mode_options[0]:
                        st.selectbox("Copy from column", df.columns, key=f"tpl_direct_{header}")

                    elif mode == mode_options[1]:
                        if not other_sources:
                            st.warning("Load a reference file (another data source) to use a lookup here.")
                        else:
                            ref_source = st.selectbox("Reference source", other_sources, key=f"tpl_ref_{header}")
                            ref_df = st.session_state.data_sources[ref_source]
                            col1, col2 = st.columns(2)
                            with col1:
                                st.selectbox("Match using column (current data)", df.columns, key=f"tpl_key_{header}")
                            with col2:
                                st.selectbox("Match against column (reference)", ref_df.columns, key=f"tpl_refkey_{header}")
                            st.selectbox("Return this column from reference", ref_df.columns, key=f"tpl_return_{header}")
                            st.text_input("Default value if not found (optional)", key=f"tpl_default_{header}")

                    elif mode == mode_options[2]:
                        st.text_area(
                            "Formula (supports [Columns], XLOOKUP('Source','Key','Return'), IF/THEN/ELSE, etc.)",
                            "[Column1]", key=f"tpl_formula_{header}"
                        )

                    else:
                        st.text_input("Static value for every row", key=f"tpl_static_{header}")

            if st.button("🏗️ Build Output From Template", use_container_width=True, key="tpl_build_btn"):
                all_sources = st.session_state.data_sources
                column_configs = {}
                for header in st.session_state.template_headers:
                    mode = st.session_state.get(f"tpl_mode_{header}", mode_options[0])
                    if mode == mode_options[0]:
                        column_configs[header] = {
                            'mode': 'direct',
                            'source_column': st.session_state.get(f"tpl_direct_{header}")
                        }
                    elif mode == mode_options[1]:
                        column_configs[header] = {
                            'mode': 'lookup',
                            'ref_source': st.session_state.get(f"tpl_ref_{header}"),
                            'key_column': st.session_state.get(f"tpl_key_{header}"),
                            'ref_key_column': st.session_state.get(f"tpl_refkey_{header}"),
                            'ref_return_column': st.session_state.get(f"tpl_return_{header}"),
                            'default': st.session_state.get(f"tpl_default_{header}") or None
                        }
                    elif mode == mode_options[2]:
                        formula_text = st.session_state.get(f"tpl_formula_{header}", "")
                        column_configs[header] = {
                            'mode': 'formula',
                            'formula': formula_text,
                            'referenced_sources': detect_referenced_source_names(formula_text, all_sources.keys())
                        }
                    else:
                        column_configs[header] = {
                            'mode': 'static',
                            'value': st.session_state.get(f"tpl_static_{header}", '')
                        }

                result_df = MultiSheetProcessor.build_template_output(
                    df, st.session_state.template_headers, column_configs, all_sources
                )
                result_name = f"{source_name}_template_output"
                st.session_state.data_sources[result_name] = result_df
                st.session_state.source_metadata[result_name] = {
                    'file': 'Template Output', 'sheet': 'Built',
                    'rows': len(result_df), 'columns': len(result_df.columns)
                }
                build_op = {
                    'type': 'build_template',
                    'template_headers': list(st.session_state.template_headers),
                    'column_configs': column_configs
                }
                st.session_state.source_snapshots[result_name] = df.copy()
                st.session_state.operations = [build_op]
                st.session_state.operations_undo_stack = []
                st.session_state.active_source = result_name
                st.success(f"✅ Built output '{result_name}' with {len(result_df)} rows")
                st.rerun()

    # Standard Operations
    st.divider()
    st.header("⚡ Data Operations")

    if st.session_state.current_operation:
        col1, col2 = st.columns([6, 1])
        with col2:
            if st.button("❌ Close", use_container_width=True, key="close_op_btn"):
                st.session_state.current_operation = None
                st.rerun()

    if not st.session_state.current_operation:
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            if st.button("🔍 Filter Rows", use_container_width=True, key="op_filter_btn"):
                st.session_state.current_operation = "filter"; st.rerun()
            if st.button("🔍 Multiple Filters", use_container_width=True, key="op_multi_filter_btn"):
                st.session_state.current_operation = "filter_multiple"; st.rerun()
            if st.button("📐 Add Column", use_container_width=True, key="op_calc_btn"):
                st.session_state.current_operation = "calc"; st.rerun()
            if st.button("📊 Group By", use_container_width=True, key="op_group_btn"):
                st.session_state.current_operation = "group"; st.rerun()
        with col2:
            if st.button("🔃 Sort Data", use_container_width=True, key="op_sort_btn"):
                st.session_state.current_operation = "sort"; st.rerun()
            if st.button("✏️ Rename Column", use_container_width=True, key="op_rename_btn"):
                st.session_state.current_operation = "rename"; st.rerun()
            if st.button("🗑️ Drop Column", use_container_width=True, key="op_drop_btn"):
                st.session_state.current_operation = "drop"; st.rerun()
        with col3:
            if st.button("📥 Append Data", use_container_width=True, key="op_append_btn"):
                st.session_state.current_operation = "append"; st.rerun()
            if st.button("📈 Pivot Table", use_container_width=True, key="op_pivot_btn"):
                st.session_state.current_operation = "pivot"; st.rerun()
            if st.button("📊 Summarize by Ref", use_container_width=True, key="op_summarize_btn"):
                st.session_state.current_operation = "summarize"; st.rerun()
        with col4:
            if st.button("💾 Execute All", use_container_width=True, key="op_execute_btn"):
                st.session_state.current_operation = "execute"; st.rerun()
            if st.button("📥 Export", use_container_width=True, key="op_export_btn"):
                st.session_state.current_operation = "export"; st.rerun()

    if st.session_state.current_operation:
        st.divider()
        st.markdown('<div class="operation-active">', unsafe_allow_html=True)
        st.markdown(f"### 🔧 Operation: {st.session_state.current_operation.title()}")

        op_selected = st.session_state.current_operation

        if op_selected == "filter":
            col1, col2 = st.columns(2)
            with col1:
                column = st.selectbox("Select column to filter", df.columns, key="filter_column")
                # Show column info
                dtype = df[column].dtype
                st.caption(f"Data type: {dtype}")
                
                # Show unique values count
                unique_count = df[column].nunique()
                st.caption(f"Unique values: {unique_count}")
                
            with col2:
                operation = st.selectbox("Operation", ['==', '!=', '>', '<', '>=', '<=', 'contains'], key="filter_operation")
                
                # DROPDOWN FOR FILTER VALUES
                if dtype == 'object' or pd.api.types.is_string_dtype(dtype):
                    # For text columns - show dropdown with unique values
                    unique_values, has_more = get_unique_values_sorted(df, column)
                    if unique_values:
                        if has_more:
                            st.caption(f"Showing first {len(unique_values)} of {df[column].nunique()} values")
                        value = st.selectbox(
                            "Select value to filter", 
                            options=unique_values,
                            key="filter_value_dropdown"
                        )
                    else:
                        value = st.text_input("Enter value to filter", key="filter_value_text")
                
                elif pd.api.types.is_datetime64_any_dtype(dtype):
                    # For date columns - show date picker
                    st.caption("Date column detected - select a date")
                    min_date = df[column].min()
                    max_date = df[column].max()
                    st.caption(f"Date range: {min_date.date()} to {max_date.date()}")
                    min_d = min_date.date() if pd.notna(min_date) else None
                    max_d = max_date.date() if pd.notna(max_date) else None
                    # Default to today only if today actually falls within the
                    # data's range; otherwise default to the closest valid date.
                    # Using "today" unconditionally breaks st.date_input whenever
                    # the data doesn't happen to span today (a StreamlitAPIException).
                    today = pd.Timestamp.now().date()
                    if min_d is not None and today < min_d:
                        default_d = min_d
                    elif max_d is not None and today > max_d:
                        default_d = max_d
                    else:
                        default_d = today
                    value = st.date_input(
                        "Select date", 
                        value=default_d,
                        min_value=min_d,
                        max_value=max_d,
                        key="filter_date_value"
                    )
                    value = value.strftime("%Y-%m-%d")
                
                elif pd.api.types.is_numeric_dtype(dtype):
                    # For numeric columns - show number input
                    min_val = df[column].min()
                    max_val = df[column].max()
                    st.caption(f"Range: {min_val} to {max_val}")
                    
                    # For operations that compare, show number input
                    if operation in ['==', '!=']:
                        # For equality, show dropdown of unique values if not too many
                        unique_values, has_more = get_unique_values_sorted(df, column)
                        if len(unique_values) <= 50:  # Only show dropdown for reasonable number of unique values
                            if has_more:
                                st.caption(f"Showing first {len(unique_values)} of {df[column].nunique()} values")
                            value = st.selectbox(
                                "Select value", 
                                options=unique_values,
                                key="filter_numeric_dropdown"
                            )
                        else:
                            value = st.number_input(
                                "Enter value",
                                value=float(min_val) if pd.notna(min_val) else 0.0,
                                key="filter_numeric_input"
                            )
                    else:
                        # For comparison operations, use number input
                        value = st.number_input(
                            "Enter value",
                            value=float(min_val) if pd.notna(min_val) else 0.0,
                            key="filter_numeric_input"
                        )
                else:
                    # Fallback - text input
                    value = st.text_input("Enter value to filter", key="filter_value_fallback")

            if st.button("✅ Apply Filter", use_container_width=True, key="apply_filter_btn"):
                # Convert value to string for contains operation
                if operation == 'contains':
                    value = str(value)
                new_df = MultiSheetProcessor.filter_rows(df, column, operation, value)
                if len(new_df) < len(df):
                    push_undo_snapshot()
                    st.session_state.operations.append({
                        'type': 'filter', 'column': column, 'operation': operation, 'value': str(value)
                    })
                    recompute_active_source()
                    st.success(f"✅ Filtered from {len(df)} to {len(new_df)} rows")
                    st.rerun()
                else:
                    st.warning("No rows were filtered. Check your filter criteria.")

        elif op_selected == "filter_multiple":
            st.info(
                "Add multiple filter conditions, then choose how they combine. "
                "**AND** = a row must match every condition (use this for narrowing down, "
                "e.g. Status == Pending AND Region == East). "
                "**OR** = a row must match at least one condition (use this to combine "
                "alternatives on the same column, e.g. Status == Pending OR Status == In Review — "
                "AND can never match this since one column can't equal two values at once)."
            )
            st.session_state.filter_logic = st.radio(
                "Combine conditions using",
                ["AND", "OR"],
                index=0 if st.session_state.filter_logic == "AND" else 1,
                horizontal=True,
                key="multi_filter_logic_radio"
            )

            # Add new condition
            col1, col2, col3, col4 = st.columns([3, 2, 2, 1])
            with col1:
                new_column = st.selectbox("Column", df.columns, key="multi_filter_col")
            with col2:
                ops_list = ['==', '!=', '>', '<', '>=', '<=', 'contains']
                new_operation = st.selectbox("Operation", ops_list, key="multi_filter_op")
            with col3:
                # Get unique values for dropdown
                dtype = df[new_column].dtype
                if pd.api.types.is_datetime64_any_dtype(dtype):
                    # Date input
                    st.caption("Select date")
                    min_date = df[new_column].min()
                    max_date = df[new_column].max()
                    min_d = min_date.date() if pd.notna(min_date) else None
                    max_d = max_date.date() if pd.notna(max_date) else None
                    today = pd.Timestamp.now().date()
                    if min_d is not None and today < min_d:
                        default_d = min_d
                    elif max_d is not None and today > max_d:
                        default_d = max_d
                    else:
                        default_d = today
                    new_value = st.date_input(
                        "Date", 
                        value=default_d,
                        min_value=min_d,
                        max_value=max_d,
                        key="multi_filter_date"
                    )
                    new_value = new_value.strftime("%Y-%m-%d")
                elif dtype == 'object' or pd.api.types.is_string_dtype(dtype):
                    # Text dropdown
                    unique_values, has_more = get_unique_values_sorted(df, new_column)
                    if unique_values:
                        if has_more:
                            st.caption(f"Showing first {len(unique_values)} values")
                        new_value = st.selectbox(
                            "Value", 
                            options=unique_values,
                            key="multi_filter_text_dropdown"
                        )
                    else:
                        new_value = st.text_input("Value", key="multi_filter_text")
                elif pd.api.types.is_numeric_dtype(dtype):
                    # Numeric input
                    min_val = df[new_column].min()
                    max_val = df[new_column].max()
                    st.caption(f"Range: {min_val} to {max_val}")
                    unique_values, has_more = get_unique_values_sorted(df, new_column)
                    if len(unique_values) <= 50:
                        new_value = st.selectbox(
                            "Value", 
                            options=unique_values,
                            key="multi_filter_num_dropdown"
                        )
                    else:
                        new_value = st.number_input(
                            "Value",
                            value=float(min_val) if pd.notna(min_val) else 0.0,
                            key="multi_filter_num"
                        )
                else:
                    new_value = st.text_input("Value", key="multi_filter_fallback")
            with col4:
                if st.button("➕ Add", key="add_filter_condition"):
                    if new_column and new_value is not None:
                        st.session_state.filter_conditions.append({
                            'column': new_column,
                            'operation': new_operation,
                            'value': str(new_value)
                        })
                        st.rerun()
            
            # Show existing conditions
            if st.session_state.filter_conditions:
                joiner = f" {st.session_state.filter_logic} "
                st.write(f"Current filter conditions (combined with **{st.session_state.filter_logic}**):")
                for i, cond in enumerate(st.session_state.filter_conditions):
                    col1, col2 = st.columns([5, 1])
                    with col1:
                        prefix = "" if i == 0 else f"{st.session_state.filter_logic} "
                        st.write(f"{i+1}. {prefix}{cond['column']} {cond['operation']} {cond['value']}")
                    with col2:
                        if st.button("🗑️", key=f"remove_filter_{i}"):
                            st.session_state.filter_conditions.pop(i)
                            st.rerun()
                
                if st.button("✅ Apply All Filters", use_container_width=True):
                    logic = st.session_state.filter_logic
                    new_df = MultiSheetProcessor.filter_multiple_rows(df, st.session_state.filter_conditions, logic)
                    push_undo_snapshot()
                    st.session_state.operations.append({
                        'type': 'filter_multiple',
                        'conditions': st.session_state.filter_conditions.copy(),
                        'logic': logic
                    })
                    st.session_state.filter_conditions = []  # Clear after applying
                    recompute_active_source()
                    st.success(f"✅ Filtered from {len(df)} to {len(new_df)} rows")
                    st.rerun()
            else:
                st.info("Add at least one filter condition above")

        elif op_selected == "calc":
            st.info("""
            **Formula Examples:**
            - `[Sales] * 0.1` (10% commission)
            - `([Price] - [Cost]) / [Cost] * 100` (profit margin %)
            - `[First Name] + ' ' + [Last Name]` (full name)
            - `IF [Sales] > 1000 THEN 'High' ELSE 'Low' END`
            - `abs([Value])` (absolute value)
            - `round([Price], 2)` (round to 2 decimals)
            - `upper([Name])` / `lower([Name])`
            - `XLOOKUP([Product ID], 'Products', 'ID', 'Unit Price')` — Excel-XLOOKUP-style
              lookup into any other loaded sheet/file. Optional 5th argument is a
              fallback value for no match, e.g. `XLOOKUP([SKU], 'Catalog', 'SKU', 'Price', 0)`
            """)
            new_column = st.text_input("New column name", f"calculated_{len(df.columns)}", key="calc_newcol")
            formula = st.text_area("Formula", "[Column1] + [Column2]", key="calc_formula")

            if st.button("✅ Add Column", use_container_width=True, key="apply_calc_btn"):
                if new_column and formula:
                    all_sources = st.session_state.data_sources
                    new_df = MultiSheetProcessor.add_calculated_column(df, new_column, formula, all_sources)
                    if new_column in new_df.columns:
                        push_undo_snapshot()
                        referenced = detect_referenced_source_names(formula, all_sources.keys())
                        st.session_state.operations.append({
                            'type': 'calc_column', 'new_column': new_column, 'formula': formula,
                            'referenced_sources': referenced
                        })
                        recompute_active_source()
                        st.success(f"✅ Added column: {new_column}")
                        st.rerun()
                else:
                    st.warning("Please provide both column name and formula")

        elif op_selected == "group":
            col1, col2 = st.columns(2)
            with col1:
                group_col = st.selectbox("Group by column", df.columns, key="group_by_col")
            with col2:
                agg_col = st.selectbox("Column to aggregate", [c for c in df.columns if c != group_col], key="group_agg_col")
            agg_func = st.selectbox("Aggregation function", ['sum', 'mean', 'median', 'min', 'max', 'count', 'std'], key="group_agg_func")

            if st.button("✅ Apply Group By", use_container_width=True, key="apply_group_btn"):
                new_df = MultiSheetProcessor.group_by(df, group_col, agg_col, agg_func)
                push_undo_snapshot()
                st.session_state.operations.append({
                    'type': 'group_by', 'group_column': group_col,
                    'aggregate_column': agg_col, 'function': agg_func
                })
                recompute_active_source()
                st.success(f"✅ Grouped data: {len(new_df)} groups")
                st.rerun()

        elif op_selected == "sort":
            sort_col = st.selectbox("Sort by column", df.columns, key="sort_col")
            ascending = st.radio("Order", ["Ascending", "Descending"], horizontal=True, key="sort_order")

            if st.button("✅ Apply Sort", use_container_width=True, key="apply_sort_btn"):
                push_undo_snapshot()
                st.session_state.operations.append({
                    'type': 'sort', 'column': sort_col, 'ascending': ascending == "Ascending"
                })
                recompute_active_source()
                st.success(f"✅ Sorted by {sort_col}")
                st.rerun()

        elif op_selected == "rename":
            old_name = st.selectbox("Select column to rename", df.columns, key="rename_old")
            new_name = st.text_input("New name", f"{old_name}_new", key="rename_new")

            if st.button("✅ Apply Rename", use_container_width=True, key="apply_rename_btn"):
                if new_name and new_name not in df.columns:
                    push_undo_snapshot()
                    st.session_state.operations.append({
                        'type': 'rename', 'old_name': old_name, 'new_name': new_name
                    })
                    recompute_active_source()
                    st.success(f"✅ Renamed '{old_name}' to '{new_name}'")
                    st.rerun()
                else:
                    st.warning("New name is empty or already exists")

        elif op_selected == "drop":
            col_to_drop = st.selectbox("Select column to drop", df.columns, key="drop_col")
            confirmed = st.checkbox(f"Confirm dropping '{col_to_drop}'", key="drop_confirm")

            if st.button("✅ Apply Drop", use_container_width=True, key="apply_drop_btn", disabled=not confirmed):
                push_undo_snapshot()
                st.session_state.operations.append({'type': 'drop', 'column': col_to_drop})
                recompute_active_source()
                st.success(f"✅ Dropped column: {col_to_drop}")
                st.rerun()

        elif op_selected == "append":
            st.info("Append rows from another data source to the current data")
            append_source = st.selectbox(
                "Source to append",
                [s for s in st.session_state.data_sources.keys() if s != source_name],
                key="append_source_select"
            )
            if append_source:
                append_df = st.session_state.data_sources[append_source]
                st.write(f"Will append {len(append_df)} rows from '{append_source}'")
                missing_cols = [c for c in df.columns if c not in append_df.columns]
                if missing_cols:
                    st.warning(f"Missing columns in append source: {missing_cols}")
                    st.write("They will be filled with NaN")

                if st.button("✅ Append Data", use_container_width=True, key="apply_append_btn"):
                    new_df = MultiSheetProcessor.append_data(df, append_df)
                    push_undo_snapshot()
                    st.session_state.operations.append({'type': 'append', 'append_source': append_source})
                    recompute_active_source()
                    st.success(f"✅ Appended {len(append_df)} rows. New total: {len(new_df)}")
                    st.rerun()

        elif op_selected == "pivot":
            col1, col2 = st.columns(2)
            with col1:
                index = st.selectbox("Index column", df.columns, key="pivot_index")
                columns = st.selectbox("Columns to pivot", [c for c in df.columns if c != index], key="pivot_columns")
            with col2:
                values = st.selectbox("Values column", [c for c in df.columns if c not in [index, columns]], key="pivot_values")
                aggfunc = st.selectbox("Aggregation", ['sum', 'mean', 'count', 'min', 'max'], key="pivot_aggfunc")

            if st.button("✅ Create Pivot Table", use_container_width=True, key="apply_pivot_btn"):
                new_df = MultiSheetProcessor.pivot_table(df, index, columns, values, aggfunc)
                push_undo_snapshot()
                st.session_state.operations.append({
                    'type': 'pivot', 'index': index, 'columns': columns,
                    'values': values, 'aggfunc': aggfunc
                })
                recompute_active_source()
                st.success(f"✅ Pivot table created: {len(new_df)} rows")
                st.rerun()

        elif op_selected == "summarize":
            st.info("Summarize data and join with reference data")
            ref_source = st.selectbox(
                "Reference source",
                [s for s in st.session_state.data_sources.keys() if s != source_name],
                key="summarize_ref_source"
            )
            if ref_source:
                ref_df = st.session_state.data_sources[ref_source]
                col1, col2 = st.columns(2)
                with col1:
                    group_col = st.selectbox("Group by column in current data", df.columns, key="summarize_group_col")
                    ref_col = st.selectbox("Reference column", ref_df.columns, key="summarize_ref_col")
                with col2:
                    agg_col = st.selectbox("Column to aggregate", [c for c in df.columns if c != group_col], key="summarize_agg_col")
                    agg_func = st.selectbox("Aggregation function", ['sum', 'mean', 'count', 'min', 'max'], key="summarize_agg_func")

                if st.button("✅ Summarize", use_container_width=True, key="apply_summarize_btn"):
                    new_df = MultiSheetProcessor.summarize_by_reference(df, group_col, agg_col, ref_df, ref_col, agg_func)
                    push_undo_snapshot()
                    st.session_state.operations.append({
                        'type': 'summarize_by_ref', 'ref_source': ref_source,
                        'group_col': group_col, 'agg_col': agg_col,
                        'ref_col': ref_col, 'agg_func': agg_func
                    })
                    recompute_active_source()
                    st.success(f"✅ Summarized data: {len(new_df)} rows")
                    st.rerun()

        elif op_selected == "execute":
            if st.session_state.operations:
                st.info(
                    "Each step already reruns automatically as you build. Use this if a "
                    "reference source changed and you want to refresh the result."
                )
                st.subheader(f"Current recipe on '{source_name}' ({len(st.session_state.operations)} step(s)):")
                for i, op in enumerate(st.session_state.operations, 1):
                    st.write(f"{i}. {describe_operation(op)}")

                if st.button("🔄 Recompute From Scratch", use_container_width=True, key="execute_all_btn"):
                    recompute_active_source()
                    new_df = st.session_state.data_sources[source_name]
                    st.success(f"✅ Recomputed! New shape: {new_df.shape}")
                    st.rerun()
            else:
                st.info("No operations to execute")

        elif op_selected == "export":
            st.subheader("📥 Export Data")
            export_option = st.radio(
                "Export:", ["Current source only", "All sources (multiple sheets)"],
                horizontal=True, key="export_option"
            )

            if export_option == "Current source only":
                col1, col2 = st.columns(2)
                with col1:
                    csv = df.to_csv(index=False)
                    st.download_button(
                        "📊 Download CSV", data=csv, file_name=f"{source_name}.csv",
                        mime="text/csv", use_container_width=True, key="download_csv_btn"
                    )
                with col2:
                    excel_data = to_excel_bytes({source_name: df})
                    st.download_button(
                        "📊 Download Excel", data=excel_data, file_name=f"{source_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True, key="download_xlsx_btn"
                    )
            else:
                excel_data = to_excel_bytes(st.session_state.data_sources)
                st.download_button(
                    "📊 Download All Sources",
                    data=excel_data,
                    file_name=f"all_sources_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, key="download_all_btn"
                )

            st.info(f"📊 Exporting {len(df)} rows and {len(df.columns)} columns")

        st.markdown('</div>', unsafe_allow_html=True)

else:
    st.divider()
    if st.session_state.data_sources:
        st.markdown(f"""
        <div style="text-align: center; padding: 30px;">
            <h2>📊 Select a Data Source</h2>
            <p>Click on any source in the sidebar to start processing</p>
            <p style="color: #666; font-size: 0.9rem;">
                You have <strong>{len(st.session_state.data_sources)}</strong> data sources loaded
            </p>
        </div>
        """, unsafe_allow_html=True)
        st.subheader("📋 All Loaded Sources")
        for name, df in st.session_state.data_sources.items():
            with st.expander(f"📊 {name} ({len(df)} rows, {len(df.columns)} cols)"):
                st.dataframe(df.head(), use_container_width=True)
    else:
        st.markdown("""
        <div style="text-align: center; padding: 50px;">
            <h2>🚀 Welcome to Multi-Sheet Data Processing!</h2>
            <p style="font-size: 1.2rem; color: #666;">
                Upload Excel files or CSVs to get started
            </p>
            <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; margin: 20px 0;">
                <h4>✨ What you can do:</h4>
                <p>📚 Load multiple Excel sheets and files</p>
                <p>🔗 Create relationships between data sources (like Excel VLOOKUP)</p>
                <p>🔄 Reference data across sheets for filtering and calculations</p>
                <p>📊 Process data step by step with a visual interface</p>
                <p>💾 Save your process as a recipe you can reuse</p>
                <p>🤖 Run a saved recipe automatically on new reports</p>
                <p>📥 Export as single or multiple sheets</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

st.divider()
st.markdown("""
<div style="text-align: center; color: #666; font-size: 0.9rem;">
    Multi-Sheet Data Processing Wizard v2.1 | Built with Streamlit and Pandas
</div>
""", unsafe_allow_html=True)