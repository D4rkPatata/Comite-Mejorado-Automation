import pandas as pd


def pivot_data(df, index, columns=None, values="TOTAL_VALORIZADO", aggfunc="sum", filters=None):
    """
    Builds a pivot table, the pandas equivalent of an Excel TablaDinámica.

    Args:
        df (pd.DataFrame): Source data.
        index (str | list): Row fields.
        columns (str | list): Column fields.
        values (str | list): Value fields.
        aggfunc (str | dict): Aggregation, e.g. "sum" or {"CANTIDAD": "sum"}.
        filters (dict): Page filters, e.g. {"TIPO_VENTA": ["PISO"], "AÑO": [2025, 2026]}.

    Returns:
        pd.DataFrame: The pivot table.
    """
    if filters:
        for field, allowed in filters.items():
            df = df[df[field].isin(allowed)]
    return pd.pivot_table(
        df, index=index, columns=columns, values=values,
        aggfunc=aggfunc, observed=True, fill_value=0,
    )


def group_data(df, schema):
    """
    Groups by the schema's dimension fields ("str" / "int") and sums its
    "float" fields, collapsing the cache detail (month, city, cluster, ...).

    Returns:
        pd.DataFrame: One row per combination of dimensions.
    """
    dims = [c for c, t in schema.items() if t != "float"]
    measures = [c for c, t in schema.items() if t == "float"]
    return df.groupby(dims, observed=True, as_index=False)[measures].sum()
