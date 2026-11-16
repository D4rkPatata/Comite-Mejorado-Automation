import zipfile

import numpy as np
import pandas as pd
from lxml import etree

from config import SCHEMA_DATA_BASE


NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
CACHE_DEF = "xl/pivotCache/pivotCacheDefinition1.xml"
CACHE_RECORDS = "xl/pivotCache/pivotCacheRecords1.xml"


def get_connection_info(file_path="input/data_lock.xlsx"):
    """
    Reads the external data connection stored in the workbook (xl/connections.xml).

    Returns:
        dict: name, connection string parts (server, database, user) and SQL command.
    """
    with zipfile.ZipFile(file_path) as z:
        root = etree.fromstring(z.read("xl/connections.xml"))

    conn = root.find(f"{NS}connection")
    db_pr = conn.find(f"{NS}dbPr")
    conn_str = db_pr.get("connection")
    parts = dict(
        p.split("=", 1) for p in conn_str.split(";") if "=" in p
    )
    return {
        "name": conn.get("name"),
        "server": parts.get("Data Source"),
        "database": parts.get("Initial Catalog"),
        "user": parts.get("User ID"),
        "command": db_pr.get("command"),
        "connection_string": conn_str,
    }


def _shared_value(el):
    tag = etree.QName(el).localname
    if tag == "m":
        return None
    v = el.get("v")
    if tag == "n":
        return float(v)
    if tag == "b":
        return v in ("1", "true")
    return v  # s, d (ISO date), e (error)


def extract_from_pivot_cache(file_path="input/data_lock.xlsx"):
    """
    Rebuilds the source table from the pivot cache saved inside the workbook
    (snapshot of the last refresh). Works offline, no credentials needed.

    Returns:
        pd.DataFrame: one row per cache record, indexed fields as categoricals.
    """
    with zipfile.ZipFile(file_path) as z:
        definition = etree.fromstring(z.read(CACHE_DEF))
        fields = definition.find(f"{NS}cacheFields").findall(f"{NS}cacheField")
        names = [f.get("name") for f in fields]
        shared = []
        for f in fields:
            items = f.find(f"{NS}sharedItems")
            shared.append([_shared_value(i) for i in items] if items is not None else [])

        n_records = int(definition.get("recordCount"))
        n_fields = len(names)
        codes = np.full((n_records, n_fields), -1, dtype=np.int32)
        values = np.full((n_records, n_fields), np.nan, dtype=np.float64)

        with z.open(CACHE_RECORDS) as fh:
            row = 0
            for _, rec in etree.iterparse(fh, tag=f"{NS}r", huge_tree=True):
                for col, el in enumerate(rec):
                    v = el.get("v")
                    tag = el.tag[len(NS):]
                    if tag == "x":
                        codes[row, col] = int(v)
                    elif tag == "n":
                        values[row, col] = float(v)
                rec.clear()
                while rec.getprevious() is not None:
                    del rec.getparent()[0]
                row += 1

    data = {}
    for col, name in enumerate(names):
        col_codes = codes[:row, col]
        items = shared[col]
        if not (col_codes >= 0).any():
            data[name] = values[:row, col]
        elif all(isinstance(i, str) for i in items):
            data[name] = pd.Categorical.from_codes(col_codes, categories=items)
        else:
            nums = np.array(items, dtype=np.float64)[col_codes]
            data[name] = nums.astype(np.int64) if np.all(nums == np.round(nums)) else nums
    return pd.DataFrame(data)


DTYPES = {"str": "category", "int": "int64", "float": "float64"}


def extract_data(file_path="input/data_lock.xlsx", schema=SCHEMA_DATA_BASE):
    """
    Extracts the workbook's source data from its pivot cache, keeping only the
    schema fields. The SQL connection is not reachable, so the cache saved by
    Excel at its last refresh is the source.

    Args:
        file_path (str): Path to the workbook.
        schema (dict): Field name -> "str" | "int" | "float".

    Returns:
        pd.DataFrame: The extracted data.
    """
    df = extract_from_pivot_cache(file_path)
    missing = [c for c in schema if c not in df.columns]
    if missing:
        raise KeyError(f"Fields not found in the pivot cache: {missing}")
    df = df[list(schema)].astype({c: DTYPES[t] for c, t in schema.items()})
    # Excel keeps items no longer present in the source (u="1"); drop them.
    for c in df.select_dtypes("category"):
        df[c] = df[c].cat.remove_unused_categories()
    return df
