import re
import pandas as pd

def find_column(df, possible_matches, how="exact", required=True):
    """
    Searches df.columns for the first column that matches any of the possible_matches.
    
    Parameters
    ----------
    df : pd.DataFrame
        The DataFrame in which to search for columns.
    possible_matches : list of str
        A list of potential column names or patterns to look for.
    how : str, optional
        Matching mode:
         - "exact"  => requires exact match
         - "partial" => checks if `possible_match` is a substring of the column name
         - "regex"   => interprets `possible_match` as a regex
    required : bool, optional
        If True, raises an error if no column is found; otherwise returns None.

    Returns
    -------
    str or None
        The actual column name in df.columns that was matched, or None if not found
        (and required=False).
    """
    df_cols = list(df.columns)
    for col in df_cols:
        for pattern in possible_matches:
            if how == "exact":
                if col.strip().lower() == pattern.strip().lower():
                    return col
            elif how == "partial":
                if pattern.strip().lower() in col.lower():
                    return col
            elif how == "regex":
                if re.search(pattern, col, flags=re.IGNORECASE):
                    return col
    
    if required:
        raise ValueError(
            f"Could not find a required column. Tried {possible_matches} in columns: {df.columns.tolist()}"
        )
    return None


def rename_columns(df, rename_map, how="exact"):
    """
    Given a dictionary { new_col_name: [list of possible appearances] },
    search & rename them in the DataFrame if found.

    Parameters
    ----------
    df : pd.DataFrame
        The DataFrame whose columns will be renamed in-place.
    rename_map : dict
        Keys = new/standardized column name,
        Values = list of possible matches for that column in the DF.
    how : str
        "exact", "partial", or "regex".

    Returns
    -------
    df : pd.DataFrame
        Same DataFrame with renamed columns.
    """
    for new_col_name, patterns in rename_map.items():
        old_name = find_column(df, patterns, how=how, required=False)
        if old_name:
            df.rename(columns={old_name: new_col_name}, inplace=True)
    return df
