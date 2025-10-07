def sanitize_excel_values(df):
    df = df.copy()
    for col in df.columns:
        df[col] = df[col].apply(
            lambda x: f"'{x}" if isinstance(x, str) and x.strip().startswith('=') else x
        )
    return df