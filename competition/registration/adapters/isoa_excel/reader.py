import pandas as pd
from io import BytesIO


def read_isoa_excel(file_storage):
    """Read an ISOA startlist.xlsx and yield raw dict rows keyed by column index.

    ISOA startlist layout is positional:
      0=member ID, 1=name, 2=club, 3=class, 6=SI card, 9=note,
      10=rental flag, 15=phone
    """
    stream = file_storage.stream
    if hasattr(stream, "read"):
        stream = BytesIO(stream.read())
    df = pd.read_excel(stream, dtype=str, header=None)
    df = df.fillna("")
    # ISOA startlist files include a header row (STNO, NAME, CLUB, ...).
    # Skip it when the first column contains "STNO".
    start_at = 0
    if len(df) > 0:
        first_val = str(df.iloc[0, 0] or "").strip().lower()
        if first_val == "stno":
            start_at = 1
    for _, row in df.iloc[start_at:].iterrows():
        yield {str(idx): value for idx, value in enumerate(row.tolist())}
