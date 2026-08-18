import pandas as pd
import json
import sys
from pathlib import Path


def parse_recon_file(file_path: str, delimiter: str = '\t') -> list[dict]:
    """
    Parse a reconciliation/matching file and produce a JSON comparison
    between source systems grouped by MATCH ID.

    Args:
        file_path: Path to the input file (TSV, CSV, etc.)
        delimiter: Column delimiter (default: tab)

    Returns:
        List of comparison result dicts per MATCH ID
    """
    # Read the file
    df = pd.read_csv(file_path, sep=delimiter, dtype=str, keep_default_na=False)

    # Strip whitespace from column names
    df.columns = df.columns.str.strip()

    # Identify key columns
    match_id_col = 'MATCH ID'
    source_col = 'SOURCE SYSTEM'

    # Metadata columns that go into the header (not compared field-by-field)
    meta_columns = {match_id_col, source_col, 'REC', 'INSTANCE NO', 'BUSINESS DATE',
                    'BREAK COLUMN COUNT', 'BREAK COLUMNS', 'MATCH PASS', 'MATCH TYPE'}

    # Columns to compare (everything that's not metadata)
    compare_columns = [c for c in df.columns if c not in meta_columns]

    # Group by MATCH ID
    results = []

    for match_id, group in df.groupby(match_id_col, sort=False):
        # Get unique source systems for this match_id
        sources = group[source_col].unique().tolist()

        # Build a dict of source -> row(s) data
        # If multiple rows per source, take the first one (or you can customise)
        source_data = {}
        for src in sources:
            src_rows = group[group[source_col] == src]
            # Take first row per source for comparison
            source_data[src] = src_rows.iloc[0]

        # Extract common metadata from the first row that has it
        first_row = group.iloc[0]
        record_id = first_row.get('RecordID', '')
        message_id = first_row.get('MessageID', '')
        use_case_id = first_row.get('UseCaseID', '')

        # Build comparison for each field
        comparisons = []
        for col in compare_columns:
            field_entry = {"field": col}

            # Add value from each source system dynamically
            for src in sources:
                if src in source_data:
                    field_entry[src] = source_data[src].get(col, '')
                else:
                    field_entry[src] = ''

            # Determine match: all source values are equal (case-insensitive, stripped)
            values = [field_entry.get(src, '').strip().upper() for src in sources]
            if len(values) >= 2:
                field_entry['match'] = all(v == values[0] for v in values)
            else:
                field_entry['match'] = None  # Can't compare with single source

            comparisons.append(field_entry)

        result = {
            "match_id": match_id,
            "RecordID": record_id,
            "MessageID": message_id,
            "UseCaseID": use_case_id,
            "source_systems": sources,
            "match_type": first_row.get('MATCH TYPE', ''),
            "comparisons": comparisons
        }

        results.append(result)

    return results


def filter_breaks_only(results: list[dict]) -> list[dict]:
    """Optional: filter to only show fields where match is False."""
    filtered = []
    for r in results:
        breaks = [c for c in r['comparisons'] if c.get('match') is False]
        if breaks:
            filtered.append({**r, 'comparisons': breaks})
    return filtered


if __name__ == '__main__':
    # Usage: python recon_compare.py <file_path> [--breaks-only]
    if len(sys.argv) < 2:
        print("Usage: python recon_compare.py <file_path> [--breaks-only]")
        sys.exit(1)

    file_path = sys.argv[1]
    breaks_only = '--breaks-only' in sys.argv

    results = parse_recon_file(file_path)

    if breaks_only:
        results = filter_breaks_only(results)

    # Output JSON
    output = json.dumps(results, indent=2, ensure_ascii=False)
    print(output)

    # Optionally write to file
    out_path = Path(file_path).stem + '_compa.json'
    with open(out_path, 'w') as f:
        f.write(output)
    print(f"\nOutput written to: {out_path}")
