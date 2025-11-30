#!/usr/bin/env bash
set -euo pipefail

# Script to export .mdb files to CSV using mdbtools or python
# Usage: ./export_mdb_to_csv.sh <path-to-mdb-file> <output-dir>

MDB_FILE="${1:?Error: mdb file path required}"
OUTPUT_DIR="${2:?Error: output directory required}"

if [ ! -f "$MDB_FILE" ]; then
  echo "Error: $MDB_FILE not found" >&2
  exit 1
fi

mkdir -p "$OUTPUT_DIR"

# Try using mdbtools first
if command -v mdbexport &> /dev/null; then
  echo "Using mdbtools to export tables..."
  
  # Get list of tables
  tables=$(mdblist "$MDB_FILE" | grep -v "^Database:" || true)
  
  for table in $tables; do
    echo "Exporting $table..."
    mdbexport "$MDB_FILE" "$table" > "$OUTPUT_DIR/${table}.csv"
  done
  
  echo "✓ Export completed to $OUTPUT_DIR"
else
  echo "Error: mdbtools not found. Install with: apt-get install mdbtools" >&2
  exit 1
fi
