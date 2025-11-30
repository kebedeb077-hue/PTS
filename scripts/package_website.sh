#!/usr/bin/env bash
set -euo pipefail

SRC="Publications_Tracking_System"
OUT_DIR="deploy"
OUT_ZIP="$OUT_DIR/${SRC}.zip"

if [ ! -d "$SRC" ]; then
  echo "Source folder '$SRC' not found. Run this script from repository root." >&2
  exit 2
fi

mkdir -p "$OUT_DIR"

echo "Scanning for .mdb files in $SRC..."
find "$SRC" -type f -iname "*.mdb" -print > "$OUT_DIR/mdb_files.txt" || true

if [ -s "$OUT_DIR/mdb_files.txt" ]; then
  echo "Found .mdb files (listed in $OUT_DIR/mdb_files.txt):"
  cat "$OUT_DIR/mdb_files.txt"
  echo "Ensure you handle these database files securely before publishing to a public host. They will be included in the zip by default." 
fi

echo "Creating zip archive: $OUT_ZIP"
# Change into source dir so zip paths are clean
(cd "$SRC" && zip -r "../$OUT_ZIP" . -x "*_vti_cnf/*" "*_vti_pvt/*" "*_vti/*" "Thumbs.db" "desktop.ini")

echo "Archive created: $OUT_ZIP"
echo "Deployable artifact and a list of .mdb files are in the '$OUT_DIR' directory. Review $OUT_DIR/mdb_files.txt before publishing."

exit 0
