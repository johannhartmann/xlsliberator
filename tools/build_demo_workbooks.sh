#!/bin/sh
set -eu

if [ ! -f /.dockerenv ] || [ "${XLSLIBERATOR_APPLICATION_CONTAINER:-}" != "1" ]; then
  echo "Demo workbooks must be built inside the XLSLiberator Docker test container." >&2
  exit 2
fi

root=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)
package="$root/demos/hostile-workbook/source-package"
output="$root/demos/hostile-workbook/source/HostileButInert.xlsx"
temporary_directory=$(mktemp -d)
temporary_output="$temporary_directory/HostileButInert.xlsx"

mkdir -p "$(dirname -- "$output")"
find "$package" -type f -exec touch -t 202607180000 {} +
(
  cd "$package"
  zip -X -0 -q "$temporary_output" \
    '[Content_Types].xml' \
    '_rels/.rels' \
    'docProps/app.xml' \
    'docProps/core.xml' \
    'xl/workbook.xml' \
    'xl/_rels/workbook.xml.rels' \
    'xl/worksheets/sheet1.xml'
)
mv "$temporary_output" "$output"
rmdir "$temporary_directory"
