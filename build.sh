#!/bin/bash
set -e
cd "$(dirname "$0")"

VERSION=$(grep '"version"' manifest.json | head -1 | sed 's/.*: *"\(.*\)".*/\1/')
OUTPUT="refdetect-${VERSION}.xpi"

rm -f "$OUTPUT"
zip -r "$OUTPUT" manifest.json bootstrap.js content/ -x "*.DS_Store"
echo "Built $OUTPUT"
