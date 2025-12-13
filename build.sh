#!/bin/bash

set -e

VERSION=$(grep '"version"' manifest.json | head -1 | sed 's/.*"version": "\(.*\)".*/\1/')
OUTPUT="dist/session-sushi-v${VERSION}.zip"

echo "Building Session Sushi v${VERSION}..."

rm -rf build dist
mkdir -p build dist

cp -r src icons manifest.json README.md LICENSE PRIVACY.md RELEASE.md build/

cd build
zip -rq "../${OUTPUT}" .
cd ..

echo "Done: ${OUTPUT}"
sha256sum "${OUTPUT}"
