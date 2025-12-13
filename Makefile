.PHONY: build clean help

VERSION := $(shell grep '"version"' manifest.json | head -1 | sed 's/.*"version": "\(.*\)".*/\1/')
OUTPUT := dist/session-sushi-v$(VERSION).zip

build:
	@echo "Building Session Sushi v$(VERSION)..."
	@rm -rf build dist
	@mkdir -p build dist
	@cp -r src icons manifest.json README.md LICENSE PRIVACY.md RELEASE.md build/
	@cd build && zip -rq "../$(OUTPUT)" .
	@echo "Done: $(OUTPUT)"
	@sha256sum "$(OUTPUT)"

clean:
	@rm -rf build dist
	@echo "Cleaned build artifacts"

help:
	@echo "Session Sushi Build"
	@echo ""
	@echo "make build  - Build extension package"
	@echo "make clean  - Remove build artifacts"
