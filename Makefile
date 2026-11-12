CHROME ?= $(or $(lastword $(sort $(wildcard $(HOME)/.cache/ms-playwright/chromium_headless_shell-*/*/chrome-headless-shell))),$(shell command -v chromium || command -v google-chrome))

.PHONY: help build update docx clean
.DEFAULT_GOAL := help

help: ## Show this help
	@awk 'BEGIN {FS = ":.*## "} /^[a-z]+:.*## / {printf "  make %-8s %s\n", $$1, $$2}' $(MAKEFILE_LIST)

build: build/resume.pdf ## Render index.html to build/resume.pdf for preview
	@echo "preview: $(CURDIR)/build/resume.pdf ($$(pdfinfo $< | awk '/^Pages/{print $$2}') pages)"

build/resume.pdf: index.html
	@mkdir -p build
	@"$(CHROME)" --headless --disable-gpu --no-pdf-header-footer \
		--print-to-pdf="$(CURDIR)/$@" "file://$(CURDIR)/index.html" 2>/dev/null

update: build ## Render, then copy the preview over the published resume.pdf
	cp build/resume.pdf resume.pdf

docx: build/resume.docx ## Write build/resume.docx for importing into Google Docs
	@echo "docx: $(CURDIR)/build/resume.docx"

build/resume.docx: index.html html2docx.py
	@mkdir -p build
	@python3 -I html2docx.py index.html $@

clean: ## Delete the build/ directory
	rm -rf build
