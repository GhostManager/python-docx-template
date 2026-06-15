# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.21.0] * 15 June 2026

### Added

* Regression coverage for custom Jinja delimiters in headers and footers
* Regression coverage for mutable file-like `InlineImage` descriptors
* Regression coverage for fast `InlineImage` XML template generation and fallback behavior
* Repository ignore rules for macOS `.DS_Store` files

### Changed

* Mirrored upstream `dev` branch rendering optimizations, including faster XML parsing and body replacement, reduced header/footer processing when no Jinja tags are present, precompiled tag-stripping regexes, and early exit for listing resolution when no listing control characters are present
* Improved `InlineImage` performance by prebuilding reusable image XML, deduplicating image parts by stable descriptors, caching image metadata instead of full XML, deriving image part counters from existing media names, and initializing drawing IDs from existing document, header, footer, and footnote IDs
* Updated fast `InlineImage` XML template generation to write placeholders into explicit XML attributes instead of replacing numeric sentinel values in serialized XML
* Added a native `python-docx` fallback for `InlineImage` XML generation if the optimized template is incompatible with a future `python-docx` XML shape
* Merged upstream PR #626 to emit RichText font properties in the OOXML order expected by Word

### Fixed

* Fixed header and footer rendering when callers provide a `jinja_env` with custom delimiters, such as `variable_start_string="[[", variable_end_string="]]"`
* Fixed stale `InlineImage` cache reuse for hashable file-like descriptors such as `io.BytesIO` and open file handles
* Fixed `InlineImage` handling for `None` filenames
* Fixed Poetry project metadata required by the build configuration
