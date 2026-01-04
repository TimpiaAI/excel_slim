# Contributing

## Module interface

Modules live in `crates/excel-slim-core/src/modules` and implement:

- `name()`
- `is_applicable(profile, options)`
- `run(ctx) -> ModuleResult`

Keep modules deterministic and avoid writing outside the output path.

## Fixtures

Test fixtures should be stored under `crates/excel-slim-core/tests/fixtures`.
Add new fixtures for any bug fix or module change.
