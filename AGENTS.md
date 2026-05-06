# Ironsheet Agent Rules

These repo-local rules apply to future automated work in this project.

## Product Direction

- Ironsheet is a preservation-first OOXML workbook engine, not a generic XLSX writer.
- The product promise is: "Move fast and break no spreadsheets."
- Optimize for lossless mutation of real Excel workbooks before broad feature checklists.
- Treat diagnostics, validation, and compatibility fixtures as first-class product surfaces.

## Implementation Rules

- Use TypeScript for implementation, scripts, tests, fixtures, and tooling logic.
- Do not add JavaScript scripts under `scripts/`; `npm run verify` enforces this.
- Keep runtime code dependency-light and justify every dependency.
- Do not import Node built-ins from runtime-neutral core packages.
- Prefer structured parsers and typed models over ad hoc string manipulation.

## Workflow

- Run `npm run ci` before handing work back.
- Use `npm run commit:safe -- "message"` for local commits.
- Do not commit `IRONSHEET_SPEC.md`; it is intentionally ignored planning material.
- Add or update tests with behavior changes.
