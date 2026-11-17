# Release Notes — Archon Tech Sewer Contributions

## 0.3.1 — source baseline (2026-10-28)
- Preserves the current Access VBA builder, finance services, query builder, UI builder, and legacy design/reference files.
- Static test suite: 20/20 passing in the current Linux environment.
- Audited source counts: 40 `CREATE TABLE` statements, 59 relationship definitions, 24 index ensure calls, 16 saved-query definitions, and 24 form-builder invocations in `AT_BuildInitialUI`.
- Added an explicit environment diagnostic and a requirements/status matrix to this handoff copy.
- Included prior Excel workbooks and original review/design files as historical references, not as substitutes for the Access system.

## Not a final product release
- No `.accdb` or `.accde` is included because this environment lacks Microsoft Access, Windows/Access automation, the ACE engine, and a viable offline Access file writer.
- The original source ZIP and its stored Library copy were not overwritten.
- No runtime or compile tests in Access are claimed.
