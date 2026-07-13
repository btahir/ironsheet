# Security Policy

## Reporting a Vulnerability

Please do not report security vulnerabilities through public GitHub issues.

Use GitHub private vulnerability reporting:

- Open the repository's **Security** tab.
- Select **Advisories**, then **Report a vulnerability**.
- Start a private draft security advisory with the details below.

If the private reporting button is unavailable, do not open a public issue or include exploit details in a public discussion. Contact the repository owner privately first.

Include as much of the following as you can:

- A description of the issue and its impact
- A minimal reproduction (a crafted workbook file or the code that generates one)
- The affected package(s) and version(s)
- Any suggested remediation

We will acknowledge reports within a few business days, keep you informed of progress, and credit reporters in release notes unless you prefer to stay anonymous.

## Scope

Ironsheet parses untrusted workbook input by design, so parsing is in scope:

- Memory-safety or denial-of-service issues while opening untrusted XLSX/XLSM files (ZIP bombs, malformed ZIP structures, pathological XML)
- Path traversal via crafted ZIP entry names or OPC part names
- XML parsing issues (entity expansion, resource exhaustion)
- Any case where mutating a workbook silently corrupts or leaks content from other workbook parts

Out of scope:

- Vulnerabilities in Excel, LibreOffice, Numbers, or other applications that open workbooks produced by Ironsheet
- Issues that require the victim to run attacker-supplied JavaScript/TypeScript
- Macro (VBA) payloads inside workbooks: Ironsheet preserves macro bytes but never executes them

## Supported Versions

Ironsheet is pre-1.0. Security fixes land on `main` and ship in the next release; we do not backport to older 0.x versions.
