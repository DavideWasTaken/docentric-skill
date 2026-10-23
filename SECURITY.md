# Security

The skill's tool (`skills/docentric/scripts/docentric.mjs`) runs locally, makes no network calls and sends no
telemetry. It reads the templates it is pointed at and writes only new files (previews, edited copies, created
templates). The file is generated from `src/` by `npm run build`; CI checks that the committed file matches the
sources, so what you read in `src/` is what runs.

- **File access**: set `DOCENTRIC_ALLOWED_ROOTS` (folders separated by `;` on Windows, `:` elsewhere) to restrict
  every read and write to those folders.
- **External programs**: for PDF/PNG previews it may start LibreOffice (`soffice --headless`), Microsoft Word through
  COM automation (its own `/Automation` instance, never your open documents) or Edge/Chrome in headless mode, each
  with a timeout.
- **Templates contain personal data** (sample data from D365). Treat previews like the templates themselves.

To report a vulnerability, please use GitHub's *Report a vulnerability* (Security tab) rather than a public issue.
