# tax-helper

Single-file browser tool that reads Nordnet exports, classifies each security under Danish tax rules and outputs the amounts and field numbers to enter on the forskudsopgørelse (preliminary tax return) at skat.dk. Everything runs client-side; no storage.

**Live app:** https://larspedersen.github.io/tax-helper/

## Files

- `index.html` – the app, one file, no build step
- `DEVELOPING.md` – tax rules implemented, input formats, code map, release procedure, open items
- `build-abis.py` – embeds a new ABIS list (skat.dk xlsx) into `index.html`

## Usage

Open the live page and follow the step-by-step guide at the top. You need Nordnet's holdings and transaction CSV exports and the Skatteoplysninger PDF from TastSelv. Files are read in the browser only.

## Maintenance

- Bump `APP_VERSION` in `index.html` on every change; open pages show a reload banner when the hosted version differs.
- New ABIS list: `python3 build-abis.py <xlsx>` and commit.
- Never commit real exports or tax prints; `.gitignore` blocks csv, pdf and xlsx.
