# Forskud · aktieindkomst — developer reference

Single-file browser tool (`index.html`) that reads Nordnet exports, classifies each
security under Danish tax rules and outputs the amounts and field numbers to enter on the
forskudsopgørelse (preliminary tax return) at skat.dk. Everything runs client-side; no storage.

Status: working for Nordnet + ABIS list + manual Jyske Bank entry. PDF parser verified
against the line structure of a real print (2026-09-27).

## 1. Domain rules the code implements

Tax year defaults to current year. Amounts in DKK.

### Income types → forskud fields

| Type key   | Meaning                                                        | Dividend → felt | Realised → felt | Lager → felt |
|------------|----------------------------------------------------------------|-----------------|-----------------|--------------|
| `stock_dk` | Danish listed share                                            | 501             | 502             | –            |
| `stock_for`| Foreign listed share                                           | 509             | 502             | –            |
| `fund_dk`  | Danish distributing equity fund (udloddende, aktiebaseret)     | 501             | 502             | –            |
| `abis`     | Investeringsselskab on skat.dk ABIS list (positivliste)        | 501 if DK ISIN else 509 | –       | 345          |
| `nonabis`  | Investeringsselskab NOT on ABIS list → kapitalindkomst         | 239             | –               | 239          |
| `bond`     | Bond-based fund, minimumsbeskattet                             | 233             | 330             | –            |
| `exclude`  | Ignore (ASK, pension, errors)                                  |                 |                 |              |

Forskud felt ↔ årsopgørelse rubrik: 501↔61, 505↔62, 509↔63, 502↔66, 345↔345, 239↔38, 330↔30, 233↔31.

### Auto-classification (`classify(p)`)
1. Manual override wins.
2. No ISIN → guess by name (fund-like words → `abis`, else `stock_for`), flag "ISIN missing".
3. ISIN on ABIS list for the tax year → `abis`. The list is embedded (`ABIS_EMBED`); a dropped-in xlsx overrides it. (`BUILTIN_ABIS`, 7 ISINs, is only a fallback if the embed is empty.)
4. DK ISIN → `fund_dk` if name looks like a fund, else `stock_dk`.
5. Foreign ISIN, not on list → `nonabis` if fund-like name, else `stock_for`.

Fund-like name regex: `\b(kl|indeks|index|invest|fond|fund|fds|afd|sparindex|etf|ucits|acc|dis|udb|akk|sicav|portefølje|portfolio)\b|a-dis|a-eur|a-usd`

### Lager (mark-to-market) gain for `abis`/`nonabis`
```
gain = (value_now + sale_proceeds_this_year) − (value_31Dec_prev_year + purchases_this_year)
```
- If all units were bought this year, baseline = 0 (purchase price already in `purchases`).
- Else baseline must come from skat.dk Skatteoplysninger (kursværdi ultimo) — PDF or manual input.
- Positions that appear only in transactions (sold out) still contribute dividends/realised.

### Buffer, threshold, tax estimate
- Buffer = `pct × Σ value_now of lager positions`, added to felt 345 (or 239 if no ABIS positions). Default 10%.
- Aktieindkomst = 501 + 502 + 509 + 345. Threshold 2026: 79.400 kr. (2025: 67.500; table in `THRESHOLDS`), doubled if married.
- Tax = 27% up to threshold, 42% above; negative income → 27% negative tax. Withheld dividend tax (UDBYTTESKAT rows + manual) is netted off.

### ABIS list eligibility for a tax year
`abisEligible()`: current skat.dk layout has `Registrerede år` per row (e.g. `2021,2022,2026`; Excel may store `2021,2022` as the number `2021.2022`) → on list if taxYear is among those years (union over all rows/sheets of the ISIN). Older layout fallback: `Første registreringsår ≤ taxYear` and (`Fjernet` empty or year(Fjernet) > taxYear).
Note (fixed 2026-09-27): the old code looked for `Første`/`Fjernet` columns that the real file does not have, so every ISIN on any sheet 2021–2026 counted as ABIS (6.073 instead of 5.298 for 2026). Example: Jupiter Financial Innovation LU0262307720 is registered 2021–2022 only → `nonabis` in 2026.

## 2. Input file formats

### Nordnet holdings CSV (`Aktietabel, Nordnet kontonummer N, D.M.YYYY.csv`, `Fondstabel, …` – one per account and tab)
- UTF-16 LE with BOM, tab-separated, decimal comma, no thousands separators. Header may contain NBSP.
- No ISIN column. ISIN resolved via: manual override → `KNOWN_ISIN` map → transaction file name match.
- Aktietabel columns: `Navn, Valuta, Antal, GAK, I dag %, Seneste kurs, Belåningsværdi DKK, Værdi, Værdi DKK, Ureal.afkast %, Afkast DKK`  → cost = Værdi DKK − Afkast DKK
- Fondstabel columns: `Navn, Valuta, Antal, GAK, 1 dag %, Indre værdi, Belåningsværdi DKK, Anskaffelsessum DKK, Værdi DKK, Ureal.afkast %, Afkast DKK`
- Depot type (`autoDepot`, re-run in `rebuild()` for every file not overridden by the user): filename `aktiesparekonto`/`ask` → ASK; `pension|rate|alder` → pension. Otherwise linked by account number: `Depot` values of all `AFKASTSKAT ASK` rows form `S.askAccounts`; a holdings file whose name holds that account (`…kontonummer 12345678…`, `accountOf`) → ASK; a tx file whose Depot values are all ASK accounts → ASK. Else free depot. User can override per file (`depotManual`).

### Nordnet transactions CSV (`transactions-and-notes-export*.csv`)
- Same encoding. Columns incl. `Bogføringsdag, Handelsdag, Valørdag, Depot, Transaktionstype, Værdipapirer, ISIN, Antal, Kurs, Beløb, Valuta.1, Indkøbsværdi, Resultat, Valuta.3, Totalt antal, Vekslingskurs, Transaktionstekst`.
- Types used: `KØBT` (purchases, Beløb negative DKK), `SOLGT` (sales; realised = Resultat × Vekslingskurs if foreign), `UDBYTTE` (gross, positive), `UDBYTTESKAT` (withheld, negative), `AFKASTSKAT ASK` (marks that row's `Depot` account as ASK). Rows of ASK accounts are dropped per row from auto-detected free tx files, so a multi-account export works; a tx file forced to free keeps all rows.
- Only rows with Handelsdag in the tax year count.

### ABIS list xlsx (skat.dk)
- Download: skat.dk → Erhverv → eKapital → Værdipapirer → "Beviser og aktier i investeringsforeninger og selskaber (IFPA)" → "Liste over aktiebaserede investeringsselskaber". Sept 2026 file: `august-2026-abis-liste-2021-2026.xlsx` (~5.400 rows).
- Sheets `Forside`, `2026` … `2021`. Columns: `Skattemæssigt hjemsted, ISIN-kode, Navn andelsklasse, LEI-kode, CVR/SE/TIN, Navn afdeling, TIN, Navn, Registrerede år, Ikke registrerede år`. Header row located by "isin".
- **Built in**: `python3 build-abis.py [xlsx] [--label "august 2026"]` (python3 + openpyxl) applies the same header detection and rewrites the block between `/* ABIS_EMBED_START */` and `/* ABIS_EMBED_END */` in the HTML: `const ABIS_EMBED = {label, source, groups:[["2021,2022,2026","ISIN,ISIN,…"],…]}` (one group per set of registered years, ~78 KB). August 2026 list: 6.073 ISIN in 34 groups; 5.298 on the list for 2026, 4.833 for 2025. At startup and on tax-year change `loadBuiltinAbis()` builds `S.abis` for `S.taxYear` (`S.abisSrc="builtin"`) unless an xlsx was dropped in (`S.abisSrc="file"`, parsed by `ingestAbis` with SheetJS; computed for the tax year at upload time).

### Skatteoplysninger PDF (skat.dk TastSelv → Årsopgørelse → Skatteoplysninger, R75 print)
Text rebuilt from pdf.js items grouped by y (verified against a real print 2026-09-27). Relevant blocks:
```
Beholdning (optaget til handel)                         ← single shares
Regnr/CVRnr Depotnummer Depot Ejerst. Antal aktier Aktiekurs Kursværdi R.dato
Nordnet, filial af nordnet bank ab
31/12-25                                                ← R.dato on its own line
<regnr 8> <depotnr 15–17> [ejerst.] <antal> <kurs,6dec> <kursværdi>   ← ejerst. token seen in ASK section
Aktie
Isin-kode Papirnavn Landenavn
<ISIN> <name> <country>
AS-ident Ejerandel Valuta Valutakurs

Udlodning / tilskrivning                                ← ETF / investeringsselskab
Regnr/CVRnr Konto/depotnummer Ejerst. Kursværdi Indestående Udbytte R.dato
Nordnet, filial af nordnet bank ab
31/12-25
<regnr> <depotnr> <kursværdi> [indestående] [udbytte]
Investeringsselskab
Isinkode Papirnavn Landenavn
<ISIN> <name> <country>
AS-ident Antal beviser/aktier
<antal>
Diverse øvrige oplysninger
Lagerprincip
Aktiebaseret|Obligationsbaseret, Optaget til handel, SEL § 3.1.19
```
Sections: "Aktieoplysninger", "Aktiesparekonto" (skip), "Investeringsforeninger og -selskaber". Numbers use `.` thousands and `,` decimals (`12.345`, `1.234,50`).
- Page headers/footers can split a block anywhere; lines matching `PDF_NOISE` (`27/09/2026, 16:12`, `TastSelv - Skatteoplysninger`, `…tastselv.skat.dk…`, page `12/18`) are removed first.
- Data row = first line in the block starting with two long numbers; drop regnr + depotnr. Beholdning: kursværdi = last number, antal = number before the 6-decimal price. Udlodning: kursværdi = first remaining number; antal = first bare number after "Antal beviser/aktier".
- Same ISIN twice outside ASK is summed. Only Beholdning/Udlodning blocks are read (Gebyr/Køb/Salg/Udbytte ignored).
- `PDF_NAME`/`PDF_KIND` keep papirnavn (without ISIN and country) and block kind per ISIN, all sections. `rebuild()` uses `pdfIsinByName` to give a free-depot holding without ISIN the PDF ISIN when name tokens match (lowercase, no punctuation, generic words and <3-char tokens dropped; all tokens of the shorter list, min. 2, in the longer; exactly one hit; ISIN not used by another holding). Marked `isinSrc="pdf"`, rule column shows `rIsinPdf`.

## 3. Verified reference data (public identifiers)

| Name (Nordnet)                                                 | ISIN         | ABIS 2026 |
|----------------------------------------------------------------|--------------|-----------|
| iShares S&P 500 Information Technology Sector UCITS ETF USD (Acc) | IE00B3WJKG14 | yes |
| iShares S&P 500 Communication Sector UCITS ETF USD (Acc)       | IE00BDDRF478 | yes |
| iShares Digital Security UCITS ETF USD (Acc)                   | IE00BG0J4C88 | yes |
| iShares Edge MSCI EM Value Factor UCITS ETF USD (Acc)          | IE00BG0SKF03 | yes |
| Amundi MSCI Korea ETF Acc                                      | LU1900066975 | yes |
| Fidelity Fds- Asian Special Sits A Udb (tx name: Fidelity Asian Special Sits A-Dis-USD) | LU0054237671 | yes (2026 only; not 2020–2025) |
| Fidelity Global Technology A-Dis-EUR                           | LU0099574567 | yes (2026 only; not 2020–2025) |
| Sparindex INDEX OMX C25 KL                                     | DK0060442556 | n/a (Danish udloddende fund) |
| Nordnet Global Indeks                                          | IE00BMTD2K75 | check against list |

Test expectations (synthetic, same shapes as the real exports): four purchases of one ETF in July for 20.000 kr. in total, worth 25.000 kr. now → lager gain 5.000 kr. (all bought this year, baseline 0); a Danish fund dividend of 100,00 kr. gross with 27,00 kr. withheld; an ETF held all year with baseline 50.000 kr. and value 60.000 kr. now → 10.000 kr.

## 4. Code map (`index.html`)
- `APP_VERSION` (top of script) — shown in the footer; `checkForUpdate()` fetches `location.href` (no-store) on load and every 5 min and shows `#updBanner` when the hosted file has another version (skipped for `file:`, errors ignored).
- `I18N` da/en dictionaries, `t()`, `applyLang()` — all UI strings incl. guide HTML (`{abisLabel}` filled from `ABIS_EMBED.label`).
- `ABIS_EMBED` (generated block), `abisYears`, `abisEligible`, `loadBuiltinAbis` — embedded ABIS list; `ingestAbis` uses the same `abisEligible`.
- `decodeBuffer`, `parseCSV`, `num`, `get` — encoding/CSV helpers.
- `handleFiles` → `ingestAbis` / `ingestPdf` / CSV kind detection.
- `parseSkatteoplysninger`, `dkNum`, `PDF_NOISE`, `PDF_ANTAL`, `PDF_NAME`, `PDF_KIND`, `isinByName`/`pdfIsinByName` — PDF parser and name → ISIN match (also used for transaction names).
- `rebuild()` — computes `S.askAccounts`, sets depot per file (`autoDepot`, `accountOf`), then builds `S.holdings` from holdings + transactions.
- `classify`, `fieldsFor`, `lagerGain`, `compute`, `taxCalc`.
- `render*()` — files, status, classification table, baselines, Jyske manual rows, result strip, evidence.
- State in `S`; overrides keyed by normalised name (`nkey`); baselines keyed by ISIN.

External scripts (CDN): SheetJS 0.18.5, pdf.js 3.11.174 (+ worker). Font: IBM Plex Sans (Google Fonts).

## 5. Open items / next steps
1. PDF parser verified against a real Skatteoplysninger print 2026-09-27 (own-line R.dato, footer stripping, token rule). Added name matching PDF papirnavn → ISIN for holdings without ISIN (e.g. "iShares Core S&P 500 ETF USD Acc" → IE00B5BMR087). Cross-check PDF antal (item 4) still open.
2. Jyske Bank: no export sample yet. Panel is manual entry; add parser when a netbank export (CSV/PDF) is available. Likely source: Netbank → Investering → Afkast/Skat.
3. Realised gains use Nordnet's `Resultat` column, not Skat's gennemsnitsmetode. Acceptable for forskud; note in UI.
4. Cross-check PDF `antal` vs holdings `Antal` and warn on mismatch (data exists in `PDF_ANTAL`, not yet displayed).
5. Consider a "copy per-field summary" button for pasting into skat.dk.
6. Possible folder mode for Claude Code: a small Node/Python CLI reusing the same parsing rules, reading a directory of exports and emitting the field table as markdown.
7. Fixed 2026-09-27: ASK holdings file (`Aktietabel, Nordnet kontonummer N, …`) was treated as free depot (positions showed "mangler") because only the tx file was detected via `AFKASTSKAT ASK`. Now linked by account number; baseline table shows "ingen transaktionsfil" instead of "mangler" when no free tx file is loaded.
   Also fixed: the same security in several free-depot holdings files (same ISIN or normalised name) is now summed (units, value, cost) instead of the last file overwriting the first.
   Status panel warns (`wAskForced`) when a file auto-detected as ASK has been manually set to free depot.
8. Fixed 2026-09-27: positions that exist only in transactions without KØBT/SOLGT (e.g. `TILBAGEBETALT FOND GEBYR` rows for funds held in Nordnet's separate Fonde tab) got lager gain = −baseline (fabricated loss). Now `notInHoldings(p)` → no lager gain, not in buffer value, class-table tag `tagNotInHoldings`; status warning `wPdfNotHeld` lists baseline ISINs without a real position (hint: export Fondstabel), plus `wPdfFeeHint` when fee-refund rows exist. Name match also accepts one identical token ≥5 chars (Ørsted ↔ ØRSTED); ADR/ADS never matches an ordinary share.

9. Added 2026-09-27: `APP_VERSION` + update banner (for GitHub Pages hosting), embedded ABIS list (`build-abis.py`, `abisEligible` – also fixed the ABIS year rule, see section 1), Fondstabel handled like Aktietabel (same columns incl. `Anskaffelsessum DKK`), guide rewritten (per-account Aktier/Fonde exports with unchanged filenames, "Alle konti" transactions, full Skatteoplysninger print, built-in ABIS list).
   2026-09-27.2: holdings without ISIN are also token-matched against transaction names (`isinByName`, same rules as the PDF match) before the PDF match, so "Nordnet Teknologi Indeks" (Fondstabel) merges with "Nordnet Teknologi Indeks DKK" (tx, IE00BNNLSN94) instead of becoming a second, sold-out row; `isinSrc="tx"`, rule text `rIsinTx`. Resolution order: override → `KNOWN_ISIN` → exact tx name → token match tx names → token match PDF names.

### Release procedure
- Bump `APP_VERSION` (`YYYY-MM-DD.n`) on every change to the HTML – open pages compare it with the hosted file and ask the user to reload.
- When skat.dk publishes a new ABIS list: `python3 build-abis.py <new.xlsx>` (label from the filename, or `--label`), check the printed counts, bump `APP_VERSION`.

## 6. Privacy
Never commit or store real exports, Skatteoplysninger prints or CPR/account numbers. The tool is
memory-only by design. Sample data above contains only public fund identifiers.
