# Automatizimi i Konvertimeve Offline

Aplikacion lokal (Streamlit) që merr raportin e Track360 dhe gjeneron 7 raportet e konvertimeve offline për Sport, Bingo, dhe UK casino (Google + Bing).

## Instalimi

Kërkohet Python 3.10+.

```bash
cd auto_conversions
pip install -r requirements.txt
```

## Si të niset

```bash
streamlit run app.py
```

Aplikacioni hapet në shfletues te `http://localhost:8501`.

## Struktura

```
auto_conversions/
├── app.py              # Streamlit UI
├── processor.py        # Logjika kryesore
├── config.py           # Rregullat e raporteve (site_id, kanale, etiketa)
├── brand_map.csv       # Hartë: emri në Track360 → emri i pastër + kategoria
├── templates/          # 7 template-et që përdoren si bazë
└── requirements.txt
```

## Si shtohen brand-e të reja

Modifiko `brand_map.csv`. Tre kolona:
- `report_brand_name` — emri ekzakt që del në kolonën `TrackingParams - BrandName` të Track360
- `clean_name` — emri i pastër (lowercase) që do të shfaqet në raport (p.sh. `betfredcasino`)
- `category` — `casino`, `bingo`, ose `sport`

Nëse një brand mungon në CSV, aplikacioni e përpunon me rregullin e parazgjedhur:
lowercase + heqje e prapashtesave si `-UK-Casino`, `-Bingo-GB`, etj.
Të gjitha brand-et e tilla shfaqen në një paralajmërim në UI.

## Si shtohet një vertical i ri

Shto një hyrje në listën `OUTPUTS` te `config.py` me:
- `template` (skedari në `templates/`)
- `site_ids`
- `channel` (`google` ose `bing`)
- etiketat e signup/ftd/casino sipas template-it

Asnjë ndryshim në kod nuk nevojitet.

## Rregullat aktuale

| Output | SiteId | Channel | Brand-e casino |
|---|---|---|---|
| sport_google | 24 | gclid + gbraid | bashkohen në `offline - ftdcasino` |
| sport_bing | 24 | msclkid | bashkohen në `offline - ftdcasino` |
| bingo_google | 568 | gclid + gbraid | bashkohen në `offline - ftdscasino` |
| bingo_bing | 568 | msclkid | bashkohen në `offline - ftdcasino` |
| uk1_google | 25 | gclid + gbraid | trajtim me 2 rreshta |
| uk2_bing | 39 | msclkid | trajtim me 2 rreshta |
| uk3_bing | 25 | msclkid | trajtim me 2 rreshta |

- Signup time: **23:55**, FTD time: **23:58**
- Cid `m_...` për Bing → prefiksi `m_` hiqet
- Cid = `na` → rreshti përjashtohet
- Revenue rrumbullakohet në numër të plotë (USD)

---

# Offline Conversions Automation

Local Streamlit app that takes the Track360 report and generates 7 offline conversion reports for Sport, Bingo, and UK casino (Google + Bing).

## Installation

Requires Python 3.10+.

```bash
cd auto_conversions
pip install -r requirements.txt
```

## How to run

```bash
streamlit run app.py
```

The app opens in the browser at `http://localhost:8501`.

## Structure

```
auto_conversions/
├── app.py              # Streamlit UI
├── processor.py        # Core logic
├── config.py           # Report rules (site_id, channels, labels)
├── brand_map.csv       # Map: Track360 name → clean name + category
├── templates/          # 7 templates used as the base
└── requirements.txt
```

## How to add new brands

Edit `brand_map.csv`. Three columns:
- `report_brand_name` — the exact name that appears in the `TrackingParams - BrandName` column from Track360
- `clean_name` — the clean (lowercase) name that should show up in the report (e.g. `betfredcasino`)
- `category` — `casino`, `bingo`, or `sport`

If a brand is missing from the CSV, the app processes it with the default rule:
lowercase + removal of suffixes like `-UK-Casino`, `-Bingo-GB`, etc.
All such brands are flagged in a warning in the UI.

## How to add a new vertical

Add an entry to the `OUTPUTS` list in `config.py` with:
- `template` (the file in `templates/`)
- `site_ids`
- `channel` (`google` or `bing`)
- the signup/ftd/casino labels matching the template

No code changes needed.

## Current rules

| Output | SiteId | Channel | Casino brands |
|---|---|---|---|
| sport_google | 24 | gclid + gbraid | collapsed into `offline - ftdcasino` |
| sport_bing | 24 | msclkid | collapsed into `offline - ftdcasino` |
| bingo_google | 568 | gclid + gbraid | collapsed into `offline - ftdscasino` |
| bingo_bing | 568 | msclkid | collapsed into `offline - ftdcasino` |
| uk1_google | 25 | gclid + gbraid | 2-row treatment |
| uk2_bing | 39 | msclkid | 2-row treatment |
| uk3_bing | 25 | msclkid | 2-row treatment |

- Signup time: **23:55**, FTD time: **23:58**
- Cid `m_...` for Bing → `m_` prefix is stripped
- Cid = `na` → row is excluded
- Revenue is rounded to a whole number (USD)