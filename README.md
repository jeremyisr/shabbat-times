# Shabbat Times

Generate an Excel file containing Jewish holiday (Yom Tov) schedules for building access control systems.

## What it does

Produces an Excel file with start/end datetimes for all days where electricity use is forbidden according to Jewish law. Designed to feed into access control systems.

### Holidays included

- Rosh Hashana (2 days)
- Yom Kippur (1 day)
- Sukkot I & II (first 2 days)
- Shemini Atzeret & Simchat Torah (2 days)
- Pesach I & II and VII & VIII (first and last 2 days)
- Shavuot I & II (2 days)

Consecutive Yom Tov days and adjacent Shabbat are merged into a single block.

**Not included:** Hanukkah, Purim, fast days, memorial days, Chol HaMoed, standalone Shabbat.

### Excel columns

| Column | Description |
|---|---|
| Date et heure de début | Candle lighting time minus safety margin |
| Date et heure de fin | Havdalah time plus safety margin |
| Nom de la fête | Holiday name(s), joined with " + " |

## Configuration

Edit the constants at the top of `generate_holidays.py`:

| Parameter | Default | Description |
|---|---|---|
| `LAT` / `LON` | 48.8534 / 2.4727 | GPS coordinates (Fontenay-sous-Bois) |
| `MARGIN_MINUTES` | 6 | Safety margin in minutes (before start, after end) |
| `TODAY` | 2026-05-28 | Start date |
| `END_YEAR` | 2046 | Last year to generate |

Timezone is set to `Europe/Paris` and handles DST automatically.

## Usage

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install openpyxl requests
python generate_holidays.py
```

## Data source

[Hebcal API](https://www.hebcal.com/hebcal) — diaspora mode, French holiday names.
