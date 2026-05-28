# shabbat-times

## Objectif
Générer un fichier Excel (`fetes_juives_2026_2046.xlsx`) contenant les horaires des fêtes juives où l'utilisation de l'électricité est interdite (Yom Tov uniquement). Destiné à programmer un contrôle d'accès bâtiment.

## Fêtes incluses (Yom Tov uniquement)
- Roch Hachana (2 jours)
- Yom Kippour (1 jour)
- Souccot I + II (2 premiers jours)
- Chemini Atseret + Simhat Torah (2 jours)
- Pessah I + II et VII + VIII (premiers et derniers jours)
- Chavouot I + II (2 jours)

**Exclus** : Hanoukah, Pourim, jeûnes, jours commémoratifs, Hol Hamoed, Shabbat seul.

## Logique de fusion
Les jours consécutifs de Yom Tov sont fusionnés en un seul bloc (une seule ligne Excel). Si un Shabbat est adjacent (vendredi soir → samedi soir collé à un Yom Tov), il est intégré au bloc.

## Colonnes Excel
1. **Date et heure de début** — allumage des bougies (coucher du soleil) MOINS la marge
2. **Date et heure de fin** — havdalah (sortie des étoiles) PLUS la marge
3. **Nom de la fête** — noms concaténés avec " + "

## Paramètres clés (`generate_holidays.py`)
- `LAT`, `LON` : Fontenay-sous-Bois (48.8534, 2.4727)
- `TZ` : Europe/Paris (gère automatiquement heure d'été/hiver)
- `MARGIN_MINUTES` : 6 (marge de sécurité avant/après)
- `TODAY` : date de début de génération
- `END_YEAR` : 2046

## Source de données
API Hebcal (https://www.hebcal.com/hebcal) — diaspora (`i=off`), noms en français (`lg=fr`).

## Exécution
```bash
source .venv/bin/activate
python generate_holidays.py
```

## Dépendances
- Python 3 avec venv dans `.venv/`
- `openpyxl`, `requests` (installés dans le venv)
