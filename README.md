# CCAT v2 – Comprehensive CDR Analysis Tool

Cross-platform Python desktop app. Runs on Windows, macOS, and Linux.

## Quick Start

1. Make sure Python 3.8+ is installed.
2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
3. Launch the app:
   ```
   python ccat.py
   ```

## How to Use

1. **Open CDR** – Click "Open CDR…" and select your `.xlsx` or `.xlsm` CDR workbook.
2. **Analyse** – Click "▶ Analyse". The Summary tab fills with stats, the Contacts tab shows top numbers.
3. **Resolve Towers** – Go to the Cell Towers tab, paste your Google Geolocation API key, and click "Resolve Towers". The app calls the Google API for each unique cell tower and retrieves lat/lng.
4. **Export Excel** – Saves a 3-sheet Excel report (Summary, Top Contacts, All CDR Records).
5. **Export KML** – Saves a `.kml` file ready for Google Earth with tower placemarks and a movement path line.

## Expected CDR Format

The app is calibrated to the CCAT sample CDR schema:
| Col | Field |
|-----|-------|
| 1   | CALLG PARTY NO |
| 2   | CALLD PARTY NO |
| 3   | START DATE |
| 4   | CALL TIME |
| 5   | BILL DURATION |
| 6   | FIRST_CELL_ID (format: prefix--LAC--CellID) |
| 7   | LAST_CELL_ID |
| 8   | CALL DIRECTION (IN___CALL / IN___SMS / OUT_CALL / OUT_SMS) |
| 9   | ESN_or_IMEI_NO |
| 10  | MIN_or_IMSI_NO |
| 14  | BTS_ADDRESS |

Row 1 = headers, Row 2 = metadata (skipped), data starts Row 3.

## Google API Key

Get a key at: https://console.cloud.google.com/
Enable: **Geolocation API**
The key is never stored to disk — enter it each session.

## Dependencies
- openpyxl – Excel read/write
- requests – HTTP calls to Google API (stdlib urllib used as fallback)

## Original Repo
https://github.com/sharad1126/ccat
