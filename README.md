# DateTime Picker for Google Sheets™

Google Sheets™ has a built-in date picker, but no way to pick a **time**. DateTime Picker is a free, open-source add-on that lets you pick a date and a time together, in any time zone, and drop it into the selected cells as a real date value.

## Features

- Pick date and time in one dialog
- Inserts ISO 8601 text by default, e.g. `2026-11-12T16:00:00.000Z`, or a formatted date value if you prefer
- Any time zone (GMT by default), with the live UTC offset shown
- Optional "Convert to GMT" so every cell is stored in GMT
- Works on a single cell or a whole selected range
- "Insert current date and time" in one click
- Per-spreadsheet settings: default zone, `dd/mm/yyyy`, `mm/dd/yyyy` or `yyyy-mm-dd`, 12 or 24-hour clock
- Cells stay real dates, so sorting, filters and formulas keep working
- No data leaves your spreadsheet. No tracking, no external servers.

## Install

From the Google Workspace Marketplace (link added after approval), or install it yourself from source below.

- Website: https://mustkeemk.com/datetime-picker
- Privacy policy: https://mustkeemk.com/datetime-picker/privacy
- Terms of service: https://mustkeemk.com/datetime-picker/terms
- Support: https://mustkeemk.com/datetime-picker/support

## Usage

1. Select one or more cells.
2. **Extensions > DateTime Picker > Pick date and time**.
3. Choose the date, time and zone, then click **Insert**.

The cell shows exactly what you picked, for example `12/11/2026 04:00 PM GMT`. Change the defaults in **Extensions > DateTime Picker > Settings**.

## Run from source

Using [clasp](https://github.com/google/clasp):

```bash
npm install -g @google/clasp
clasp login
clasp create --type sheets --title "DateTime Picker" --rootDir src
clasp push
clasp open
```

Or copy each file in `src/` into **Extensions > Apps Script** of any spreadsheet.

## Permissions

| Scope | Why |
|---|---|
| `spreadsheets.currentonly` | Write the picked value into the spreadsheet you are using, nothing else |
| `script.container.ui` | Show the menu and the picker dialog |

## Privacy and data

- The add-on requests only two OAuth scopes: `spreadsheets.currentonly` (read the selected cell to prefill the picker, write the value you pick, and save settings in the spreadsheet's document properties) and `script.container.ui` (show the menu and dialogs).
- It makes no external network calls and has no servers, analytics or tracking.
- It uses no third-party AI or machine learning services, and no user data is used to train AI models.
- Its use of information received from Google APIs adheres to the [Google API Services User Data Policy](https://developers.google.com/terms/api-services-user-data-policy), including the Limited Use requirements.

Full policy: https://mustkeemk.com/datetime-picker/privacy

## Project layout

```
src/        Apps Script source (Code.gs, dialogs, manifest)
docs/       Legacy GitHub Pages site (canonical site: https://mustkeemk.com/datetime-picker)
```

## Contributing

Issues and pull requests are welcome.

## License

[MIT](LICENSE)

## Support

Email mo.mustkem@gmail.com or see https://mustkeemk.com/datetime-picker/support

---

Google Sheets™ and Google Workspace™ are trademarks of Google LLC. This project is not affiliated with Google.
