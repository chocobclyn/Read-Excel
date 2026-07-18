# Read-Excel — Dynamic Source Controller

This Python script allows OBS Studio to dynamically update Text, Image, Media, Browser, and Color sources — plus source visibility and the active scene — based on data from an Excel/CSV file or a Google Sheet. You can easily manage and update multiple OBS sources using a well structured sheet, making it ideal for automated setups like esports tournaments or live productions.



## Features

- Dynamically update Text, Image, Media, Browser, and Color sources in OBS.
- Show/hide sources, move/scale/rotate them, change bounding-box style, and switch scenes from the sheet.
- Automatically reload sources at a user-defined interval.
- Sheet and column settings are dropdowns read from your actual file — no manual typing, no spelling mistakes.
- Supported data sources: `.xlsx` `.xlsm` `.csv` files, and Google Sheets (link-shared or private via API).
- Auto-installs its own Python dependencies on first load.
- Status line in the script settings shows what happened each cycle (updated / skipped / errors).
- Blank rows or cells are skipped safely — they never touch OBS.
- Editing and saving the file in Excel while the script runs no longer triggers "file in use" errors.
- Works on Windows, macOS, and Linux.
- Sources list generator: creates an Excel file showing which sources can be controlled by the script.

## Requirements

- **OBS Studio** with Python scripting enabled (a Python 3 install configured under Tools > Scripts > Python Settings).
- That's it — the script installs its own libraries (`openpyxl`; `gspread`/`google-auth` only if you use the Google Sheets API route). Manual fallback:
  ```bash
  python -m pip install -r requirements.txt
  ```

## Script Setup in OBS

1. **Open OBS Studio**.
2. Go to **Tools** > **Scripts**.
3. Click the **+** button and select the downloaded Python script (`read-excel.py`). DO NOT LOAD `template_generator.py` INTO SCRIPTS!!!
4. On first load the script installs its dependencies — watch the status line, then reload the script (Scripts > Reload) when it says so.
5. In the script settings, fill in the following options:

   - **Data Source**: Excel, CSV, or Google Sheet.
   - **File Path**: The database file you'll use to manage your sources (Excel/CSV only).
   - **Sheet Name/Label**: Dropdown of the worksheets found in the selected file — pick the one with your source data.
   - **Google Sheet URL**: For the Google Sheet option — paste the sheet's share link.
   - **Service Account JSON** (optional): Only for *private* Google Sheets — see below.
   - **Source Type Column** / **Source Name Column** / **Value Column**: Dropdowns auto-filled with the selected sheet's column headers — no typing needed. (For Google Sheets there's no local file to scan, so type the header names into the same boxes.)
   - **Reload Interval**: How often (in seconds) the script re-reads the data.

6. Press the **Start/Stop Script** button to start or stop the script. Use **Refresh Status** to see the latest result.

## How to Use

1. **Prepare the sheet** (you can customize the column names):

   | Source Type | Source Name | Value                             | Visible | Pos X | Pos Y | Scale | Rotation | Bounds Type |
   |-------------|-------------|-----------------------------------|---------|-------|-------|-------|----------|-------------|
   | Text        | Title       | Welcome to the Stream             |         |       |       |       |          |             |
   | Image       | Logo        | C:/path/to/logo.png               | TRUE    | 100   | 50    | 0.5   |          |             |
   | Media       | VideoTrack  | C:/path/to/video.mp4              |         |       |       |       |          |             |
   | Browser     | VDO Ninja 1 | https://www.path.to/browser       |         |       |       |       |          |             |
   | Color       | Backdrop    | #FF8800                           |         |       |       |       |          |             |
   |             | Sponsor Bar |                                   | FALSE   |       |       |       | 45       |             |
   | Scene       | Intermission| FALSE                             |         |       |       |       |          |             |

   - **Source Type** values: `text` `image` `media` `browser` `color` `scene` (case-insensitive).
   - **Source Name** must match the EXACT name of the source in OBS (or the scene name for `scene` rows).
   - **Value** semantics:
     - `text`: the text to display; `image`/`media`: a local file path; `browser`: a URL.
     - `color`: a hex color like `#FF8800` (for Color sources).
     - `scene`: set to `TRUE` to switch to that scene. Keep several scene rows in the sheet and flip one to TRUE to change scenes; the script won't re-trigger the transition while it's already live.
   - **Setting columns** (all optional — add only the ones you want): `Visible`, `Pos X`, `Pos Y`, `Scale`, `Rotation`, `Bounds Type`. Column names are matched case-insensitively; a **blank cell leaves that setting untouched**. They apply to the source in every scene containing it. A row can be settings-only (blank Type/Value) — see the Sponsor Bar row above.
     - `Visible`: `TRUE`/`FALSE` (also `yes`/`no`, `on`/`off`, `show`, `1`/`0`).
     - `Pos X` / `Pos Y`: position in pixels (either one alone works — the other axis is kept).
     - `Scale`: uniform scale factor (`0.5` = half size, `2` = double).
     - `Rotation`: degrees.
     - `Bounds Type`: OBS's bounding-box style as a number — `0` none, `1` stretch, `2` scale to inner, `3` scale to outer, `4` scale to width, `5` scale to height, `6` max size only. Note: styles only have a visible effect if the item's bounds size is set (Edit Transform in OBS) — this column switches the style, it doesn't invent a box size.
   - Rows with a blank Source Name — or nothing to apply — are skipped (counted in the status line), so partial rows are safe.
   - **Breaking change** from earlier versions: the `visibility` Source Type is gone — move those rows to the `Visible` column.

2. **Run the script in OBS**:
   - Edit the sheet, save (CTRL+S), and the sources refresh within one reload interval. You can keep the file open in Excel the whole time.

3. **Google Sheets**:
   - **Easiest (no credentials)**: Share the sheet as "Anyone with the link can view", pick *Google Sheet* as the Data Source, and paste the link into *Google Sheet URL*.
   - **Private sheets (API)**: Create a Google Cloud service account with the Sheets API enabled, download its JSON key, share the sheet with the service account's email, and set *Service Account JSON* to the key file. The script installs `gspread` automatically the first time you use this.

4. **Source List Generator** (optional):
   - Click **Generate Source Template** to scan your scene collection.
   - The file is saved to your Downloads folder as `source_template.xlsx`, listing every controllable source with a `CHANGE_ME` placeholder value.
   - Use it as the starting point for your control sheet.

## Example Use Cases

- **Esports Tournaments**: Dynamically update player names, team logos, and stats from an Excel file without manually adjusting OBS.
- **Live Productions**: Quickly switch between sponsor images, text overlays, or full scenes using data managed in a sheet — even remotely via Google Sheets.

## Troubleshooting

- **Sources not updating**: Check the status line (press *Refresh Status*). Ensure the source names in your sheet match the EXACT names of the sources in OBS, and the three column settings match your sheet's headers.
  ![Refresh](https://github.com/user-attachments/assets/990f4686-622e-4511-b3c0-6361a5e69787)

- **File paths/URLs not working**: Double-check that the file paths or URLs are correct and accessible.

- **Dependency install failed**: Run the command shown in the status line manually with the same Python that OBS uses, then reload the script.

## Contributing

Feel free to fork this repository, improve the code, and create a pull request if you'd like to contribute. Any feedback is welcome!

## License

This project is licensed under the MIT License. See the LICENSE file for more details.
