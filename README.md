# OBS Dynamic Source Controller

This Python script allows OBS Studio to dynamically update Text, Image, Video, and Browser sources based on data from an Excel file. You can easily manage and update multiple OBS sources using a structured Excel sheet, making it ideal for automated setups like esports tournaments or live productions.

## Features

- Dynamically update Text and Image sources in OBS.
- Automatically reload sources at user-defined intervals.
- Supports `.xlsm` Excel files for source management.
- Easy setup with dropdown menus for selecting worksheets and column headers.

## Requirements

### Software
- **OBS Studio** with Python Scripting enabled
- **Python 3.10+** installed (with `pandas` and `openpyxl` libraries)
- **Excel file** (.xlsm) for data management

### Python Libraries
You'll need the following Python libraries installed:
```bash
pip install pandas openpyxl
```

## Installation

1. **Download the Script**: Clone or download the repository from GitHub.
2. **Python Setup**:
   - Ensure Python 3.10 or higher is installed.
   - Install the required libraries:
     ```bash
     pip install pandas openpyxl
     ```
3. **OBS Studio**:
   - Ensure OBS is installed with scripting support enabled.

## Script Setup in OBS

1. **Open OBS Studio**.
2. Go to **Tools** > **Scripts**.
3. Click the **+** button and select the downloaded Python script (`obs_dynamic_source_controller.py`).
4. In the script settings, fill in the following options:

   - **Excel File**: Select the `.xlsm` file you’ll use to manage your sources.
   - **Worksheet**: Choose the worksheet from the Excel file containing your source data.
   - **Source Type Column**: Select the column in the Excel file that contains the source types (Text & Image).
   - **Source Name Column**: Select the column containing the names of the sources in OBS.
   - **Value Column**: Select the column containing the values (text, image file paths, video file paths, or URLs).

5. Choose the **Reload Interval** to set how often the script updates OBS from the Excel file.

6. Press the **Start/Stop Script** button to start or stop the script.

## How to Use

1. **Prepare the Excel File**:
   - Your Excel file should have the following structure (you can customize the column names):
   
   | Source Type | Source Name | Value                             |
   |-------------|-------------|-----------------------------------|
   | Text        | Title       | Welcome to the Stream             |
   | Image       | Logo        | C:/path/to/logo.png               |

   - The **Source Type** column can have the values: `text` or `image`.
   - The **Source Name** must match the name of the source in OBS.
   - The **Value** column should contain the text for `text` sources, and file paths for `image` sources.

2. **Run the Script in OBS**:
   - The script will automatically update your OBS sources based on the data from the Excel file.
   - You can change the Excel file data, and the script will refresh the sources at the interval you set.

## Example Use Cases

- **Esports Tournaments**: Dynamically update player names, team logos, and stats from an Excel file without manually adjusting OBS.
- **Live Productions**: Quickly switch between video intros, sponsor images, or text overlays using data managed in Excel.
- **Automated Streaming**: Automate browser overlays like live chat or scoreboards based on changing data.

## Troubleshooting

- **Sources not updating**: Ensure that the source names in your Excel file match the exact names of the sources in OBS.
  ![Refresh](https://github.com/user-attachments/assets/990f4686-622e-4511-b3c0-6361a5e69787)

- **File paths/URLs not working**: Double-check that the file paths for images/videos are correct and accessible. Ensure URLs for browser sources are valid.
- **Drop down properties not loading**: Press the refresh button after loading the Excel file and worksheet

## Contributing

Feel free to fork this repository, improve the code, and create a pull request if you'd like to contribute. Any feedback is welcome!

## To-Do

- Add audio file source
- Add browser source
- Add video source

## License

This project is licensed under the MIT License. See the LICENSE file for more details.
