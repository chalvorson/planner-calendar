# Planner Calendar Generator

Generate a printable, year-at-a-glance HTML calendar from a Microsoft Planner Excel export. Tasks are displayed across their start and due dates, with each unique task name assigned a distinct color.

## Features

- Reads task data (`Task Name`, `Start date`, `Due date`) directly from a Planner Excel export file (`.xlsx`).
- Generates a single HTML file displaying a full year calendar (12 months).
- Renders tasks as blocks spanning the days from their start date to their due date (inclusive) within the target year.
- Assigns a unique, visually distinct background color to each task name for easy identification.
- Uses hashing to ensure somewhat stable color assignments between runs.
- Optionally uses task labels or buckets for color assignment.
- Includes CSS optimized for printing on large format paper (e.g., 11" x 17") in landscape mode.
- Allows customization of task color saturation and lightness via command-line arguments.
- Automatically determines the calendar year based on the earliest task start date, or allows specifying a year manually.
- Optionally generates a calendar for a specific month instead of the full year.

## Unsupported

- **Repeating tasks:** The script currently cannot support tasks that repeat due to the way the data is structured in the Excel export.

## Prerequisites

- **Python:** Version 3.12 or higher recommended (check with `python --version`).
- **uv:** A fast Python package installer and resolver. If you don't have it, follow the [official uv installation guide](https://astral.sh/docs/uv#installation).

## Installation

1.  **Clone the repository (or download the files):**
    ```bash
    git clone <your-repository-url> 
    cd planner-calendar
    ```
    *(If you just downloaded the files, navigate to the directory containing `main.py` and `pyproject.toml`)*

2.  **Create a virtual environment (Recommended):**
    Using `uv`:
    ```bash
    uv venv
    ```
    Activate the environment:
    - Windows (Command Prompt): `.venv\Scripts\activate`
    - Windows (PowerShell): `.venv\Scripts\Activate.ps1` 
    - macOS/Linux: `source .venv/bin/activate`

3.  **Install dependencies:** (Optional: uv will handle this automatically)
    Using `uv`:
    ```bash
    uv pip install -r requirements.txt 
    ```
    *(Note: This project currently uses uv and `pyproject.toml` for dependency management.  The `requirements.txt` file is included for compatibility with older tools.)*

## Usage

Run the script from your activated virtual environment, providing the path to your Planner Excel export file as the main argument.

```bash
uv run main.py path/to/your/PlannerExport.xlsx [OPTIONS]
```

Or to run the GUI version:
```bash
uv run gui.py
```

**Arguments:**

- `excel_file`: (Required) The path to the Microsoft Planner Excel export file (`.xlsx`). This file **must** contain a sheet named `Tasks` with columns `Task Name`, `Start date`, and `Due date`.

**Options:**

- `-o OUTPUT`, `--output OUTPUT`: Specifies the name for the generated HTML file.
  (Default: `planner_calendar.html`)
- `-y YEAR`, `--year YEAR`: Forces the calendar to be generated for a specific `YEAR` (e.g., `2024`). If omitted, the script uses the year of the earliest task start date found in the file.
- `--month MONTH`: Generates a calendar for a specific month (e.g., `1` for January, `2` for February, etc.). If omitted, the script generates a full year calendar.
- `--no-wrap-text`: Prevents task names from wrapping to the next line if they exceed the available space.
- `--color-saturation SATURATION`: Sets the saturation level for generated task colors (a float between 0.0 and 1.0). Higher values mean more intense colors.
  (Default: `0.7`)
- `--color-lightness LIGHTNESS`: Sets the lightness level for generated task colors (a float between 0.0 and 1.0). Higher values mean lighter colors (better contrast with black text).
  (Default: `0.85`)
- `--color-by-bucket`: Use the task bucket for color generation.
- `--color-by-label`: Use the task label for color generation.
- `--prefix-labels`: Prefix task names with their labels.
- `--alternate-colors`: Use alternate colors algorithm for tasks.
- `-h`, `--help`: Show the help message and exit.

**Examples:**

- **Basic Usage (output to `planner_calendar.html`, auto-detect year):**
  ```bash
  uv run main.py "C:\Downloads\My Planner Export.xlsx"
  ```

- **Specify Output File and Year:**
  ```bash
  uv run main.py "My Planner Export.xlsx" -o TeamCalendar2024.html -y 2024
  ```

- **Adjust Task Colors (Lighter, Less Saturated):**
  ```bash
  uv run main.py tasks.xlsx --color-lightness 0.9 --color-saturation 0.5
  ```

- **Prevent Task Text Wrapping:**
  ```bash
  uv run main.py tasks.xlsx --no-wrap-text
  ```

- **Use Task Label for Color Generation:** (can't be used with `--color-by-bucket`)
  ```bash
  uv run main.py tasks.xlsx --color-by-label
  ```

- **Use Task Bucket for Color Generation:** (can't be used with `--color-by-label`)
  ```bash
  uv run main.py tasks.xlsx --color-by-bucket
  ```

- **Prefix Task Names with Labels:**
  ```bash
  uv run main.py projects.xlsx -y 2025 --color-lightness 0.9 --color-by-bucket --prefix-labels
  ```

- **Use Alternate Colors Algorithm:**
  ```bash
  uv run main.py tasks.xlsx --alternate-colors
  ```

## Output

The script generates an HTML file (e.g., `planner_calendar.html`) in the current directory (or the path specified by `-o`).

- Open this file in a web browser (like Chrome, Firefox, Edge).
- To print, use your browser's print function. Ensure settings are configured for:
    - **Layout:** Landscape
    - **Paper Size:** Tabloid (11" x 17") or Ledger (17" x 11")
    - **Margins:** Adjust as needed (default CSS sets 0.5 inches).
    - **Scale:** Usually "Fit to printable area" or 100% works well.
    - **Options:** Enable "Print backgrounds" or "Background graphics" to ensure task colors are printed.

## Notes

- Ensure your Planner export file has the correct sheet name (`Tasks`) and the required columns (`Task Name`, `Start date`, `Due date`).
- Date parsing errors are handled by skipping tasks with invalid dates.
- For tasks with a completed date, the completed date will override the due date.
- Tasks with no due date will be skipped.
- Tasks with no start date are assumed to start and end on the due date.
- If no valid start dates are found and no year is specified with `-y`, the script defaults to the current year and prints a warning.