import argparse
import calendar
import colorsys
import hashlib
import sys
from datetime import datetime, timedelta

import pandas as pd

# Constant for hue step
HUE_STEP = 29

def generate_calendar_html(tasks_df, no_wrap_text, year, task_colors, prefix_label, month=None):
    """
    Generates a year-at-a-glance HTML calendar or a single month calendar.

    Args:
        tasks_df (pd.DataFrame): DataFrame containing tasks with 'Start date', 'Due date', and 'Task Name'.
        no_wrap_text (bool): Whether to prevent task names from wrapping to the next
                             line if they exceed the available space.
        year (int): The year for which the calendar is generated.
        task_colors (dict): A dictionary mapping task names to their respective color codes.
        prefix_label (bool): Whether to prefix task names with their labels.
        month (int or str, optional): If provided, generates only the calendar for this month.
                                     Can be an integer (1-12) or a month name (e.g., "January").

    Returns:
        str: The generated HTML content as a string.
    """

    # --- Wrap tasks ---
    wrapping = "nowrap" if no_wrap_text else "wrap"

    # --- Task Processing ---
    tasks_by_day = {}  # {date_obj: [task_name1, task_name2, ...]}

    # Ensure date columns are datetime objects, coercing errors
    tasks_df["Start date"] = pd.to_datetime(tasks_df["Start date"], errors="coerce")
    tasks_df["Due date"] = pd.to_datetime(tasks_df["Due date"], errors="coerce")
    tasks_df["Completed Date"] = pd.to_datetime(tasks_df["Completed Date"], errors="coerce")

    # For tasks with a valid Completed Date, set the Due Date to the Completed Date
    tasks_df.loc[tasks_df["Completed Date"].notna(), "Due date"] = tasks_df["Completed Date"]

    # For tasks with a valid due date but invalid start date, set start date to due date
    tasks_df.loc[tasks_df["Start date"].isna(), "Start date"] = tasks_df["Due date"]

    # Drop tasks with invalid dates
    tasks_df = tasks_df.dropna(subset=["Start date", "Due date"])

    for _, task in tasks_df.iterrows():
        start_date = task["Start date"].date()
        due_date = task["Due date"].date()
        task_name = task["Task Name"]

        # Ensure we only process tasks relevant to the target year
        if start_date.year > year or due_date.year < year:
            continue

        current_date = max(
            start_date, datetime(year, 1, 1).date()
        )  # Start from Jan 1st if task started earlier
        end_date = min(due_date, datetime(year, 12, 31).date())  # End at Dec 31st if task ends later

        while current_date <= end_date:
            if current_date not in tasks_by_day:
                tasks_by_day[current_date] = []
            if (
                task_name not in tasks_by_day[current_date]
            ):  # Avoid duplicates on the same day if logic repeats
                tasks_by_day[current_date].append(task_name)
            current_date += timedelta(days=1)

    # --- HTML Generation ---
    # Using a list to build the HTML is more efficient than string concatenation
    
    # Convert month to integer if it's a string
    month_num = None
    if month is not None:
        if isinstance(month, str):
            try:
                # Try to convert month name to number
                month_names = {name.lower(): num for num, name in enumerate(calendar.month_name) if num > 0}
                month_num = month_names.get(month.lower())
                if month_num is None:
                    # Try to parse as a number
                    month_num = int(month)
            except (ValueError, KeyError):
                # If conversion fails, default to None (full year)
                month_num = None
        else:
            # If month is already a number
            month_num = month
    
    # Title based on whether we're showing a single month or the full year
    title_text = f"Yearly Planner Calendar - {year}"
    grid_columns = "repeat(3, 1fr)"  # Default for year view (3 columns)
    
    if month_num is not None:
        month_name = calendar.month_name[month_num]
        title_text = f"{month_name} {year} Calendar"
        grid_columns = "1fr"  # Single column for month view
    
    html_parts = [
        f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Planner Calendar - {year}</title>
    <style>
        body {{ font-family: sans-serif; font-size: 8px; }}
        .year-grid {{ display: grid; grid-template-columns: {grid_columns}; gap: 20px; }}
        .month {{ border: 1px solid #ccc; padding: 5px; }}
        .month-title {{ text-align: center; font-weight: bold; font-size: 12px; margin-bottom: 5px; }}
        .calendar-grid {{ display: grid; grid-template-columns: repeat(7, 1fr); border-collapse: collapse; 
        table-layout: fixed; width: 100%; }}
        .day-header {{ text-align: center; font-weight: bold; background-color: #f0f0f0; font-size: 9px; 
        padding: 2px; border: 1px solid #ddd; }}
        .day {{ border: 1px solid #ddd; vertical-align: top; height: 70px; /* Adjust as needed */ 
        padding: 2px; overflow: hidden; position: relative; }}
        .day.other-month {{ background-color: #f9f9f9; color: #aaa; }}
        .day-number {{ position: absolute; top: 1px; left: 1px; font-weight: bold; font-size: 9px; 
        color: #333; }}
        .tasks {{ margin-top: 12px; /* Space below day number */ line-height: 1.1; }}
        .task {{ display: block; white-space: {wrapping}; overflow: hidden; text-overflow: ellipsis; 
        margin-bottom: 1px; padding: 0 1px; border-radius: 2px; border: 1px solid #ccc; color: #000; 
        /* Ensure text is black */}}
        /* Single month view */
        .single-month .day {{ height: 100px; }}
        .single-month .day-number {{ font-size: 12px; }}
        .single-month .day-header {{ font-size: 12px; }}
        .single-month .task {{ font-size: 9px; }}
        /* Print specific styles */
        @media print {{
            @page {{ size: 17in 11in landscape; margin: 0.5in; }} /* 11x17 Landscape */
            body {{ font-size: 7pt; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
            .year-grid {{ gap: 15px; }}
            .month {{ page-break-inside: avoid; border: 1px solid #aaa; }}
            .day {{ height: 65px; border: 1px solid #ccc; }}
            .day-header {{ font-size: 8pt; padding: 1px; }}
            .day-number {{ font-size: 8pt; }}
            .task {{ border: 1px solid #b0dde4; font-size: 5pt;}}
            /* Hide non-essential elements for print */
            /* Add any elements here you want to hide during printing */
        }}
    </style>
</head>
<body>
    <h1 style="text-align: center;">{title_text}</h1>
    <div class="year-grid{' single-month' if month_num is not None else ''}">
"""
    ]

    cal = calendar.Calendar(firstweekday=calendar.SUNDAY)  # Start weeks on Sunday

    # If a specific month is provided, only generate that month
    month_range = [month_num] if month_num is not None else range(1, 13)

    for current_month in month_range:
        month_name = calendar.month_name[current_month]
        month_weeks = cal.monthdatescalendar(year, current_month)

        html_parts.append('<div class="month">')
        html_parts.append(f'  <div class="month-title">{month_name} {year}</div>')
        html_parts.append('  <div class="calendar-grid">')

        # Day Headers
        for day_abbr in ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]:
            html_parts.append(f'    <div class="day-header">{day_abbr}</div>')

        # Days
        for week in month_weeks:
            for day_date in week:
                day_class = "day"
                if day_date.month != month_num:
                    day_class += " other-month"

                html_parts.append(f'    <div class="{day_class}">')
                html_parts.append(f'      <div class="day-number">{day_date.day}</div>')
                html_parts.append('      <div class="tasks">')

                # Add tasks for this day
                tasks_today = tasks_by_day.get(day_date, [])
                for task_name in tasks_today:
                    # Basic HTML escaping for task name in title attribute
                    bg_color = task_colors.get(task_name, "#f0f0f0")  # Get color, fallback to light gray

                    if prefix_label:
                        labels = tasks_df.loc[tasks_df["Task Name"] == task_name, "Labels"]
                        labels = "" if labels.empty or labels.isna().all() else labels.values[0]
                        labeled_task_name = f"{labels}{task_name}" if labels else task_name
                    else:
                        labeled_task_name = task_name

                    escaped_task_name = (
                        labeled_task_name.replace(";", "")
                        .replace("&", "&amp;")
                        .replace("<", "&lt;")
                        .replace(">", "&gt;")
                        .replace('"', "&quot;")
                        .replace("'", "&#39;")
                    )
                    html_parts.append(
                        f'        <span class="task" title="{escaped_task_name}" style="background-color: '
                        f'{bg_color};">{escaped_task_name}</span>'
                    )

                html_parts.append("      </div>")  # close tasks
                html_parts.append("    </div>")  # close day

        html_parts.append("  </div>")  # close calendar-grid
        html_parts.append("</div>")  # close month

    html_parts.append("""
    </div> <!-- close year-grid -->
</body>
</html>
""")
    return "".join(html_parts)


def main():
    parser = argparse.ArgumentParser(
        description="Generate a yearly HTML calendar from a Planner Excel export."
    )
    parser.add_argument("excel_file", help="Path to the Microsoft Planner Excel export file (.xlsx)")
    parser.add_argument(
        "-o",
        "--output",
        default="planner_calendar.html",
        help="Output HTML file name (default: planner_calendar.html)",
    )
    parser.add_argument(
        "-y",
        "--year",
        type=int,
        help="Force a specific year for the calendar. If not provided, uses the earliest start year found.",
    )
    parser.add_argument(
        "-m",
        "--month",
        help=("Generate calendar for a specific month only. "
            "Can be a month name (e.g., 'January') or number (1-12)."),
    )
    parser.add_argument(
        "--no-wrap-text",
        action="store_true",
        help="Do not wrap task names that exceed the available space.",
    )
    parser.add_argument(
        "--color-saturation",
        type=float,
        default=0.7,
        help="Saturation for generated task colors (0.0-1.0, default: 0.7)",
    )
    parser.add_argument(
        "--color-lightness",
        type=float,
        default=0.85,
        help="Lightness for generated task colors (0.0-1.0, default: 0.85)",
    )
    parser.add_argument(
        "-l",
        "--color-by-label",
        action="store_true",
        help="Use the task label for color generation",
    )
    parser.add_argument(
        "-b",
        "--color-by-bucket",
        action="store_true",
        help="Use the task bucket for color generation",
    )
    parser.add_argument(
        "-p",
        "--prefix-labels",
        action="store_true",
        help="Prefix task names with their labels",
    )
    parser.add_argument(
        "--alternate-colors",
        action="store_true",
        help="Use alternate colors for tasks",
    )

    args = parser.parse_args()

    # --- Validate Arguments ---
    if args.color_by_label and args.color_by_bucket:
        print("Error: Cannot use both --color-by-label and --color-by-bucket.", file=sys.stderr)
        sys.exit(1)

    # --- Read Excel File ---
    try:
        # Specify the sheet name explicitly
        tasks_df = pd.read_excel(args.excel_file, sheet_name="Tasks", engine="openpyxl")
    except FileNotFoundError:
        print(f"Error: File not found at {args.excel_file}", file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        if "Worksheet named 'Tasks' not found" in str(e):
            error_msg = f"Error: Worksheet named 'Tasks' not found in {args.excel_file}."
            print(
                f"{error_msg} Please ensure the Planner export format is correct.",
                file=sys.stderr,
            )
        else:
            print(f"Error reading Excel file: {e}", file=sys.stderr)
        sys.exit(1)
    except ImportError:
        print(
            "Error: Missing dependency 'openpyxl'. Please install it: pip install openpyxl", 
            file=sys.stderr
        )
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred while reading the Excel file: {e}", file=sys.stderr)
        sys.exit(1)

    required_columns = ["Task Name", "Start date", "Due date"]
    missing_cols = [col for col in required_columns if col not in tasks_df.columns]
    if missing_cols:
        print(f"Error: Missing required columns in 'Tasks' sheet: {', '.join(missing_cols)}", file=sys.stderr)
        sys.exit(1)

    # --- Generate Task Colors ---
    unique_tasks = sorted(tasks_df["Task Name"].astype(str).unique())  # Ensure consistent order
    task_colors = {}
    saturation = max(0.0, min(1.0, args.color_saturation))
    lightness = max(0.0, min(1.0, args.color_lightness))

    # Use a hash of the task name (or label) for potentially more stable color assignments
    # if the task list changes slightly run-to-run.
    for task_index, task_name in enumerate(unique_tasks):
        hue = None  # Initialize hue
        if args.color_by_label:
            matching_rows = tasks_df.loc[tasks_df["Task Name"] == task_name]
            if not matching_rows.empty:
                label = matching_rows["Labels"].iloc[0]  # Get the first label
                if pd.notna(label):  # Check if label exists and is not NaN
                    try:
                        # Hash the label
                        hash_object = hashlib.md5(str(label).encode())  # Use str() for safety
                        hash_digest = hash_object.hexdigest()
                        if args.alternate_colors:
                            hue = (int(HUE_STEP * task_index) % 360) / 360.0
                        else:
                            hue = (int(hash_digest, 16) % 360) / 360.0
                    except AttributeError:  # Handle potential encoding errors
                        pass  # Fall through to task name hashing

        if args.color_by_bucket:
            matching_rows = tasks_df.loc[tasks_df["Task Name"] == task_name]
            if not matching_rows.empty:
                bucket = matching_rows["Bucket Name"].iloc[0]  # Get the first bucket
                if pd.notna(bucket):  # Check if bucket exists and is not NaN
                    try:
                        # Hash the bucket
                        hash_object = hashlib.md5(str(bucket).encode())  # Use str() for safety
                        hash_digest = hash_object.hexdigest()
                        if args.alternate_colors:
                            hue = (int(HUE_STEP * task_index) % 360) / 360.0
                        else:
                            hue = (int(hash_digest, 16) % 360) / 360.0
                    except AttributeError:  # Handle potential encoding errors
                        pass  # Fall through to task name hashing

        print(f"Task: {task_name}, Hue: {hue}")

        # If hue wasn't set by label (either disabled, no matching rows, NaN label, or encoding error)
        if hue is None:
            # Hash the task name
            hash_object = hashlib.md5(str(task_name).encode())
            hash_digest = hash_object.hexdigest()
            hue = (int(hash_digest, 16) % 360) / 360.0

        # Generate color
        rgb = colorsys.hls_to_rgb(hue, lightness, saturation)
        hex_color = f"#{int(rgb[0] * 255):02x}{int(rgb[1] * 255):02x}{int(rgb[2] * 255):02x}"
        task_colors[task_name] = hex_color

    # Determine the year for the calendar
    target_year = args.year
    if not target_year:
        # Find the earliest year from valid start dates
        valid_start_dates = pd.to_datetime(tasks_df["Start date"], errors="coerce").dropna()
        if not valid_start_dates.empty:
            target_year = valid_start_dates.min().year
        else:
            # If no valid start dates, default to current year or ask user? Defaulting for now.
            target_year = datetime.now().year
            warning_msg = "Warning: Could not determine year from 'Start date'."
            print(
                f"{warning_msg} Defaulting to current year: {target_year}",
                file=sys.stderr,
            )

    print(f"Generating calendar for the year: {target_year}")

    # Process month argument if provided
    target_month = None
    if args.month:
        try:
            # Try to convert to integer (1-12)
            target_month = int(args.month)
            if target_month < 1 or target_month > 12:
                print(f"Error: Month must be between 1 and 12, got {target_month}", file=sys.stderr)
                sys.exit(1)
        except ValueError:
            # Try to match month name
            month_names = {name.lower(): num for num, name in enumerate(calendar.month_name) if num > 0}
            target_month = month_names.get(args.month.lower())
            if target_month is None:
                print(f"Error: Invalid month name: {args.month}", file=sys.stderr)
                valid_months = ', '.join(name for name in calendar.month_name[1:])
                print(f"Valid month names are: {valid_months}", file=sys.stderr)
                sys.exit(1)
    
    html_content = generate_calendar_html(
        tasks_df, args.no_wrap_text, target_year, task_colors, args.prefix_labels, target_month
    )

    try:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(html_content)
        print(f"Successfully generated calendar HTML: {args.output}")
    except Exception as e:
        print(f"Error writing HTML file: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
