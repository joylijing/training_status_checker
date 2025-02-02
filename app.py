from flask import Flask, request, render_template, send_file
import pandas as pd
from io import BytesIO
from fuzzywuzzy import fuzz
import re
import os
from datetime import datetime, timedelta  # New import for date checking

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Get uploaded files
        table1_file = request.files["table1"]
        table2_file = request.files["table2"]

        # Load the Excel files into DataFrames using openpyxl
        table1 = pd.read_excel(table1_file, engine="openpyxl")
        table2 = pd.read_excel(table2_file, engine="openpyxl")

        # Clean and preprocess the "Time in Session" column in Table2
        table2["Time in Session"] = table2["Time in Session"].apply(parse_time_in_session)

        # Perform the check using email and fuzzy matching
        updated_table1 = perform_check(table1, table2)

        # Third-round check: If "Date of Entry" column exists, then update training status based on date.
        # If today is two weeks past the date of entry and training is not "Webinar Training Complete",
        # mark it as "Webinar Training Incomplete."
        if "Date of entry" in updated_table1.columns:
            today = datetime.today().date()
            for index, row in updated_table1.iterrows():
                date_entry = row["Date of entry"]
                if pd.notna(date_entry):
                    try:
                        entry_date = pd.to_datetime(date_entry).date()
                    except Exception:
                        continue  # skip rows with unparseable dates
                    if (today - entry_date).days >= 14 and updated_table1.at[index, "Training status"] != "Webinar Training Complete":
                        updated_table1.at[index, "Training status"] = "Webinar Training Incomplete"

        # Save the updated Table1 to a new Excel file.
        # Note: We are keeping all columns of Table1 for the output.
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            updated_table1.to_excel(writer, index=False, sheet_name="Updated Table1")
        output.seek(0)

        # Return the updated file for download
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="Updated_Table1.xlsx",
        )

    return render_template("index.html")


def parse_time_in_session(time_str):
    """
    Convert a time string (e.g., "2 hours 8 minutes" or "59 minutes") to total minutes.
    If the input is already numeric, return it as a float.
    If the input is invalid or missing, return NaN.
    """
    if pd.isna(time_str) or time_str == "":
        return float("nan")  # Return NaN for missing or empty values

    if isinstance(time_str, (int, float)):
        return float(time_str)  # Return numeric values as-is

    # Convert text like "2 hours 8 minutes" or "59 minutes" to total minutes
    total_minutes = 0
    hours_match = re.search(r"(\d+)\s*hours?", str(time_str), re.IGNORECASE)
    minutes_match = re.search(r"(\d+)\s*minutes?", str(time_str), re.IGNORECASE)

    if hours_match:
        total_minutes += int(hours_match.group(1)) * 60  # Convert hours to minutes
    if minutes_match:
        total_minutes += int(minutes_match.group(1))  # Add minutes

    return float(total_minutes) if total_minutes > 0 else float("nan")


def perform_check(table1, table2):
    # Ensure the "Training status" and "Time in Session" columns are of the correct types
    if "Training status" not in table1.columns:
        table1["Training status"] = ""  # Add the column if it doesn't exist
    table1["Training status"] = table1["Training status"].astype(str)

    if "Time in Session" not in table1.columns:
        table1["Time in Session"] = ""  # Add the column if it doesn't exist
    table1["Time in Session"] = table1["Time in Session"].astype(str)

    # Create dictionaries from Table2 for quick lookup
    email_to_time = dict(zip(table2["Email Address"], table2["Time in Session"]))
    name_to_time = dict(zip(table2["First Name"] + " " + table2["Last Name"], table2["Time in Session"]))

    # Iterate through Table1 and update the Training status and Time in Session
    for index, row in table1.iterrows():
        email = row["Email Address"]
        name = row["Guest Editor Name"]

        # Primary check: Match by email
        if email in email_to_time:
            time_in_session = email_to_time[email]
            if pd.isna(time_in_session):
                table1.at[index, "Time in Session"] = "N/A"
            else:
                table1.at[index, "Time in Session"] = f"{int(time_in_session)} minutes"
            update_training_status(table1, index, time_in_session)
        else:
            # Secondary check: Match by name using fuzzy matching with token_set_ratio.
            best_match = None
            best_score = 0

            for table2_name, t2_time in name_to_time.items():
                # Using token_set_ratio handles extra words (like middle names) and 
                # reordering of name tokens better than a simple ratio.
                similarity_score = fuzz.token_set_ratio(name.lower(), table2_name.lower())

                # Experiment with thresholds: 70, 75, 80, 85, etc.
                if similarity_score > best_score and similarity_score >= 80:
                    best_score = similarity_score
                    best_match = table2_name

            if best_match:
                time_in_session = name_to_time[best_match]
                if pd.isna(time_in_session):
                    table1.at[index, "Time in Session"] = "N/A"
                else:
                    table1.at[index, "Time in Session"] = f"{int(time_in_session)} minutes"
                update_training_status(table1, index, time_in_session)
            else:
                # No match found
                table1.at[index, "Time in Session"] = "N/A"
                table1.at[index, "Training status"] = "Webinar Registration Pending"

    return table1


def update_training_status(table1, index, time_in_session):
    if pd.isna(time_in_session):
        table1.at[index, "Training status"] = "Webinar Registration Pending"
    elif isinstance(time_in_session, (int, float)) and time_in_session < 15:
        table1.at[index, "Training status"] = "Webinar Training Pending"
    elif isinstance(time_in_session, (int, float)) and time_in_session >= 15:
        table1.at[index, "Training status"] = "Webinar Training Complete"
    else:
        # Handle unexpected cases (e.g., strings that couldn't be parsed)
        table1.at[index, "Training status"] = "Webinar Registration Pending"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))