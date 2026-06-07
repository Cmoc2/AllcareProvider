# -*- coding: utf-8 -*-
"""
Created on Fri Apr  3 12:44:37 2026

@author: ChristianOrtiz
"""

import argparse
from datetime import datetime, timedelta
import pandas as pd

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import sys

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders


REQUIRED_COLUMNS = {
    "User",
    "Form Status",
    "Form Date",
}

REMINDER_START_DAYS = 5 #start remidners after this many days
AUDIT_RISK_DAYS = 30 #existing audit threshhold

sender_email = os.getenv("EMAIL_SENDER")
app_password = os.getenv("EMAIL_APP_PASSWORD")


def launch_gui(start_callback):
    """
    Opens a GUI to select an Excel file and start processing.

    start_callback: function that accepts a single argument (file_path)
                    and runs the summary logic.
    """

    selected_file = {"path": None}

    def select_file():
        file_path = filedialog.askopenfilename(
            title="Select Visit Report Excel File",
            filetypes=[("Excel Files", "*.xls *.xlsx")]
        )

        if not file_path:
            return  # user canceled dialog

        if not file_path.lower().endswith((".xls",".xlsx")):
            messagebox.showerror(
                "Invalid File Type",
                "Please select a valid .xlsx or .xls Excel file."
            )
            return

        selected_file["path"] = file_path
        select_button.config(text=os.path.basename(file_path))

    def start_processing():
        if selected_file["path"] is None:
            messagebox.showwarning(
                "No File Selected",
                "Please select an Excel file before starting."
            )
            return

        try:
            root.destroy()  # close GUI before processing
            start_callback(selected_file["path"])
        except Exception as e:
            messagebox.showerror(
                "Processing Error",
                f"An error occurred:\n\n{str(e)}"
            )

    def exit_app():
        root.destroy()
        sys.exit(0)

    root = tk.Tk()
    root.title("Visit Note Summary Generator")
    root.geometry("500x200")
    root.resizable(False, False)

    main_frame = tk.Frame(root, padx=20, pady=20)
    main_frame.pack(expand=True, fill="both")

    tk.Label(
        main_frame,
        text="Select the Visit Report Excel file to begin",
        font=("Segoe UI", 11)
    ).pack(pady=(0, 15))

    select_button = tk.Button(
        main_frame,
        text="Select Excel File",
        width=30,
        command=select_file
    )
    select_button.pack(pady=5)

    start_button = tk.Button(
        main_frame,
        text="Start Generating Summary",
        width=30,
        command=start_processing
    )
    start_button.pack(pady=5)

    cancel_button = tk.Button(
        main_frame,
        text="Cancel / Exit",
        width=30,
        command=exit_app
    )
    cancel_button.pack(pady=5)

    root.mainloop()


def show_status_preview(summary_df):
    """
    Displays a scrollable preview window showing summary results per user.
    summary_df: pandas DataFrame produced by generate_summary()
    """
    
    def on_send_email(row, button):
        success = send_email_to_user(row)
    
        if success:
            button.config(
                text="Sent",
                state="disabled"
            )

    preview = tk.Toplevel()
    preview.title("Status Preview")
    preview.geometry("750x550")

    container = ttk.Frame(preview)
    container.pack(fill="both", expand=True)

    canvas = tk.Canvas(container)
    scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=scrollbar.set)

    scroll_frame = ttk.Frame(canvas)
    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    def _on_mousewheel(event):
        # Windows / macOS
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_linux_scroll_up(event):
        canvas.yview_scroll(-1, "units")

    def _on_linux_scroll_down(event):
        canvas.yview_scroll(1, "units")

    # Mouse wheel bindings
    canvas.bind_all("<MouseWheel>", _on_mousewheel)      # Windows / macOS
    canvas.bind_all("<Button-4>", _on_linux_scroll_up)   # Linux scroll up
    canvas.bind_all("<Button-5>", _on_linux_scroll_down) # Linux scroll down

    ttk.Label(
        scroll_frame,
        text="Visit Note Status Preview",
        font=("Segoe UI", 14, "bold")
    ).pack(pady=10)

    for _, row in summary_df.iterrows():
        frame = ttk.Frame(scroll_frame, relief="solid", padding=10)
        frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame, text=f"User: {row['User']}", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        ttk.Label(frame, text=f"Pending: {row['Total_Pending']}").pack(anchor="w")
        ttk.Label(frame, text=f"To Be Corrected: {row['Total_To_Be_Corrected']}").pack(anchor="w")

        if row["Has_Over_30_Day_Notes"]:
            ttk.Label(frame, text="⚠ Audit Risk: Yes", foreground="red").pack(anchor="w")

        send_btn = ttk.Button(
            frame,
            text="Send Email"
            )
        send_btn.config(
            command=lambda r=row, b=send_btn: on_send_email(r,b)
        )
        send_btn.pack(anchor="e", pady=5)

    footer = ttk.Frame(scroll_frame)
    footer.pack(fill="x", pady=15)

    ttk.Button(
        footer,
        text="Email All",
        command=lambda: send_email_all(summary_df)
    ).pack(side="left", padx=10)

    ttk.Button(
        footer,
        text="Export to Excel",
        command=lambda: export_summary_via_gui(summary_df)
    ).pack(side="left", padx=10)

    ttk.Button(
        footer,
        text="Close",
        command=preview.destroy
    ).pack(side="right", padx=10)



def load_data(path: str) -> pd.DataFrame:
    """
    Load the Visit Report Excel file.
    Column headers are on row 6 (0-based index = 5).
    """
    path_lower = path.lower()

    if path_lower.endswith(".xlsx"):
        df = pd.read_excel(
            path,
            engine="openpyxl",
            skiprows=5
        )

    elif path_lower.endswith(".xls"):
        df = pd.read_excel(
            path,
            engine="xlrd",
            skiprows=5
        )

    else:
        raise ValueError("Unsupported file type. Please select an .xls or .xlsx file.")

    df.columns = df.columns.str.strip()
    return df


def validate_headers(df: pd.DataFrame):
    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(
            f"Invalid report format. Missing required columns: {', '.join(missing)}"
        )


def generate_summary(df: pd.DataFrame) -> pd.DataFrame:

    today = datetime.today()

    # Only relevant statuses
    df = df[df["Form Status"].isin(["Pending", "To Be Corrected"])].copy()

    # Parse Form Date
    df["Form Date"] = pd.to_datetime(df["Form Date"], errors="coerce")

    # Calculate note age in days
    df["Days_Old"] = (today - df["Form Date"]).dt.days

    # Reminder eligibility (grace period applied)
    df["Eligible_For_Reminder"] = df["Days_Old"] > REMINDER_START_DAYS

    # Audit risk logic (unchanged)
    df["Audit_Risk"] = df["Days_Old"] >= AUDIT_RISK_DAYS

    # **ONLY include notes eligible for reminders**
    reminder_df = df[df["Eligible_For_Reminder"]].copy()

    if reminder_df.empty:
        return pd.DataFrame(
            columns=[
                "User",
                "Total_Pending",
                "Total_To_Be_Corrected",
                "Has_Over_30_Day_Notes",
                "Email",
            ]
        )

    summary = (
        reminder_df.groupby("User")
        .agg(
            Total_Pending=("Form Status", lambda x: (x == "Pending").sum()),
            Total_To_Be_Corrected=("Form Status", lambda x: (x == "To Be Corrected").sum()),
            Has_Over_30_Day_Notes=("Audit_Risk", "any"),
        )
        .reset_index()
    )

    # Placeholder for email mapping
    summary["Email"] = "christian.ortiz@allcareprovider.com"

    return summary



def print_stdout(summary: pd.DataFrame):
    """
    Print a readable stdout report.
    """
    for _, row in summary.iterrows():
        print(f"User: {row['User']}")
        print(f"  Pending: {row['Total_Pending']}")
        print(f"  To Be Corrected: {row['Total_To_Be_Corrected']}")

        if row["Has_Over_30_Day_Notes"]:
            print("  ⚠ WARNING: One or more notes are over 30 days old.")
            print("    Risk of audit due to government-mandated completion timelines.")

        print("-" * 50)


def build_html_email_body(user, pending, corrected, risk):
    warning_section = ""

    if risk:
        warning_section = """
        <tr>
            <td colspan="2" style="
                color: #721c24;
                background-color: #f8d7da;
                padding: 12px;
                border: 1px solid #f5c6cb;
                font-weight: bold;
            ">
                ⚠ Compliance Warning<br>
                One or more visit notes are over <strong>30 days old</strong>.
                Incomplete or late documentation may place you at risk of audit
                under government‑mandated timelines.
            </td>
        </tr>
        """

    html = f"""
    <html>
    <body style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #333;">
        <p>Hello {user},</p>

        <p>
            This is a reminder that you currently have visit notes pending for more than 5 days and require action
            before they can be processed by the office. Please note that it is expected for all notes and documentation to be submitted within 48 hours of the visit.
            Corrections to also be submitted in a timely manner.
        </p>

        <table cellpadding="8" cellspacing="0" style="border-collapse: collapse; margin-top: 10px;">
            <tr>
                <th style="border-bottom: 2px solid #ccc; text-align: left;">
                    Status
                </th>
                <th style="border-bottom: 2px solid #ccc; text-align: right;">
                    Count
                </th>
            </tr>
            <tr>
                <td>Pending</td>
                <td style="text-align: right;">{pending}</td>
            </tr>
            <tr>
                <td>To Be Corrected</td>
                <td style="text-align: right;">{corrected}</td>
            </tr>
            {warning_section}
        </table>

        <p style="margin-top: 15px;">
            Please log into the system and complete or correct these notes as soon
            as possible.
        </p>

        <p>
            If you believe you have received this message in error, please contact
            the office for assistance.
        </p>

        <p style="margin-top: 20px;">
            Thank you,<br>
            <strong>HR Department</strong>
        </p>
    </body>
    </html>
    """

    return html


def send_email_to_user(user_row):
    user = user_row["User"]
    pending = user_row["Total_Pending"]
    corrected = user_row["Total_To_Be_Corrected"]
    risk = user_row["Has_Over_30_Day_Notes"]

    subject = "You have notes to Send to Office"

    msg = MIMEMultipart("alternative")
    msg["From"] = sender_email
    msg["To"] = user_row["Email"]
    msg["Subject"] = subject
    cc_value = 'susan.balmadrid@allcareprovider.com'  #glendy.arce@allcareprovider.com 
    
    if cc_value == '':
        recipients = msg['To'].split(',')
    else:
        msg['CC'] = cc_value
        recipients = msg['To'].split(',') + msg['CC'].split(',')

    html_body = build_html_email_body(
        user=user,
        pending=pending,
        corrected=corrected,
        risk=risk
    )

    msg.attach(MIMEText(html_body, "html"))

    try:
        with smtplib.SMTP("smtp.office365.com", 587) as server:
            server.starttls()
            server.login(sender_email, app_password)
            server.sendmail(sender_email, recipients, msg.as_string())

        print(f"Email sent successfully to {recipients}")
        return True

    except Exception as e:
        print(f"Failed to send email. Reason: {e}")
        return False


    
    print(f"[EMAIL] Sending email to {user}")
    print(f"Pending: {pending}, To Be Corrected: {corrected}, Audit Risk: {risk}")


def send_email_all(summary_df):
    for _, row in summary_df.iterrows():
        if row["Total_Pending"] == 0 and row["Total_To_Be_Corrected"] == 0:
            continue
        send_email_to_user(row)


def export_excel(summary: pd.DataFrame, output_path: str):
    """
    Export summary to Excel for manual email review.
    """
    export_columns = [
        "User",
        "Total_Pending",
        "Total_To_Be_Corrected",
        "Email",
        "Has_Over_30_Day_Notes",
    ]

    summary[export_columns].to_excel(
        output_path,
        index=False,
        engine="openpyxl"
    )


def export_summary_via_gui(summary_df):
    path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not path:
        return

    export_excel(summary_df, path)


def run_summary_from_gui(file_path):
    df = load_data(file_path)
    validate_headers(df)
    
    summary = generate_summary(df)
    # Show preview Window
    show_status_preview(summary)
    #Still Print to Terminal stdout
    print_stdout(summary)
    #wait = input("Press Enter to continue.\n") #No longer needed since we have previewWIndow.

def main():
    parser = argparse.ArgumentParser(description="Visit Note Status Summary")
    parser.add_argument("input", help="Path to Visit Report Excel file")
    parser.add_argument(
        "--excel",
        help="Optional output Excel file (e.g. summary.xlsx)",
        required=False
    )

    args = parser.parse_args()

    df = load_data(args.input)
    summary = generate_summary(df)

    print_stdout(summary)

    if args.excel:
        export_excel(summary, args.excel)
        print(f"\nExcel summary written to: {args.excel}")


if __name__ == "__main__": 
    try:
        if not sender_email or not app_password: 
            messagebox.showerror(
                    "Email Configuration Error",
                    "Email credentials are not configured on this machine.\n"
                    "Please contact IT."
                )

            raise RuntimeError("Environment variables undefined.")
        launch_gui(run_summary_from_gui)
    except Exception as e:
        messagebox.showerror("Startup Error", str(e))
