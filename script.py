import tkinter as tk
from tkinter import filedialog, messagebox
import smtplib
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Global variable to store recipients
recipients = []

def send_email_from_gui():
    """Send email using GUI-provided details."""
    email = email_entry.get()
    password = password_entry.get()
    smtp_server = smtp_server_entry.get()
    smtp_port = smtp_port_entry.get()
    subject = subject_entry.get()
    body = body_text.get("1.0", tk.END).strip()

    if not (email and password and smtp_server and smtp_port and subject and body and recipients):
        messagebox.showerror("Error", "Please fill in all fields and load a CSV file with email addresses.")
        return

    try:
        smtp_port = int(smtp_port)  # Ensure the port is an integer
        send_email(subject, body, email, password, smtp_server, smtp_port, recipients)
    except ValueError:
        messagebox.showerror("Error", "SMTP port must be a number.")

def send_email(subject, body, email, password, smtp_server, smtp_port, recipients):
    """Send an email to a list of recipients one at a time and handle errors."""
    try:
        # Set up the server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(email, password)

        success_count = 0
        failed_recipients = []

        for recipient in recipients:
            try:
                # Clone the email template for the current recipient
                msg = MIMEMultipart()  # Create a new email instance for each recipient.
                msg.attach(MIMEText(body, 'plain'))  # Attach the body content again.
                msg['Subject'] = subject
                msg['From'] = email
                msg['To'] = recipient  # Set 'To' for the current recipient.
                server.sendmail(email, recipient, msg.as_string())
                success_count += 1
            except Exception as e:
                # Log the failed recipient and error
                failed_recipients.append((recipient, str(e)))

        server.quit()

        # Show success and failure summary
        if not failed_recipients:
            messagebox.showinfo("Success", f"Emails sent successfully to all {len(recipients)} recipients!")
        else:
            failed_count = len(failed_recipients)
            failed_details = "\n".join([f"{recipient}: {error}" for recipient, error in failed_recipients])
            messagebox.showwarning(
                "Partial Success",
                f"Emails sent successfully to {success_count} recipients, but failed for {failed_count}:\n\n{failed_details}"
            )
    except Exception as e:
        messagebox.showerror("Error", f"Failed to send emails: {e}")

def load_csv():
    """Allow user to select a CSV file and read email addresses for clients accepting email marketing."""
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if not file_path:
        return

    try:
        # Read the CSV file
        email_list = []
        with open(file_path, newline='', encoding='utf-8') as csvfile:
            csvreader = csv.DictReader(csvfile)
            for row in csvreader:
                # Check if the client accepts email marketing
                email = row.get("Email", "").strip()
                accepts_marketing = row.get("Accepts Email Marketing", "").strip().lower()
                
                if email and accepts_marketing == "yes":
                    email_list.append(email)

        if not email_list:
            messagebox.showerror("Error", "No email addresses found for clients accepting email marketing.")
            return

        global recipients
        recipients = email_list
        messagebox.showinfo("Success", f"Loaded {len(email_list)} email addresses.")

        # Enable the "Send Email" button
        send_button["state"] = "normal"

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load or read the CSV file: {e}")

# Initialize the GUI application
root = tk.Tk()
root.title("Email Sender")

# Form fields for credentials
tk.Label(root, text="Email Address:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
email_entry = tk.Entry(root, width=30)
email_entry.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Password:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
password_entry = tk.Entry(root, show="*", width=30)
password_entry.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="SMTP Server:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
smtp_server_entry = tk.Entry(root, width=30)
smtp_server_entry.insert(0, "smtp.gmail.com")  # Default for Outlook
smtp_server_entry.grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="SMTP Port:").grid(row=3, column=0, padx=10, pady=5, sticky="e")
smtp_port_entry = tk.Entry(root, width=30)
smtp_port_entry.insert(0, "587")  # Default for Outlook
smtp_port_entry.grid(row=3, column=1, padx=10, pady=5)

# Form fields for email content
tk.Label(root, text="Subject:").grid(row=4, column=0, padx=10, pady=5, sticky="e")
subject_entry = tk.Entry(root, width=30)
subject_entry.grid(row=4, column=1, padx=10, pady=5)

tk.Label(root, text="Body:").grid(row=5, column=0, padx=10, pady=5, sticky="e")
body_text = tk.Text(root, width=30, height=5)
body_text.grid(row=5, column=1, padx=10, pady=5)

# Button to load CSV file
load_csv_button = tk.Button(root, text="Load CSV File", command=load_csv)
load_csv_button.grid(row=6, column=0, columnspan=2, pady=10)

# Send email button that will appear after CSV is loaded
send_button = tk.Button(root, text="Send Emails", command=send_email_from_gui, state="disabled")
send_button.grid(row=8, column=0, columnspan=2, pady=10)

# Run the application
root.mainloop()
