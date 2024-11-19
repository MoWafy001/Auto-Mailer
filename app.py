import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Toplevel
from tkinterhtml import HtmlFrame
import pandas as pd
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail
import os
import json


class EmailSenderApp:
    STORAGE_FILE = ".storage"

    def __init__(self, root):
        self.root = root
        self.root.title("Email Sender App")
        self.root.geometry("800x600")

        # Variables
        self.api_token = tk.StringVar()
        self.template = tk.StringVar()
        self.from_email = tk.StringVar()
        self.title = tk.StringVar()
        self.file_data = None
        self.search_timer = None  # Timer for debounce

        # Load settings and template from storage
        self.load_storage()

        # Create Notebook (Tab Container)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True)

        # Create Tabs
        self.create_people_tab()
        self.create_email_template_tab()
        self.create_send_emails_tab()
        self.create_settings_tab()

    def create_people_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="People")

        # File Selection
        file_frame = ttk.Frame(tab)
        file_frame.pack(fill="x", pady=5, padx=5)

        ttk.Label(file_frame, text="Select a CSV or Excel file:").pack(side="left", padx=5)
        ttk.Button(file_frame, text="Browse", command=self.load_file).pack(side="left", padx=5)
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side="left", padx=5)

        # Search Bar
        search_frame = ttk.Frame(tab)
        search_frame.pack(fill="x", pady=5, padx=5)

        ttk.Label(search_frame, text="Search:").pack(side="left", padx=5)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side="left", padx=5)
        search_entry.bind("<KeyRelease>", lambda event: self.debounced_search())

        # Treeview to display data
        self.tree = ttk.Treeview(tab, columns=("Name", "Email"), show="headings")
        self.tree.heading("Name", text="Name")
        self.tree.heading("Email", text="Email")
        self.tree.pack(fill="both", expand=True, pady=5)

        # Total Count
        self.count_label = ttk.Label(tab, text="Total People: 0")
        self.count_label.pack(pady=5)

    def create_email_template_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Email Template")

        ttk.Label(tab, text="Edit HTML Template (use {column_name} for placeholders):").pack(pady=5)

        # HTML Template Editor
        self.template_text = tk.Text(tab, height=15, width=70, wrap="none")
        self.template_text.pack(pady=5)
        self.template_text.insert("1.0", self.template.get())
        self.template_text.bind("<KeyRelease>", lambda event: self.debounced_save_template())

        # Render Preview
        ttk.Label(tab, text="Preview with Dummy Data:").pack(pady=5)
        self.preview_frame = HtmlFrame(tab, horizontal_scrollbar="auto")
        self.preview_frame.pack(fill="both", expand=True)

        self.update_preview()

    # vertical scrollable tab
    def create_send_emails_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Send Emails")

        # scrollable frame
        canvas = tk.Canvas(tab)
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        inner_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=inner_frame, anchor="nw")

        # Email Summary
        self.email_count_label = ttk.Label(inner_frame, text="Emails to be sent: 0")
        self.email_count_label.pack(pady=10)

        # Example Email
        ttk.Label(inner_frame, text="Example Email:").pack(pady=5)
        self.example_email_frame = HtmlFrame(inner_frame, horizontal_scrollbar="auto", vertical_scrollbar="auto", width=799, height=200)
        self.example_email_frame.grid_propagate(False)
        self.example_email_frame.pack_propagate(False)
        self.example_email_frame.pack(fill='x', expand=False)
        self.example_email_frame.pack(pady=5)

        # Send Button
        send_button = ttk.Button(inner_frame, text="Send Emails", command=self.show_send_status_window)
        send_button.pack(pady=10)

    def create_settings_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Settings")

        ttk.Label(tab, text="SendGrid API Token:").pack(pady=10)
        api_token_entry = ttk.Entry(tab, textvariable=self.api_token, show="*")
        api_token_entry.pack(pady=5)
        api_token_entry.bind("<KeyRelease>", lambda event: self.save_storage())

        ttk.Label(tab, text="From Email:").pack(pady=10)
        from_email_entry = ttk.Entry(tab, textvariable=self.from_email)
        from_email_entry.pack(pady=5)
        from_email_entry.bind("<KeyRelease>", lambda event: self.save_storage())

        ttk.Label(tab, text="Email Title:").pack(pady=10)
        title_entry = ttk.Entry(tab, textvariable=self.title)
        title_entry.pack(pady=5)
        title_entry.bind("<KeyRelease>", lambda event: self.save_storage())

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[
            ("CSV Files", "*.csv"),
            ("Excel Files", "*.{xls,xlsx}")
        ])
        if not file_path:
            return

        try:
            if file_path.endswith('.csv'):
                self.file_data = pd.read_csv(file_path)
            else:
                self.file_data = pd.read_excel(file_path)

            self.file_label.config(text=f"Loaded: {file_path}")
            self.populate_people_list()
            self.update_email_summary()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

    def populate_people_list(self):
        if self.file_data is not None:
            # Clear existing rows
            for row in self.tree.get_children():
                self.tree.delete(row)

            # Insert rows
            for _, row in self.file_data.iterrows():
                name = row.get("Name", "")
                email = row.get("Email", "")
                self.tree.insert("", "end", values=(name, email))

            self.count_label.config(text=f"Total People: {len(self.file_data)}")

    def debounced_search(self):
        if self.search_timer:
            self.root.after_cancel(self.search_timer)
        self.search_timer = self.root.after(300, self.search)

    def search(self):
        query = self.search_var.get().strip().lower()
        if self.file_data is None or query == "":
            self.populate_people_list()
            return

        filtered_data = self.file_data[self.file_data.apply(
            lambda row: query in str(row["Name"]).lower() or query in str(row["Email"]).lower(),
            axis=1
        )]

        for row in self.tree.get_children():
            self.tree.delete(row)

        for _, row in filtered_data.iterrows():
            self.tree.insert(
                "",
                "end",
                values=(row.get("Name", ""), row.get("Email", ""))
            )

        self.count_label.config(text=f"Total People: {len(filtered_data)}")

    def debounced_save_template(self):
        if self.search_timer:
            self.root.after_cancel(self.search_timer)
        self.search_timer = self.root.after(300, self.save_template)

    def save_template(self):
        self.template.set(self.template_text.get("1.0", "end").strip())
        self.update_preview()
        self.save_storage()

    def update_preview(self):
        dummy_data = {"Name": "John Doe", "Email": "john.doe@example.com"}
        filled_template = self.template.get().format(**dummy_data)
        self.preview_frame.set_content(filled_template)

    def update_email_summary(self):
        if self.file_data is not None:
            self.email_count_label.config(text=f"Emails to be sent: {len(self.file_data)}")
            self.update_example_email()

    def update_example_email(self):
        if self.file_data is not None and not self.file_data.empty:
            example_row = self.file_data.iloc[0].to_dict()
            filled_template = self.template.get().format(**example_row)
            self.example_email_frame.set_content(filled_template)

    def show_send_status_window(self):
        if self.file_data is None:
            messagebox.showerror("Error", "No data loaded!")
            return

        if not self.api_token.get() or not self.from_email.get() or not self.title.get():
            messagebox.showerror("Error", "SendGrid API token, From Email or Title is missing!")
            return

        status_window = Toplevel(self.root)
        status_window.title("Send Status")
        ttk.Label(status_window, text="Sending emails...").pack(pady=10)

        progress = ttk.Progressbar(status_window, orient="horizontal", mode="determinate", length=300)
        progress.pack(pady=10)
        progress["maximum"] = len(self.file_data)
        progress["value"] = 0

        # log text
        self.log_text = tk.Text(status_window, height=10, width=60, wrap="word")
        self.log_text.pack(pady=10)
        self.log_text.config(state=tk.DISABLED)

        try:
            for _, row in self.file_data.iterrows():
                email_content = self.template.get().format(**row)
                email = Mail(
                    from_email=self.from_email.get(),
                    to_emails=row["Email"],
                    subject=self.title.get(),
                    html_content=email_content
                )
                sg_client = SendGridAPIClient(self.api_token.get())
                response = sg_client.send(email)

                progress["value"] += 1
                if response.status_code != 202:
                    self.log(f"Failed to send email to {row['Email']}: {response.body}")
                    print(response.body)
                else:
                    self.log(f"Email sent to {row['Email']}")
                self.root.update_idletasks()

            messagebox.showinfo("Success", "Emails sent successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Error sending emails: {e}")
            print(e)
        finally:
            status_window.destroy()

    def log(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.config(state=tk.DISABLED)

    def load_storage(self):
        if os.path.exists(self.STORAGE_FILE):
            with open(self.STORAGE_FILE, "r") as file:
                data = json.load(file)
                self.api_token.set(data.get("api_token", "").replace("\n", "").replace("\r", "").strip())
                self.template.set(data.get("template", "").replace("\n", "").strip())
                self.from_email.set(data.get("from_email", "").replace("\n", "").replace("\r", "").strip())
                self.title.set(data.get("title", "").replace("\n", "").replace("\r", "").strip())

    def save_storage(self):
        self.api_token.set(self.api_token.get().replace("\n", "").replace("\r", "").strip())
        self.template.set(self.template.get().strip())
        self.from_email.set(self.from_email.get().replace("\n", "").replace("\r", "").strip())
        self.title.set(self.title.get().replace("\n", "").replace("\r", "").strip())

        data = {
            "api_token": self.api_token.get().replace("\n", "").replace("\r", "").strip(),
            "template": self.template.get().strip(),
            "from_email": self.from_email.get().replace("\n", "").replace("\r", "").strip(),
            "title": self.title.get().replace("\n", "").replace("\r", "").strip()
        }
        with open(self.STORAGE_FILE, "w") as file:
            json.dump(data, file, indent=4)


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()

