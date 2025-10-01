"""
Created by: Randy Grizzelli
Version: 0.3 (Added progress bar and verbose logging)
Date: 2025-10-01
Purpose: Export Outlook emails (Inbox + Sent Items) to/from/cc a specific address into a PDF,
         including attachments saved locally. Prompts user for email at runtime.
GitHub: https://github.com/rsgrizz
"""


import os
import win32com.client
from tqdm import tqdm  # Import the tqdm library for the progress bar
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet

def export_outlook_emails():
    # ---- DYNAMIC CONFIGURATION ----
    target_email = input("Enter the target email address to search for: ")

    if not target_email or "@" not in target_email:
        print("\n❌ Invalid or empty email address. Script will now exit.")
        return

    safe_filename_base = target_email.replace("@", "_at_").replace(".", "_")
    output_file = f"{safe_filename_base}_export.pdf"
    attachments_dir = f"{safe_filename_base}_attachments"

    print(f"\n▶️  Starting Export Process...")
    print(f"Searching for emails related to: {target_email}")
    print(f"Output PDF will be saved as: {output_file}")
    print(f"Attachments will be saved to: {attachments_dir}/")
    # --------------------------------

    print("\n⏳ Connecting to Outlook application...")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    print("✅ Successfully connected to Outlook.")

    print("Accessing Inbox and Sent Items folders...")
    inbox = outlook.GetDefaultFolder(6)
    sent_items = outlook.GetDefaultFolder(5)
    
    # Collect all items first to provide a single, accurate progress bar
    print("Collecting emails to scan (this may take a moment for large mailboxes)...")
    all_items = list(inbox.Items) + list(sent_items.Items)
    print(f"Found a total of {len(all_items)} emails to scan.")
    
    if not os.path.exists(attachments_dir):
        os.makedirs(attachments_dir)

    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(output_file, pagesize=letter)
    story = []
    found_count = 0

    # Use tqdm to create a dynamic progress bar for the main loop
    for item in tqdm(all_items, desc="Scanning emails", unit="item"):
        try:
            if item.Class == 43:  # Ensure it's a MailItem
                sender = item.SenderEmailAddress or "Unknown"
                subject = item.Subject or "No Subject"
                body = item.Body or ""
                date = item.SentOn.strftime("%Y-%m-%d %H:%M") if item.SentOn else "Unknown"

                to_list = [r.Address for r in item.Recipients if r.Type == 1]
                cc_list = [r.Address for r in item.Recipients if r.Type == 2]

                match = False
                if target_email.lower() in (sender or "").lower():
                    match = True
                elif any(target_email.lower() in (r or "").lower() for r in to_list + cc_list):
                    match = True

                if match:
                    found_count += 1
                    tqdm.write(f"  [+] Match #{found_count}: Found email from '{sender}' with subject '{subject}'")

                    story.append(Paragraph(f"<b>Subject:</b> {subject}", styles["Normal"]))
                    story.append(Paragraph(f"<b>From:</b> {sender}", styles["Normal"]))
                    story.append(Paragraph(f"<b>To:</b> {', '.join(to_list)}", styles["Normal"]))
                    story.append(Paragraph(f"<b>CC:</b> {', '.join(cc_list)}", styles["Normal"]))
                    story.append(Paragraph(f"<b>Date:</b> {date}", styles["Normal"]))
                    story.append(Paragraph(f"<b>Body:</b><br/>{body.replace(chr(10), '<br/>').replace(chr(13), '')}", styles["Normal"]))

                    if item.Attachments.Count > 0:
                        tqdm.write(f"    -> Found {item.Attachments.Count} attachment(s). Saving...")
                        saved_files = []
                        for att in item.Attachments:
                            safe_name = "".join(c for c in att.FileName if c.isalnum() or c in (' ', '.', '_')).rstrip()
                            save_path = os.path.join(attachments_dir, safe_name)
                            att.SaveAsFile(save_path)
                            saved_files.append(save_path)
                            tqdm.write(f"       - Saved '{safe_name}'")
                        story.append(Paragraph(f"<b>Attachments saved:</b> {', '.join(saved_files)}", styles["Normal"]))
                    else:
                        story.append(Paragraph("<b>Attachments:</b> None", styles["Normal"]))

                    story.append(Spacer(1, 20))

        except Exception as e:
            tqdm.write(f"\n⚠️ Could not process an item. Error: {e}")

    if found_count > 0:
        print(f"\nFound {found_count} total matching emails. Building PDF...")
        doc.build(story)
        print(f"✅ Export complete: {output_file}")
        print(f"✅ Attachments saved in: {os.path.abspath(attachments_dir)}")
    else:
        print("\nNo matching emails were found.")

if __name__ == "__main__":
    export_outlook_emails()
