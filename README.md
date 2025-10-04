<h2 align="center" style="color:#0078D4;">📧 Outlook Operations</h2>

Here are some key Outlook automation capabilities handled via Python and Outlook COM integration:

---

### 📨 1. Send Email through Outlook App
- Emails are sent directly via the **Outlook desktop application**.
- Sent emails are visible in the **Outlook "Sent Items"** folder (unlike SMTP-based emails).
- Reply Recipients Address can be configured
- Email can be sent either through Plain Format or HTML Format

---

### 💾 2. Save Outlook Emails as `.msg` Files
- Extract and store Outlook emails in **`.msg` format** for archiving or later processing.

---

### 📂 3. Move Outlook Emails
- Automate moving emails between **folders** (Inbox → Processed, Archive, etc.).
- Useful for organizing and managing mail flow after processing.

---

### 👥 4. Work with Shared Mailbox Addresses
- Access and send emails from **Shared Mailboxes**.
- Enables team collaboration and automation for common inboxes.

---

### 👥 5. Reading Emails from Outlook Folder 
- Access emails from **Inbox (Any Other Folder)**.
- Filter Emails through Subject, Body Content
- Save Attachments of Emails 
- Extracting 'To Address' and 'Sender Address' from email

---

### 👥 6. Reading Emails from Outlook Folder 
- Access emails from **Inbox (Any Other Folder)**.
- Filter Emails through Subject, Body Content
- Save Attachments of Emails 
- Extracting To Address List

---

> 💡 **Tip1:** Use `win32com.client` in Python to interact with Outlook objects such as `MailItem`, `Namespace`, and `MAPIFolder`.
> 💡 **Tip2:** Use `os.startfile('outlook')` from the os library in Python to ensure the Outlook application is properly initialized and synchronized before sending emails. Sometimes, emails sent using email_object.Send() may get stuck in the Outbox if Outlook isn’t fully synced. Launching Outlook via os.startfile('outlook') helps avoid this issue by ensuring the client is active and connected.

