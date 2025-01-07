![image](https://github.com/user-attachments/assets/ff899cac-3a48-4269-9b3c-14abf0849b9a)# OutlookEmailAutomation
A VBA macro to automate sending emails with attachments using Outlook, including dynamic content and personalized salutations. Make sure to give required permissions to your macro to access your Outlook application on your system to integrate with VBA macro.


# OutlookEmailAutomation

This repository contains a VBA macro to automate sending emails with attachments using Outlook or Gmail. The macro includes features such as dynamic content, personalized salutations, and HTML signatures.

## Features
- Send emails to multiple recipients with a single click.
- Attach files dynamically based on Excel sheet data.
- Include personalized salutations using recipient names.
- Embed images in the email signature without showing them as attachments.
- Compatible with both Outlook and Gmail.

## Setup
1. **Clone the repository:**
   ```sh
   git clone https://github.com/NeerajSugara/OutlookEmailAutomation.git
2. **Open the VBA editor in Excel:**

- Press Alt + F11 to open the VBA editor.
- Import the SendEmailsWithAttachments.vba file.
3. **Configure the macro:**

- Update the email addresses, paths, and other details in the macro as needed.
4. **Run the macro:**

- Press F5 to run the macro and send emails.
## Usage
- Ensure that your Excel sheet is set up with the required data:

- Column A: Email addresses <!-- mandatory field -->
- Column B: Recipient names <!-- mandatory field -->
- Column C: Message bodies <!-- mandatory field -->
- Column D: Attachment paths <!-- Not mandatory field -->
- Column E: Subject <!-- mandatory field -->
- Column F: Sender Name <!-- mandatory field -->
- Customize the macro as needed to fit your requirements.
