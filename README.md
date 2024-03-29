# Excel-Email-Reader

**Overview:**

Excel Email Reader is a Java-based utility designed to simplify the process of sending bulk emails with personalized content. The tool reads email-related data from a CSV file, generates individualized emails, and employs Microsoft Outlook for efficient delivery.

**Key Features:**

- **CSV File Parsing:** Extract recipient email addresses, subjects, and email bodies from a designated CSV file.
  
- **Data Validation:** Verify the CSV file for valid entries in each email, subject, and body. Invalid lines are flagged and reported in the console.

- **Email Composition:** Dynamically compose personalized emails by combining recipient addresses, subjects, and bodies.

- **Outlook Integration:** Leverage the Runtime class to execute Microsoft Outlook processes, opening new email windows with recipient addresses, subjects, and bodies.

- **Bulk Email Handling:** Effectively manage bulk emails, simplifying the process of sending tailored emails to multiple recipients concurrently.

- **Error Handling:** Implement robust error-handling mechanisms to catch exceptions during file reading, email composition, and Outlook process execution. Detailed error messages are printed to the console for troubleshooting.

**Usage Instructions:**

1. **CSV File Configuration:**
   - Modify the `filePath` variable to specify the path of the CSV file containing email information.
   - Ensure the CSV file includes columns for email addresses, subjects, and bodies.

2. **Outlook Path Configuration:**
   - Update the `outlookCommand` variable with the correct path to the Outlook executable (`OUTLOOK.EXE`). Adjust the path enclosed in quotes according to your Outlook installation directory.

3. **Running the Program:**
   - Execute the program to read data from the CSV file, compose emails, and open Microsoft Outlook with the prepared emails.


