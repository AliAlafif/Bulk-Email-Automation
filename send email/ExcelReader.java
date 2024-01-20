import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class ExcelReader {

    public static void main(String[] args) {
        String filePath = "gmail.csv";

        List<String> emails = new ArrayList<>();
        List<String> subjects = new ArrayList<>();
        List<String> bodies = new ArrayList<>();

        try {
            FileInputStream file = new FileInputStream(new File(filePath));
            Scanner scanner = new Scanner(file);

            // skip header
            if (scanner.hasNextLine()) {
                scanner.nextLine();
            }

            while (scanner.hasNextLine()) {
                String line = scanner.nextLine();
                String[] values = line.split(",");

                if (values.length >= 3) {
                    String email = values[0].trim();
                    String subject = values[1].trim();
                    String body = values[2].trim();

                    emails.add(email);
                    subjects.add(subject);
                    bodies.add(body);
                } else {
                    System.out.println("Invalid line: " + line);
                }
            }

            file.close();
            scanner.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        // Open Outlook with new email
        for (int i = 0; i < emails.size(); i++) {
            openOutlookNewEmail(emails.get(i), subjects.get(i), bodies.get(i));
        }
    }

    private static void openOutlookNewEmail(String to, String subject, String body) {
        try {
            String outlookCommand = "\"C:/Program Files/Microsoft Office/root/Office16/OUTLOOK.EXE\" /c ipm.note /m " +
                    "\"" + "mailto:" + to + "?subject=" + subject + "&body=" + body + "\"";
    
            Process process = Runtime.getRuntime().exec(outlookCommand);
            int exitCode = process.waitFor();
    
            if (exitCode == 0) {
                System.out.println("Outlook opened with new email to: " + to);
            } else {
                System.out.println("Failed to open Outlook. Exit code: " + exitCode);
            }
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }
}
