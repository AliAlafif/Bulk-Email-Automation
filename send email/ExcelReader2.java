import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class ExcelReader2 {

    public static void main(String[] args) {
        // change the file name
        String filePath = "gmail.csv";

        List<String> emails = new ArrayList<>();
        List<String> subjects = new ArrayList<>();
        List<String> bodies = new ArrayList<>();

        try {
            FileInputStream file = new FileInputStream(new File(filePath));
            Scanner scanner = new Scanner(file);

            // skip header line(if there is no header remove line 23-25)
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

      
        openAllEmailWindows(emails, subjects, bodies);
    }

    private static void openAllEmailWindows(List<String> toList, List<String> subjectList, List<String> bodyList) {
        List<Process> processes = new ArrayList<>();

        for (int i = 0; i < toList.size(); i++) {
            try {// change the path to your outlook.exe path
                            // from here  "                                                            " to here
                String outlookCommand = "\"C:/Program Files/Microsoft Office/root/Office16/OUTLOOK.EXE\" /c ipm.note /m " +
                        "\"" + "mailto:" + toList.get(i) + "?subject=" + subjectList.get(i) + "&body=" + bodyList.get(i) + "\"";

                Process process = Runtime.getRuntime().exec(outlookCommand);
                processes.add(process);

                System.out.println("Outlook opened with new email to: " + toList.get(i));
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        // Wait for all processes to finish
        for (Process process : processes) {
            try {
                int exitCode = process.waitFor();

                if (exitCode != 0) {
                    System.out.println("Failed to open Outlook. Exit code: " + exitCode);
                }
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
    }
}
