import java.io.BufferedReader;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class CSVProcessor {

    public static void main(String[] args) {
        processCSV("/home/alexander/Downloads/member_details.csv", "/home/alexander/members.xlsx");
    }

    private static void processCSV(String inputFilePath, String outputFilePath) {
        // Define regex patterns
        String idPattern = "\\b\\d{8}\\b";
        String namePattern = "^[^0-9]+$";
        String mobilePattern = "^(\\+254|254)?\\d{10}$";
        String emailPattern = "^[a-zA-Z0-9_+&*-]+(?:\\.[a-zA-Z0-9_+&*-]+)*@(?:[a-zA-Z0-9-]+\\.)+[a-zA-Z]{2,7}$";

        // Create regex pattern objects
        Pattern idRegex = Pattern.compile(idPattern);
        Pattern nameRegex = Pattern.compile(namePattern);
        Pattern mobileRegex = Pattern.compile(mobilePattern);
        Pattern emailRegex = Pattern.compile(emailPattern);

        try (BufferedReader br = new BufferedReader(new FileReader(inputFilePath))) {
            // TODO: Use an Excel library to create and manage Excel sheets
            // For simplicity, we are using FileWriter to create CSVs

            FileWriter invalidSheet = new FileWriter("invalid_data.xlsx");
            FileWriter maleSheet = new FileWriter("male_data.xlsx");
            FileWriter femaleSheet = new FileWriter("female_data.xlsx");
            FileWriter otherSheet = new FileWriter("other_data.xlsx");

            String line;
            while ((line = br.readLine()) != null) {
                String[] fields = line.split(",");

                // Validate ID, Name, Mobile, and Email
                if (validateField(fields[0], idRegex) &&
                        validateField(fields[1], nameRegex) &&
                        validateField(fields[2], mobileRegex) &&
                        validateField(fields[3], emailRegex)) {

                    // Group data based on gender
                    String gender = fields[4].toLowerCase();
                    FileWriter sheet;
                    switch (gender) {
                        case "male":
                            sheet = maleSheet;
                            break;
                        case "female":
                            sheet = femaleSheet;
                            break;
                        default:
                            sheet = otherSheet;
                            break;
                    }

                    // Write valid data to the corresponding sheet
                    sheet.write(String.join(",", fields) + "\n");
                } else {
                    // Write invalid data to the invalid sheet
                    invalidSheet.write(String.join(",", fields) + "\n");
                }
            }

            // Close the file writers
            invalidSheet.close();
            maleSheet.close();
            femaleSheet.close();
            otherSheet.close();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static boolean validateField(String value, Pattern pattern) {
        Matcher matcher = pattern.matcher(value);
        return matcher.matches();
    }
}
