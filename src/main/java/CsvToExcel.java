import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.ListDocument;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.List;
import java.util.regex.Pattern;
import java.util.stream.Collectors;




public class CsvToExcel {
    private static final String ID_REGEX = "\\d{8}";
    private static final String MOBILE_REGEX = "^(\\+254|254|07)\\d{9}$";
    private static final String EMAIL_REGEX = "^[a-zA-Z0-9_+&*-]+(?:\\.[a-zA-Z0-9_+&*-]+)*@(?:[a-zA-Z0-9-]+\\.)+[a-zA-Z]{2,7}$";

    public static void main(String[] args) {
        // Input and output file paths
        String csvFilePath = "/home/brutal/Downloads/member_details.csv";
        String excelFilePath = "/mnt/hgfs/Ubuntu back/csv/members.xlsx";

        System.out.println("....................Process is starting....................");

        long startTime = System.currentTimeMillis();

        try (CSVReader reader = new CSVReader(new FileReader(csvFilePath))) {
            ListDocument.List<String[]> data;
            try {
                data = reader.readAll();
            } catch (CsvException e) {
                System.err.println("Error reading CSV file!!");
                e.printStackTrace();
                return;
            }

            // Skip the second row of headers
            data.remove(1);

            // Separate gender data
            List<String[]> maleData = filterAndValidateData(data, "Male");
            List<String[]> femaleData = filterAndValidateData(data, "Female");
            List<String[]> invalidData = data.stream()
                    .filter(row -> !"Male".equalsIgnoreCase(row[4]) && !"Female".equalsIgnoreCase(row[4]))
                    .collect(Collectors.toList());

            // Write sheets
            writeSheet(excelFilePath, "Male", maleData);
            writeSheet(excelFilePath, "Female", femaleData);
            writeSheet(excelFilePath, "Invalid data", invalidData);

            System.out.println("....................Process has come to an end....................");

        } catch (Exception e) {
            e.printStackTrace();
        }

        long endTime = System.currentTimeMillis();
        double durationInSeconds = (endTime - startTime) / 1000.0;
        System.out.println("Total execution time: " + formatExecutionTime(durationInSeconds));
    }

    // Helper method to format time
    private static String formatExecutionTime(double seconds) {
        if (seconds >= 60) {
            long minutes = (long) (seconds / 60);
            long remainingSeconds = (long) (seconds % 60);
            return String.format("%d minutes %d seconds", minutes, remainingSeconds);
        } else {
            return String.format("%.2f seconds", seconds);
        }
    }

    // Filter and validate data based on gender
    private static List<String[]> filterAndValidateData(List<String[]> data, String gender) {
        return data.stream()
                .filter(row -> gender.equalsIgnoreCase(row[4]))
                .filter(CsvToExcel::isValid)
                .collect(Collectors.toList());
    }

    // Write data to Excel sheet
    private static void writeSheet(String excelFilePath, String sheetName, List<String[]> data) {
        try (Workbook workbook = getWorkbook(excelFilePath);
             FileOutputStream fileOut = new FileOutputStream(excelFilePath)) {

            Sheet sheet = workbook.getSheet(sheetName);

            if (sheet == null || sheet.getPhysicalNumberOfRows() == 0) {
                sheet = workbook.createSheet(sheetName);

                // Add titles to the sheet
                Row titleRow = sheet.createRow(0);
                String[] titles = {"ID", "Name", "Mobile Number", "Email Address", "Gender"};
                for (int i = 0; i < titles.length; i++) {
                    Cell titleCell = titleRow.createCell(i, CellType.STRING);
                    titleCell.setCellValue(titles[i]);
                }

                // Auto-fit column widths for titles
                for (int i = 0; i < titles.length; i++) {
                    sheet.autoSizeColumn(i);
                }
            }

            // Add data to the sheet
            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(sheet.getLastRowNum() + 1);
                String[] rowData = data.get(i);
                for (int j = 0; j < rowData.length; j++) {
                    Cell cell = row.createCell(j, CellType.STRING);
                    cell.setCellValue(rowData[j]);
                }
            }

            // Auto-fit column
            for (int i = 0; i < sheet.getRow(0).getPhysicalNumberOfCells(); i++) {
                sheet.autoSizeColumn(i);
            }

            workbook.write(fileOut);

        } catch (IOException | EncryptedDocumentException e) {
            e.printStackTrace();
        }
    }

    private static Workbook getWorkbook(String excelFilePath) throws IOException {
        try {
            return WorkbookFactory.create(new FileInputStream(excelFilePath));
        } catch (IOException e) {
            return new XSSFWorkbook();
        }
    }

    private static boolean isValid(String[] row) {
        if (row.length != 5) {
            return false;
        }

        String id = row[0].trim();
        String name = row[1].trim();
        String mobileNumber = row[2].trim();
        String emailAddress = row[3].trim();
        String gender = row[4].trim();

        return Pattern.matches(ID_REGEX, id) &&
                (mobileNumber.isEmpty() || Pattern.matches(MOBILE_REGEX, mobileNumber)) &&
                (emailAddress.isEmpty() || Pattern.matches(EMAIL_REGEX, emailAddress)) &&
                !id.isEmpty() && !name.isEmpty() && !gender.isEmpty() && !mobileNumber.isEmpty() && !emailAddress.isEmpty();
    }
}