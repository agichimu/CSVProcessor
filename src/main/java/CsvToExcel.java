import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.regex.Pattern;

public class CsvToExcel {
    private static final String ID_REGEX = "\\d{8}";
    private static final String MOBILE_REGEX = "^(\\+254|254|07)\\d{9}$";
    private static final String EMAIL_REGEX = "^[a-zA-Z0-9_+&*-]+(?:\\.[a-zA-Z0-9_+&*-]+)*@(?:[a-zA-Z0-9-]+\\.)+[a-zA-Z]{2,7}$";
    private static final int CHUNK_SIZE = 4500;

    public static void main(String[] args) {
        // Input and output file paths
        String csvFilePath = "/home/brutal/Downloads/member_details.csv";
        String excelFilePath = "/mnt/hgfs/Ubuntu back/csv/members.xlsx";

        System.out.println("....................Process is starting....................");

        long startTime = System.currentTimeMillis();

        try (CSVReader reader = new CSVReader(new FileReader(csvFilePath))) {
            // Skip the header
            String[] header = reader.readNext();

            processAndWriteInChunks(reader, header, excelFilePath);

            System.out.println("....................Process has come to an end....................");

        } catch (CsvException csvException) {
            System.err.println("CSV Exception: " + csvException.getMessage());
        } catch (IOException e) {
            e.printStackTrace();
        }

        long endTime = System.currentTimeMillis();
        double durationInSeconds = (endTime - startTime) / 1000.0;
        System.out.println("Total execution time: " + formatExecutionTime(durationInSeconds));
    }

    private static void processAndWriteInChunks(CSVReader reader, String[] header, String excelFilePath) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        String sheetNameMale = "Male";
        String sheetNameFemale = "Female";
        String sheetNameInvalid = "InvalidData";

        try {
            // Create sheets
            Sheet sheetMale = workbook.createSheet(sheetNameMale);
            Sheet sheetFemale = workbook.createSheet(sheetNameFemale);
            Sheet sheetInvalid = workbook.createSheet(sheetNameInvalid);

            // Add header to the sheets
            Row titleRowMale = sheetMale.createRow(0);
            Row titleRowFemale = sheetFemale.createRow(0);
            Row titleRowInvalid = sheetInvalid.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell titleCellMale = titleRowMale.createCell(i, CellType.STRING);
                Cell titleCellFemale = titleRowFemale.createCell(i, CellType.STRING);
                Cell titleCellInvalid = titleRowInvalid.createCell(i, CellType.STRING);

                titleCellMale.setCellValue(header[i]);
                titleCellFemale.setCellValue(header[i]);
                titleCellInvalid.setCellValue(header[i]);
            }

            String[] row;
            int rowCount = 0;
            while ((row = reader.readNext()) != null) {
                if (!isValid(row)) {
                    addRowToSheet(sheetInvalid, row);
                    continue;
                }

                if ("Male".equalsIgnoreCase(row[4])) {
                    addRowToSheet(sheetMale, row);
                } else if ("Female".equalsIgnoreCase(row[4])) {
                    addRowToSheet(sheetFemale, row);
                } else {
                    addRowToSheet(sheetInvalid, row);
                }

                rowCount++;
                if (rowCount % CHUNK_SIZE == 0) {
                    writeWorkbookToFile(workbook, excelFilePath);
                    workbook = new XSSFWorkbook();
                    sheetMale = workbook.createSheet(sheetNameMale);
                    sheetFemale = workbook.createSheet(sheetNameFemale);
                    sheetInvalid = workbook.createSheet(sheetNameInvalid);
                    titleRowMale = sheetMale.createRow(0);
                    titleRowFemale = sheetFemale.createRow(0);
                    titleRowInvalid = sheetInvalid.createRow(0);
                    for (int i = 0; i < header.length; i++) {
                        Cell titleCellMale = titleRowMale.createCell(i, CellType.STRING);
                        Cell titleCellFemale = titleRowFemale.createCell(i, CellType.STRING);
                        Cell titleCellInvalid = titleRowInvalid.createCell(i, CellType.STRING);

                        titleCellMale.setCellValue(header[i]);
                        titleCellFemale.setCellValue(header[i]);
                        titleCellInvalid.setCellValue(header[i]);
                    }
                }
            }

        } catch (CsvException e) {
            throw new IOException("Error reading or validating CSV file.", e);
        }

        writeWorkbookToFile(workbook, excelFilePath);
    }

    private static void addRowToSheet(Sheet sheet, String[] row) {
        int rowNum = sheet.getLastRowNum() + 1;
        Row dataRow = sheet.createRow(rowNum);
        for (int i = 0; i < row.length; i++) {
            Cell cell = dataRow.createCell(i, CellType.STRING);
            cell.setCellValue(row[i]);
        }
    }

    private static void writeWorkbookToFile(Workbook workbook, String excelFilePath) throws IOException {
        try (FileOutputStream fileOut = new FileOutputStream(excelFilePath, true)) {
            workbook.write(fileOut);
        }
    }

    private static String formatExecutionTime(double seconds) {
        if (seconds >= 60) {
            long minutes = (long) (seconds / 60);
            long remainingSeconds = (long) (seconds % 60);
            return String.format("%d minutes %d seconds", minutes, remainingSeconds);
        } else {
            return String.format("%.2f seconds", seconds);
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

        return !id.isEmpty() && !name.isEmpty() && !mobileNumber.isEmpty() && !emailAddress.isEmpty() && !gender.isEmpty() && Pattern.matches(ID_REGEX, id) && Pattern.matches(MOBILE_REGEX, mobileNumber) && Pattern.matches(EMAIL_REGEX, emailAddress);
    }
}
