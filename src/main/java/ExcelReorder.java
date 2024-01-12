import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class ExcelReorder {
    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\alexa\\Downloads/Alex.xlsx";  // Path to your input Excel file
        String referenceFilePath = "C:\\Users\\alexa\\Downloads/Final2222.xlsx";  // Path to your reference Excel file
        String outputFilePath = "D:\\Power Bi/output.xlsx";  // Path to the output Excel file

        try {
            // Load input and reference Excel files
            FileInputStream inputWorkbook = new FileInputStream(inputFilePath);
            FileInputStream referenceWorkbook = new FileInputStream(referenceFilePath);

            Workbook inputWorkbookObj = new XSSFWorkbook(inputWorkbook);
            Workbook referenceWorkbookObj = new XSSFWorkbook(referenceWorkbook);

            // Get the sheets from both workbooks
            Sheet inputSheet = inputWorkbookObj.getSheetAt(0);
            Sheet referenceSheet = referenceWorkbookObj.getSheetAt(0);

            // Get titles from the input sheet
            List<String> inputTitles = getTitles(inputSheet);

            // Get titles from the reference sheet
            List<String> referenceTitles = getTitles(referenceSheet);

            // Create a map to store the order of titles based on the reference
            Map<String, Integer> titleOrderMap = new HashMap<>();
            int order = 1;
            for (String title : referenceTitles) {
                titleOrderMap.put(title, order++);
            }

            // Sort the input titles based on the order in the reference, with a default value for titles not found
            inputTitles.sort(Comparator.comparingInt(title -> titleOrderMap.getOrDefault(title, Integer.MAX_VALUE)));



            // Create a new workbook for the output
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet();

            // Write sorted titles to the output sheet
            writeTitlesToSheet(outputSheet, inputTitles);

            // Write the output workbook to a file
            try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
                outputWorkbook.write(outputStream);
            }

            System.out.println("Excel file has been created successfully!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<String> getTitles(Sheet sheet) {
        List<String> titles = new ArrayList<>();
        Row headerRow = sheet.getRow(0);

        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell != null) {
                titles.add(cell.getStringCellValue());
            }
        }

        return titles;
    }

    private static void writeTitlesToSheet(Sheet sheet, List<String> titles) {
        Row row = sheet.createRow(0);

        for (int i = 0; i < titles.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(titles.get(i));
        }
    }
}