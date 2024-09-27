package JATTask8;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class Question123 {

	public static void main(String[] args) {
		 // Creating a new workbook
        Workbook workbook = new XSSFWorkbook();

        // Creating a new sheet called "Sheet1"
        Sheet sheet = workbook.createSheet("Sheet1");

        // Creating the header row and setting the headers
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Name");
        headerRow.createCell(1).setCellValue("Age");
        headerRow.createCell(2).setCellValue("Email");

        // Adding first data row
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("John Doe");
        row1.createCell(1).setCellValue(30);
        row1.createCell(2).setCellValue("john@test.com");

        // Adding second data row
        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Jane Doe");
        row2.createCell(1).setCellValue(28);
        row2.createCell(2).setCellValue("jane@test.com");

        // Adjust column widths to fit content
        for (int i = 0; i < 3; i++) {
            sheet.autoSizeColumn(i);
        }

        // Writing the data to an Excel file
        try (FileOutputStream fileOut = new FileOutputStream("priya.xlsx")) {
            workbook.write(fileOut);
            System.out.println("Excel file has been generated successfully!");
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

	}

}
