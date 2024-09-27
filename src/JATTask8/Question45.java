package JATTask8;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Question45 {

	public static void main(String[] args) {
		String excelFilePath = "Task8.xlsx";

        // Write data to Excel file
        writeDataToExcel("C:\\Users\\pc\\OneDrive\\Desktop\\Task8.xlsx");

       
    }

    // Method to write data to an Excel file
    public static void writeDataToExcel(String filePath) {
        Workbook workbook = new XSSFWorkbook(); // Create a new workbook
        Sheet sheet = workbook.createSheet("Sheet1"); // Create a new sheet

        // Create header row
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Name");
        headerRow.createCell(1).setCellValue("Age");
        headerRow.createCell(2).setCellValue("Email");

        // Add rows with data
        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue("John Doe");
        row1.createCell(1).setCellValue(30);
        row1.createCell(2).setCellValue("john@test.com");

        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue("Jane Doe");
        row2.createCell(1).setCellValue(28);
        row2.createCell(2).setCellValue("jane@test.com");

        // Auto-size the columns
        for (int i = 0; i < 3; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the workbook to a file
        try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
            System.out.println("Data written to Excel file successfully.");
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



