package JATTask8;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;

public class Question5 {

	public static void main(String[] args) {
		// Specify the path of the existing Excel file
        String excelFilePath = "C:\\Users\\pc\\OneDrive\\Desktop\\Task8.xlsx";

        // Read data from Excel file
        readDataFromExcel(excelFilePath);
    }

    // Method to read data from an Excel file
    public static void readDataFromExcel(String filePath) {
        try (FileInputStream fileInputStream = new FileInputStream(new String("C:\\Users\\pc\\OneDrive\\Desktop\\Task8.xlsx"))) {
            // Open the Excel file as a workbook
            Workbook workbook = new XSSFWorkbook(fileInputStream);
            // Get the first sheet
            Sheet sheet = workbook.getSheetAt(0);

            // Iterate through each row
            for (Row row : sheet) {
                // Iterate through each cell in the row
                for (Cell cell : row) {
                    // Check the cell type and print the appropriate value
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            System.out.print((int) cell.getNumericCellValue() + "\t");
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "\t");
                            break;
                        case FORMULA:
                            System.out.print(cell.getCellFormula() + "\t");
                            break;
                        default:
                            System.out.print("Unknown Value" + "\t");
                            break;
                    }
                }
                System.out.println(); // New line after each row
            }
            // Close the workbook after reading
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

	}

}
