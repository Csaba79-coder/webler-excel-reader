import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class Main {

    // first step add dependency to pom.xml
    // https://mvnrepository.com/artifact/org.apache.poi/poi/5.2.4
    // the newest version does not work :) find a previous!!!

    public static void main(String[] args) {

        String filePath = "src/main/resources/webler.xlsx";

        File file = new File(filePath);

       try (FileInputStream inputStream = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(inputStream)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming you want the first sheet

            for (Row row : sheet) {
             for (Cell cell : row) {
                  System.out.print(cell.toString() + "\t"); // Print cell content
             }
             System.out.println(); // Move to the next row
            }


        } catch (IOException e) {
            System.out.println("Error reading excel file: " + e.getMessage());
        }
    }
}
