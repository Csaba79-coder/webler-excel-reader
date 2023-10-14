import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelReaderApp {

    // first step add dependency to pom.xml
    // https://mvnrepository.com/artifact/org.apache.poi/poi/5.2.4
    // the newest version does not work :) find a previous!!!

    public static void main(String[] args) {

        String filePath = "src/main/resources/webler.xlsx";

        try {
            Workbook workbook = openExcelFile(filePath);
            String line = readExcelFile(workbook);
            String folderPath = createFolder();
            String file = folderPath + "/output.txt";
            writeContentToFile(line, file);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }

    private static Workbook openExcelFile(String filePath) throws IOException {
        File file = new File(filePath);
        FileInputStream inputStream = new FileInputStream(file);
        return new XSSFWorkbook(inputStream);
    }

    private static String readExcelFile(Workbook workbook) {
        StringBuilder content = new StringBuilder();
        Sheet sheet = workbook.getSheetAt(0); // Assuming you want the first sheet

        for (Row row : sheet) {
            for (Cell cell : row) {
                content.append(cell.toString()).append("\t"); // Print cell content
                // System.out.print(cell.toString() + "\t"); // Print cell content
            }
            content.append("\n"); // Move to the next row
            // System.out.println(); // Move to the next row
        }
        return content.toString();
    }

    private static String createFolder() {
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd-MM-yyyy");
        String date = dateFormat.format(new Date());
        File directory = new File("src/main/resources/" + date);

        if (!directory.exists()) {
            directory.mkdir();
        }
        return directory.getPath();
    }

    private static void writeContentToFile(String line, String file) {
        try (BufferedWriter writer = new BufferedWriter(new FileWriter(file))) {
            writer.write(line);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
