package eu.smartpeg;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.util.Properties;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class OracleToExcel
{
    public static void main( String[] args )
    {
        System.out.println( "Hello World!" );

        System.out.println("Lettura file di configurazione");


        String jdbcUrl = "";
        String username = "";
        String password = "";

        // SQL query
        String sqlQuery = readQueryFromFile("query.sql");

        Properties properties = new Properties();

        try (FileInputStream fis = new FileInputStream("config.properties")) {
            properties.load(fis);

            jdbcUrl = properties.getProperty("jdbcUrl");
            username = properties.getProperty("username");
            password = properties.getProperty("password");

            System.out.println("jdbcUrl : " + jdbcUrl);
            System.out.println("username: " + username);
            System.out.println("password: " + password);

        } catch (IOException e) {
            System.err.println("An error occurred while reading the properties file: " + e.getMessage());
        }

        // Excel file details
        String excelFilePath = "anagrafico.xlsx";
        String sheetName = "Dati1";

        System.out.println("Esecuzione query");

        try (Connection conn = DriverManager.getConnection(jdbcUrl, username, password);
             Statement stmt = conn.createStatement();
             ResultSet rs = stmt.executeQuery(sqlQuery)) {
            System.out.println("Connessione DB OK");

            // Create a new workbook and sheet
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet(sheetName);

            // Create header row
            ResultSetMetaData metaData = rs.getMetaData();
            Row headerRow = sheet.createRow(0);
            int columnCount = metaData.getColumnCount();
            for (int i = 1; i <= columnCount; i++) {
                String columnName = metaData.getColumnName(i);
                Cell headerCell = headerRow.createCell(i - 1);
                headerCell.setCellValue(columnName);
            }

            // Populate data rows
            int rowNum = 1;
            while (rs.next()) {
                Row dataRow = sheet.createRow(rowNum);
                for (int i = 1; i <= columnCount; i++) {
                    Object value = rs.getObject(i);
                    Cell dataCell = dataRow.createCell(i - 1);
                    if (value != null) {
                        dataCell.setCellValue(value.toString());
                    }
                }
                rowNum++;
            }

            // Autosize columns
            for (int i = 0; i < columnCount; i++) {
                sheet.autoSizeColumn(i);
            }

            // Save the workbook to file
            FileOutputStream fileOut = new FileOutputStream(excelFilePath);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();

            System.out.println("Excel file created successfully!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Read query from file
     * @param s file name
     * @return query
     */
    private static String readQueryFromFile(String s) {
        String content = "";
        try {
            content = new String(Files.readAllBytes(Paths.get(s)));
            System.out.println(content);
        } catch (IOException e) {
            System.err.println("An error occurred while reading the file: " + e.getMessage());
        }
        return content;
    }
}
