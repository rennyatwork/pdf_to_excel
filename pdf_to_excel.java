import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class PdfToExcelConverter {

    public static void main(String[] args) {
        String inputPdfPath = "path/to/input.pdf";
        String outputExcelPath = "path/to/output.xlsx";

        try {
            PDDocument document = PDDocument.load(new File(inputPdfPath));
            PDFTextStripper textStripper = new PDFTextStripper();
            String pdfText = textStripper.getText(document);

            List<String[]> tableData = extractTableData(pdfText);

            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("TableSheet");

            int rowNum = 0;
            for (String[] rowData : tableData) {
                Row row = sheet.createRow(rowNum++);
                for (int colNum = 0; colNum < rowData.length; colNum++) {
                    Cell cell = row.createCell(colNum);
                    String cellContent = rowData[colNum].replace("\r\n", " ");  // Replace newline with space
                    cell.setCellValue(cellContent);
                }
            }

            try (FileOutputStream outputStream = new FileOutputStream(outputExcelPath)) {
                workbook.write(outputStream);
            }

            document.close();
            System.out.println("PDF to Excel conversion completed.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<String[]> extractTableData(String pdfText) {
        List<String[]> tableData = new ArrayList<>();
        // Implement logic to extract table data from pdfText
        // This could involve splitting and parsing the text based on the table structure
        // For the purpose of this example, let's assume we have a list of arrays representing rows
        // Each array contains cell data as strings
        // Replace this with your actual logic to extract table data
        return tableData;
    }
}
