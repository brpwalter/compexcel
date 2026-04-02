package compexcel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelComparator {

    public static void main(String[] args) {
        String file1 = "old/file1.xlsx";
        String file2 = "new/file2.xlsx";
        String outputFile = "diff.xlsx";

        try (Workbook wb1 = new XSSFWorkbook(new FileInputStream(file1));
             Workbook wb2 = new XSSFWorkbook(new FileInputStream(file2));
             Workbook resultWb = new XSSFWorkbook()) {

            // Style für Markierung von Unterschieden definieren
            CellStyle diffStyle = resultWb.createCellStyle();
            diffStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            diffStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            // Alle Sheets aus Datei 1 durchlaufen
            for (int i = 0; i < wb1.getNumberOfSheets(); i++) {
                Sheet sheet1 = wb1.getSheetAt(i);
                Sheet sheet2 = wb2.getSheet(sheet1.getSheetName());

                if (sheet2 == null) {
                    System.out.println("Sheet '" + sheet1.getSheetName() + "' fehlt in Datei 2.");
                    continue;
                }

                Sheet resSheet = resultWb.createSheet(sheet1.getSheetName());
                compareSheets(sheet1, sheet2, resSheet, diffStyle);
            }

            try (FileOutputStream out = new FileOutputStream(outputFile)) {
                resultWb.write(out);
            }
            System.out.println("Vergleich abgeschlossen. Ergebnisse in: " + outputFile);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void compareSheets(Sheet s1, Sheet s2, Sheet res, CellStyle diffStyle) {
        int maxRows = Math.max(s1.getLastRowNum(), s2.getLastRowNum());

        for (int i = 0; i <= maxRows; i++) {
            Row r1 = s1.getRow(i);
            Row r2 = s2.getRow(i);
            Row resRow = res.createRow(i);

            if (r1 == null && r2 == null) continue;

            int maxCols = Math.max(
                r1 != null ? r1.getLastCellNum() : 0, 
                r2 != null ? r2.getLastCellNum() : 0
            );

            for (int j = 0; j < maxCols; j++) {
                Cell c1 = (r1 != null) ? r1.getCell(j) : null;
                Cell c2 = (r2 != null) ? r2.getCell(j) : null;
                Cell resCell = resRow.createCell(j);

                String val1 = getCellValueAsString(c1);
                String val2 = getCellValueAsString(c2);

                if (!val1.equals(val2)) {
                    resCell.setCellValue(val1 + " vs " + val2);
                    resCell.setCellStyle(diffStyle);
                } else {
                    resCell.setCellValue(val1);
                }
            }
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }
}
