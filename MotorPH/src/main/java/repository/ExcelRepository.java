package repository;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRepository {
    private static final String EXCEL_FILE_PATH = "MotorPH_EmployeeData.xlsx";
    private static final DataFormatter FORMATTER = new DataFormatter();

    public static boolean saveNewEmployee(String id, String fName, String lName) {
        try (FileInputStream fis = new FileInputStream(new File(EXCEL_FILE_PATH));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("Employee Details");
            int lastRow = sheet.getLastRowNum();
            Row newRow = sheet.createRow(lastRow + 1);

            newRow.createCell(0).setCellValue(id);
            newRow.createCell(1).setCellValue(lName);
            newRow.createCell(2).setCellValue(fName);
            newRow.createCell(10).setCellValue("Probationary");

            try (FileOutputStream fos = new FileOutputStream(new File(EXCEL_FILE_PATH))) {
                workbook.write(fos);
            }
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    public static boolean updateEmployeeStatus(String id, String newStatus) {
        try (FileInputStream fis = new FileInputStream(new File(EXCEL_FILE_PATH));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("Employee Details");
            for (Row row : sheet) {
                if (getCellValueAsString(row.getCell(0)).equals(id)) {
                    Cell statusCell = row.getCell(10);
                    if (statusCell == null) statusCell = row.createCell(10);
                    statusCell.setCellValue(newStatus);
                    break;
                }
            }
            try (FileOutputStream fos = new FileOutputStream(new File(EXCEL_FILE_PATH))) {
                workbook.write(fos);
            }
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    public static boolean removeEmployeeRecord(String id) {
        try (FileInputStream fis = new FileInputStream(new File(EXCEL_FILE_PATH));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("Employee Details");
            int rowIndex = -1;
            for (Row row : sheet) {
                if (getCellValueAsString(row.getCell(0)).equals(id)) {
                    rowIndex = row.getRowNum();
                    break;
                }
            }
            if (rowIndex != -1) {
                sheet.removeRow(sheet.getRow(rowIndex));
                try (FileOutputStream fos = new FileOutputStream(new File(EXCEL_FILE_PATH))) {
                    workbook.write(fos);
                }
                return true;
            }
            return false;
        } catch (Exception e) {
            return false;
        }
    }

    public static boolean checkEmployeeExists(String id) {
        try (FileInputStream fis = new FileInputStream(new File(EXCEL_FILE_PATH));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("Employee Details");
            for (Row row : sheet) {
                if (getCellValueAsString(row.getCell(0)).equals(id)) {
                    return true;
                }
            }
        } catch (Exception e) {}
        return false;
    }

    public static double getNumericSafe(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null) {
            return 0.0;
        }
        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return cell.getNumericCellValue();
            } else if (cell.getCellType() == CellType.FORMULA) {
                return evaluator.evaluate(cell).getNumberValue();
            }
            String val = FORMATTER.formatCellValue(cell, evaluator).replace(",", "").replaceAll("[^\\d.]", "");
            if (val.isEmpty()) {
                return 0.0;
            } else {
                return Double.parseDouble(val);
            }
        } catch (Exception e) {
            return 0.0;
        }
    }

    public static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        return FORMATTER.formatCellValue(cell).trim();
    }
}
