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

// Handles employee data operations stored in the MotorPH Excel workbook.
public class ExcelRepository {
    // The Excel file is expected in the program's working directory.
    private static final String EXCEL_FILE_PATH = "MotorPH_EmployeeData.xlsx";
    private static final DataFormatter FORMATTER = new DataFormatter();

    // Adds a new employee row with a default probationary status.
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

    // Updates the employment status of the employee with the given ID.
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

    // Deletes the employee row that matches the given ID.
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

    // Checks whether an employee ID exists in the workbook.
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

    // Builds an HTML table containing the employee roster.
    public static String getAllEmployeeProfilesString() {
        StringBuilder sb = new StringBuilder();

        try (FileInputStream fis = new FileInputStream(new File(EXCEL_FILE_PATH));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("Employee Details");

            sb.append("<h2 style='color: #4DA6FF; border-bottom: 1px solid #4DA6FF; padding-bottom: 5px;'>Company Employee Roster</h2>");
            sb.append("<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse; width: 100%; border-color: #555;'>");
            sb.append("<tr style='background-color: #2C2C2C; color: #4DA6FF;'>");

            sb.append("<th>ID Number</th><th>Full Name</th><th>Position</th><th>Status</th>");
            sb.append("</tr>");

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                String id = getCellValueAsString(row.getCell(0));
                String name = getCellValueAsString(row.getCell(2)) + " " + getCellValueAsString(row.getCell(1));
                String position = getCellValueAsString(row.getCell(11));
                String status = getCellValueAsString(row.getCell(10));

                if (!id.isEmpty()) {
                    sb.append("<tr style='background-color: #1E1E1E;'>");
                    sb.append("<td style='text-align: center;'>").append(id).append("</td>");
                    sb.append("<td>").append(name).append("</td>");
                    sb.append("<td>").append(position).append("</td>");
                    sb.append("<td style='text-align: center;'>").append(status).append("</td>");
                    sb.append("</tr>");
                }
            }
            sb.append("</table>");
        } catch (Exception e) {
            sb.append("<p style='color:red;'>Error retrieving company roster: ").append(e.getMessage()).append("</p>");
        }
        return sb.toString();
    }

    // Builds an HTML profile summary for one employee.
    public static String getEmployeeProfileString(String searchId) {
        StringBuilder sb = new StringBuilder();
        try (FileInputStream fis = new FileInputStream(new File(EXCEL_FILE_PATH));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("Employee Details");

            Row myRow = null;
            for (Row r : sheet) {
                if (getCellValueAsString(r.getCell(0)).equals(searchId)) {
                    myRow = r;
                    break;
                }
            }

            if (myRow != null) {
                String idNum      = getCellValueAsString(myRow.getCell(0));
                String firstName  = getCellValueAsString(myRow.getCell(1));
                String lastName   = getCellValueAsString(myRow.getCell(2));
                String fullName   = lastName + " " + firstName;
                String birthday   = getCellValueAsString(myRow.getCell(3));
                String address    = getCellValueAsString(myRow.getCell(4));
                String phone      = getCellValueAsString(myRow.getCell(5));

                String sss        = getCellValueAsString(myRow.getCell(6));
                String philHealth = getCellValueAsString(myRow.getCell(7));
                String tin        = getCellValueAsString(myRow.getCell(8));
                String pagIbig    = getCellValueAsString(myRow.getCell(9));

                String status     = getCellValueAsString(myRow.getCell(10));
                String position   = getCellValueAsString(myRow.getCell(11));

                sb.append("<h2 style='color: #4DA6FF; border-bottom: 1px solid #4DA6FF; padding-bottom: 5px;'>Employee Profile Record</h2>");
                sb.append("<table border='0' cellpadding='6' style='font-size: 14px;'>");
                sb.append("<tr><td style='color: #888;'>ID Number:</td><td><b>" + idNum + "</b></td></tr>");
                sb.append("<tr><td style='color: #888;'>Full Name:</td><td><b>" + fullName + "</b></td></tr>");
                sb.append("<tr><td style='color: #888;'>Birthday:</td><td>" + birthday + "</td></tr>");
                sb.append("<tr><td style='color: #888;'>Address:</td><td>" + address + "</td></tr>");
                sb.append("<tr><td style='color: #888;'>Phone:</td><td>" + phone + "</td></tr>");
                sb.append("<tr><td colspan='2'><h3 style='color: #4DA6FF; margin-top: 15px;'>Government & Tax IDs</h3></td></tr>");
                sb.append("<tr><td style='color: #888;'>SSS Number:</td><td>" + sss + "</td></tr>");
                sb.append("<tr><td style='color: #888;'>PhilHealth No:</td><td>" + philHealth + "</td></tr>");
                sb.append("<tr><td style='color: #888;'>TIN:</td><td>" + tin + "</td></tr>");
                sb.append("<tr><td style='color: #888;'>Pag-IBIG No:</td><td>" + pagIbig + "</td></tr>");
                sb.append("<tr><td colspan='2'><h3 style='color: #4DA6FF; margin-top: 15px;'>Employment Status</h3></td></tr>");
                sb.append("<tr><td style='color: #888;'>Status:</td><td>" + status + "</td></tr>");
                sb.append("<tr><td style='color: #888;'>Position:</td><td><b>" + position + "</b></td></tr>");
                sb.append("</table>");
            } else {
                sb.append("<p style='color:red;'>Employee not found.</p>");
            }
        } catch (Exception e) {
            sb.append("<p style='color:red;'>Error retrieving profile: ").append(e.getMessage()).append("</p>");
        }
        return sb.toString();
    }

    // Reads a numeric cell safely, including formula and text values.
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

    // Converts a cell value to a trimmed string for comparisons and display.
    public static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        return FORMATTER.formatCellValue(cell).trim();
    }
}
