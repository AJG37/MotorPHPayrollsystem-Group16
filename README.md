# MotorPH Payroll System (Group 16)

## Project Description

MotorPH Payroll System is a beginner-friendly Java Swing desktop application for managing employee records and calculating payroll reports. It uses a local Excel workbook as its data store and provides separate administrator and employee views.

Administrators can view, add, update, and delete employee records; generate individual payroll reports; and process a company-wide payroll summary. Employees can view their own profiles and payslips.

## Features

- Administrator and employee login views
- Employee profile and company roster reports
- Add employee with ID, duplicate, and name validation
- Update first name, last name, status, or position
- Delete employee with confirmation
- Individual payroll reports with period filters
- Company-wide payroll summary with:
  - Total Employees Processed
  - Total Gross Pay
  - Total Deductions
  - Total Net Pay
  - Average Net Pay
  - Unified employee table showing Employee Number, Name, Rate, Hours Worked, Gross Pay, Deductions, and Net Pay
- Loading message while reports are prepared
- Excel file, sheet, and column-order validation
- Simple leave request submission (UI-only)

## Requirements and Setup

1. Install JDK 21 and Maven. The existing `pom.xml` targets Java 21.
2. Open the repository in VS Code, IntelliJ IDEA, Eclipse, or another Java IDE.
3. Use the `MotorPH` folder as the application's working directory.
4. Keep `MotorPH_EmployeeData.xlsx` inside the `MotorPH` folder.
5. Do not rename the workbook, required sheets, or required headers.

### Run via Maven

From the `MotorPH` folder:

```text
mvn clean compile
mvn exec:java -Dexec.mainClass=MotorPH
```

### Run via an IDE

1. Open `MotorPH/src/main/java/MotorPH.java`.
2. Set the working directory to the `MotorPH` folder.
3. Run the `MotorPH` class.

### Default Administrator Login

- Employee ID: `999`
- Password: `admin123`

The Admin Password field is required only for administrator login. Employee login currently uses the employee ID stored in the workbook.

## Excel Workbook Format

The application expects `MotorPH_EmployeeData.xlsx` in its working directory. It validates the required sheets and headers before running file-dependent actions.

### Required Sheet: Employee Details

The headers must remain in this exact order:

1. Employee #
2. Last Name
3. First Name
4. Birthday
5. Address
6. Phone Number
7. SSS #
8. Philhealth #
9. TIN #
10. Pag-ibig #
11. Status
12. Position
13. Immediate Supervisor
14. Basic Salary
15. Rice Subsidy
16. Phone Allowance
17. Clothing Allowance
18. Gross Semi-monthly Rate
19. Hourly Rate

### Required Sheet: Attendance Record

The headers must remain in this exact order:

1. Employee #
2. Last Name
3. First Name
4. Date
5. Log In
6. Log Out

The application still uses fixed column positions after validating these headers. It does not automatically rearrange or remap columns.

## Validation and Error Handling

- Employee IDs entered in Add, Update, Delete, and single-payroll flows must contain exactly five digits.
- Duplicate employee IDs are rejected.
- First and last names accept letters, spaces, hyphens, and apostrophes.
- Status accepts only `Regular` or `Probationary`, case-insensitively.
- Position accepts letters, spaces, hyphens, apostrophes, and ampersands.
- Custom payroll periods must use `MM/YYYY`, with a month from `01` to `12`.
- Missing, damaged, open-locked, or incorrectly structured workbooks produce clearer guidance.
- Payroll reports clearly explain when the selected period has no usable records.

## Revision Notes

### Payroll and Employee Features

- Added the company-wide payroll summary required for Feature 5.
- Kept individual payroll generation separate from the bulk report.
- Expanded employee updates to safe text fields only: first name, last name, status, and position.
- Intentionally excluded salary, hourly rate, allowances, employee ID, and attendance editing from the Update screen to protect payroll data and formulas.

### Usability Improvements

- Added loading dialogs for roster, profile, individual payroll, and company payroll reports.
- Added Enter-key support to the login form.
- Added an Admin Password field directly below Employee ID instead of using a separate password pop-up.
- Improved add, update, delete, file, period, employee-not-found, and no-records messages.
- Improved the dark theme using clearer panel, field, border, button, report, and semantic status colors.

### Code and Data Handling

- Added `PayrollService` for payroll calculation methods.
- Added `ExcelRepository` for Excel-based employee operations and workbook validation.
- Added model classes for employee, attendance, and payroll-result data.
- Removed duplicate helper and CRUD logic from the main class during the earlier refactor.
- Added comments for intentionally skipped malformed attendance rows.
- Completed a Java 21 strict compiler pass without warnings.

## Recommended Manual Tests

Before submission, test the following with a backup copy of the workbook:

1. Administrator login using the visible password field and Enter key.
2. Employee login and profile viewing.
3. Add a valid employee.
4. Reject symbols or numbers in employee names.
5. Reject a duplicate employee ID.
6. Update first name, last name, status, and position individually.
7. Cancel a deletion, then delete a temporary test employee.
8. Generate single payroll for All Time, Latest Month, Previous Month, and a custom period.
9. Generate the company-wide payroll summary and verify its totals and employee table.
10. Test an invalid custom period such as `13/2024`.
11. Temporarily rename the workbook and verify the missing-file message.
12. Use a backup workbook with a missing sheet or changed header and verify the format message.

## Known Limitations

- This is an academic desktop application, not a production payroll or security system.
- Employee login uses an employee ID without an employee-specific password.
- Administrator credentials are stored in the source code.
- The Excel workbook is a local file and does not support concurrent multi-user editing. Close it in Excel before the application writes changes.
- Sheet names and column order are fixed and must match the documented format.
- Adding an employee creates the ID, name, and default `Probationary` status. Remaining personal and payroll fields are not entered through the Add screen.
- Salary, hourly rate, allowances, IDs, and attendance records are not editable through the Update screen.
- Employees without usable attendance for the selected period are skipped by the company payroll summary.
- Payroll rules are simplified for the assessment and should be reviewed before any real-world use.
- Leave requests display a confirmation but are not persisted or sent to an administrator.

## Technologies Used

- Java 21
- Java Swing
- Apache POI for Excel read/write operations
- Maven

## Group Members and Roles

- Xyrus Ezekiel Cuenca — UI/UX Designer & Developer
- Albert Joaquin Geronimo — UI/UX Designer & Developer
- Elefes Ramones Capulong — Documentation Specialist & Project Manager
- Euie Garcia — Frontend Developer & Documentation

## License

This project is for academic use and internal learning purposes.
