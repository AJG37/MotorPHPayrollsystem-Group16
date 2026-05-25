# MotorPH Payroll System (Group 16)

## Project Description
MotorPH Payroll System is a Java Swing desktop application for managing employee profiles and computing payroll summaries. The app uses a local Excel workbook as its data store and provides role-based access for administrators and employees. Administrators can manage employee records and generate company-wide payroll reports, while employees can view their own profiles and payslips.

## Features
- Role-based login (Administrator and Employee)
- Employee profile and roster viewing
- Payroll report generation with period filters
- Simple leave request submission (UI-only)
- Excel-backed data storage

## Setup Instructions
1. Install Java (JDK 8+ recommended) and Maven.
2. Open the project in your IDE (e.g., VS Code, IntelliJ, Eclipse).
3. Ensure the Excel database file is in the project root:
	- MotorPH_EmployeeData.xlsx
4. Verify the Excel workbook contains the required sheets:
	- Employee Details
	- Attendance Record

## Usage
### Run via Maven
From the MotorPH folder:
```
mvn clean compile
mvn exec:java -Dexec.mainClass=MotorPH
```

### Run via IDE
1. Open MotorPH/src/main/java/MotorPH.java.
2. Run the MotorPH class.

### Default Admin Login
- Employee ID: 999
- Password: admin123

## Technologies Used
- Java (Swing)
- Apache POI (Excel read/write)
- Maven

## Group Members and Roles
- Xyrus Ezekiel Cuenca — UI/UX Designer & Developer
- Albert Joaquin Geronimo — UI/UX Designer & Developer
- Elefes Ramones Capulong — Documentation Specialist & Project Manager
- Euie Garcia — Frontend Developer & Documentation

## Data Format Notes
- The application expects specific column order in the Excel workbook. Columns referenced include:
  - Employee Details: ID (0), Last Name (1), First Name (2), Status (10), Position (11), Basic Salary (13), Hourly Rate (18)
  - Attendance Record: ID (0), Date (3), Time In (4), Time Out (5)
- If the hourly rate cell appears to be monthly, it is normalized using a 160-hour month.

## Troubleshooting
- If the app cannot find the Excel file, place MotorPH_EmployeeData.xlsx in the project root and retry.
- If payroll results look empty, verify the Attendance Record sheet dates match the MM/YYYY filter format.

## License
This project is for academic use and internal learning purposes.