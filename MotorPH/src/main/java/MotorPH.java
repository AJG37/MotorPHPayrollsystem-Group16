/*
 * Project: MotorPH Payroll System
 * Description: An application for managing employee data and calculating payroll.
 * It features a login portal with access levels for Administrators and Employees.
 * All data is read from and written to a local Excel database.
 */

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import java.awt.*;
import java.awt.Font;  
import java.awt.Color; 
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream; 
import java.time.Duration;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

public class MotorPH extends JFrame {
    
    private CardLayout cardLayout;
    private JPanel mainContainer;
    private String currentLoggedInEmployeeId = "";

    private static final String ADMIN_ID = "999"; 
    private static final String ADMIN_PASS = "admin123";
    private static final String EXCEL_FILE_PATH = "MotorPH_EmployeeData.xlsx";
    private static final DataFormatter FORMATTER = new DataFormatter();
    
    private static final Color BG_DARK = new Color(18, 18, 18);
    private static final Color PANEL_DARK = new Color(30, 30, 30);
    private static final Color TEXT_LIGHT = new Color(240, 240, 240);
    private static final Color BTN_ACCENT = new Color(70, 70, 70);
    
    private static final String HTML_HEADER = "<html><body style='font-family: Arial, sans-serif; background-color: #1E1E1E; color: #F0F0F0; margin: 10px;'>";
    private static final String HTML_FOOTER = "</body></html>";

    public MotorPH() {
        setTitle("MotorPH Payroll System");
        
        setExtendedState(JFrame.MAXIMIZED_BOTH); 
        
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setResizable(true); 
        
        UIManager.put("OptionPane.background", PANEL_DARK);
        UIManager.put("Panel.background", PANEL_DARK);
        UIManager.put("OptionPane.messageForeground", TEXT_LIGHT);

        cardLayout = new CardLayout();
        mainContainer = new JPanel(cardLayout);

        mainContainer.add(createLoginScreen(), "LOGIN");
        mainContainer.add(createAdminDashboard(), "ADMIN_DASHBOARD");
        mainContainer.add(createEmployeeDashboard(), "EMPLOYEE_DASHBOARD");

        add(mainContainer);
        cardLayout.show(mainContainer, "LOGIN");
    }

    private JPanel createLoginScreen() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBackground(BG_DARK); 
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(10, 10, 10, 10);

        JLabel titleLabel = new JLabel("MotorPH Secure Portal");
        titleLabel.setFont(new Font("Arial", Font.BOLD, 26));
        titleLabel.setForeground(TEXT_LIGHT);
        gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 2;
        panel.add(titleLabel, gbc);

        JLabel idLabel = new JLabel("Enter Employee ID:");
        idLabel.setForeground(TEXT_LIGHT);
        gbc.gridwidth = 1; gbc.gridy = 1;
        panel.add(idLabel, gbc);
        
        JTextField idField = new JTextField(15);
        idField.setFont(new Font("Arial", Font.PLAIN, 16));
        idField.setBackground(PANEL_DARK);
        idField.setForeground(TEXT_LIGHT);
        idField.setCaretColor(Color.WHITE);
        gbc.gridx = 1;
        panel.add(idField, gbc);

        JButton loginBtn = new JButton("Login");
        loginBtn.setFont(new Font("Arial", Font.BOLD, 14));
        loginBtn.setBackground(BTN_ACCENT);
        loginBtn.setForeground(TEXT_LIGHT);
        loginBtn.setFocusPainted(false);
        gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 2;
        panel.add(loginBtn, gbc);

        loginBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    String userId = idField.getText().trim();

                    if (userId.isEmpty()) {
                        throw new IllegalArgumentException("Employee ID cannot be blank!");
                    }

                    if (userId.equals(ADMIN_ID)) {
                        JPasswordField pf = new JPasswordField();
                        pf.setBackground(PANEL_DARK);
                        pf.setForeground(TEXT_LIGHT);
                        pf.setCaretColor(Color.WHITE);
                        
                        int okCxl = JOptionPane.showConfirmDialog(panel, pf, "Enter Admin Password", JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
                        
                        if (okCxl == JOptionPane.OK_OPTION) {
                            String pass = new String(pf.getPassword());
                            if (pass.equals(ADMIN_PASS)) {
                                currentLoggedInEmployeeId = ADMIN_ID; 
                                idField.setText(""); 
                                cardLayout.show(mainContainer, "ADMIN_DASHBOARD"); 
                            } else {
                                throw new SecurityException("Incorrect Admin Password.");
                            }
                        }
                    } else {
                        if (checkEmployeeExists(userId)) {
                            currentLoggedInEmployeeId = userId; 
                            idField.setText(""); 
                            cardLayout.show(mainContainer, "EMPLOYEE_DASHBOARD");
                        } else {
                            throw new IllegalArgumentException("Employee ID not found in database.");
                        }
                    }
                } catch (IllegalArgumentException ex) {
                    JOptionPane.showMessageDialog(panel, ex.getMessage(), "Input Error", JOptionPane.WARNING_MESSAGE);
                } catch (SecurityException ex) {
                    JOptionPane.showMessageDialog(panel, ex.getMessage(), "Access Denied", JOptionPane.ERROR_MESSAGE);
                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(panel, "System error: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        return panel;
    }

    private JPanel createAdminDashboard() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBackground(BG_DARK);

        JPanel navBar = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        navBar.setBackground(PANEL_DARK);
        
        JLabel roleLabel = new JLabel("Role: Administrator | ");
        roleLabel.setForeground(TEXT_LIGHT);
        navBar.add(roleLabel);
        
        JButton logoutBtn = new JButton("Logout");
        logoutBtn.setBackground(BTN_ACCENT);
        logoutBtn.setForeground(TEXT_LIGHT);
        navBar.add(logoutBtn);
        
        panel.add(navBar, BorderLayout.NORTH);

        JPanel buttonPanel = new JPanel(new GridLayout(8, 1, 15, 15));
        
        buttonPanel.setBorder(BorderFactory.createEmptyBorder(50, 300, 50, 300));
        buttonPanel.setBackground(BG_DARK);

        Font btnFont = new Font("Arial", Font.PLAIN, 18);
        
        JButton viewAllEmpBtn = new JButton("View All Employee Profiles");
        viewAllEmpBtn.setFont(btnFont);
        viewAllEmpBtn.setBackground(PANEL_DARK);
        viewAllEmpBtn.setForeground(TEXT_LIGHT);
        
        JButton processOneBtn = new JButton("Process Single Employee Payroll");
        processOneBtn.setFont(btnFont);
        processOneBtn.setBackground(PANEL_DARK);
        processOneBtn.setForeground(TEXT_LIGHT);
        
        JButton processAllBtn = new JButton("Process Company Payroll (Bulk)");
        processAllBtn.setFont(btnFont);
        processAllBtn.setBackground(PANEL_DARK);
        processAllBtn.setForeground(TEXT_LIGHT);

        JButton addEmpBtn = new JButton("Add New Employee");
        addEmpBtn.setFont(btnFont);
        addEmpBtn.setBackground(PANEL_DARK);
        addEmpBtn.setForeground(new Color(100, 200, 100)); 
        
        JButton editEmpBtn = new JButton("Update Employee Status");
        editEmpBtn.setFont(btnFont);
        editEmpBtn.setBackground(PANEL_DARK);
        editEmpBtn.setForeground(new Color(255, 200, 100));
        
        JButton deleteEmpBtn = new JButton("Delete Employee Record");
        deleteEmpBtn.setFont(btnFont);
        deleteEmpBtn.setBackground(PANEL_DARK);
        deleteEmpBtn.setForeground(new Color(255, 100, 100));
        
        JButton editPayrollBtn = new JButton("Edit Payroll Records [WIP]"); 
        editPayrollBtn.setFont(btnFont);
        editPayrollBtn.setBackground(PANEL_DARK);
        editPayrollBtn.setForeground(TEXT_LIGHT);
        
        JButton databaseCheckBtn = new JButton("Check Database Connection");
        databaseCheckBtn.setFont(btnFont);
        databaseCheckBtn.setBackground(PANEL_DARK);
        databaseCheckBtn.setForeground(TEXT_LIGHT);

        buttonPanel.add(viewAllEmpBtn);
        buttonPanel.add(processOneBtn);
        buttonPanel.add(processAllBtn);
        buttonPanel.add(addEmpBtn);    
        buttonPanel.add(editEmpBtn);   
        buttonPanel.add(deleteEmpBtn); 
        buttonPanel.add(editPayrollBtn);
        buttonPanel.add(databaseCheckBtn);

        panel.add(buttonPanel, BorderLayout.CENTER);
        
        logoutBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                int confirm = JOptionPane.showConfirmDialog(panel, "Are you sure you want to log out?", "Logout", JOptionPane.YES_NO_OPTION);
                if (confirm == JOptionPane.YES_OPTION) {
                    currentLoggedInEmployeeId = "";
                    cardLayout.show(mainContainer, "LOGIN");
                }
            }
        });

        addEmpBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String newId = JOptionPane.showInputDialog(panel, "Enter New Employee ID:");
                if (newId != null && !newId.trim().isEmpty()) {
                    if (checkEmployeeExists(newId)) {
                        JOptionPane.showMessageDialog(panel, "Employee ID already exists!", "Error", JOptionPane.ERROR_MESSAGE);
                        return;
                    }
                    String fName = JOptionPane.showInputDialog(panel, "Enter First Name:");
                    String lName = JOptionPane.showInputDialog(panel, "Enter Last Name:");
                    
                    if (fName != null && lName != null) {
                        boolean success = saveNewEmployee(newId, fName, lName);
                        if(success) {
                            JOptionPane.showMessageDialog(panel, "Employee Successfully Added to Excel Database!");
                        } else {
                            JOptionPane.showMessageDialog(panel, "Error writing to Database.", "Error", JOptionPane.ERROR_MESSAGE);
                        }
                    }
                }
            }
        });

        editEmpBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String empId = JOptionPane.showInputDialog(panel, "Enter Employee ID to Update:");
                if (empId != null && !empId.trim().isEmpty()) {
                    if (checkEmployeeExists(empId)) {
                        String newStatus = JOptionPane.showInputDialog(panel, "Enter New Status (e.g. Regular, Probationary):");
                        if (newStatus != null) {
                            boolean success = updateEmployeeStatus(empId, newStatus);
                            if (success) {
                                JOptionPane.showMessageDialog(panel, "Employee Status Updated Successfully!");
                            } else {
                                JOptionPane.showMessageDialog(panel, "Error updating Database.", "Error", JOptionPane.ERROR_MESSAGE);
                            }
                        }
                    } else {
                        JOptionPane.showMessageDialog(panel, "Employee ID not found.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
            }
        });

        deleteEmpBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String empId = JOptionPane.showInputDialog(panel, "Enter Employee ID to Delete:");
                if (empId != null && !empId.trim().isEmpty()) {
                    if (checkEmployeeExists(empId)) {
                        int confirm = JOptionPane.showConfirmDialog(panel, "WARNING: Are you sure you want to delete employee " + empId + "?", "Delete Record", JOptionPane.YES_NO_OPTION);
                        if (confirm == JOptionPane.YES_OPTION) {
                            boolean success = removeEmployeeRecord(empId);
                            if (success) {
                                JOptionPane.showMessageDialog(panel, "Employee Record Deleted Permanently.");
                            } else {
                                JOptionPane.showMessageDialog(panel, "Error removing record from Database.", "Error", JOptionPane.ERROR_MESSAGE);
                            }
                        }
                    } else {
                        JOptionPane.showMessageDialog(panel, "Employee ID not found.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
            }
        });

        viewAllEmpBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String report = getAllEmployeeProfilesString();
                showReportWindow("All Employee Roster", HTML_HEADER + report + HTML_FOOTER);
            }
        });

        processOneBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String empId = JOptionPane.showInputDialog(panel, "Enter Employee ID to Process:");
                if (empId != null && !empId.trim().isEmpty()) {
                    if (checkEmployeeExists(empId)) {
                        handlePayrollFilterRequest(panel, empId);
                    } else {
                        JOptionPane.showMessageDialog(panel, "Employee ID not found.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                }
            }
        });

        processAllBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                handleBulkPayrollFilterRequest(panel);
            }
        });

        editPayrollBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JOptionPane.showMessageDialog(panel, "This feature is currently under development.", "Work In Progress", JOptionPane.INFORMATION_MESSAGE);
            }
        });

        databaseCheckBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    File file = new File(EXCEL_FILE_PATH);
                    if (!file.exists()) {
                        JOptionPane.showMessageDialog(panel, "Excel file not found!", "Error", JOptionPane.ERROR_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(panel, "Database Connected Successfully!", "Status", JOptionPane.INFORMATION_MESSAGE);
                    }
                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(panel, "Database Error: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        return panel;
    }

    private JPanel createEmployeeDashboard() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBackground(BG_DARK);

        JPanel navBar = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        navBar.setBackground(PANEL_DARK); 
        
        JLabel roleLabel = new JLabel("Role: Employee Self-Service | ");
        roleLabel.setForeground(TEXT_LIGHT);
        navBar.add(roleLabel);
        
        JButton logoutBtn = new JButton("Logout");
        logoutBtn.setBackground(BTN_ACCENT);
        logoutBtn.setForeground(TEXT_LIGHT);
        navBar.add(logoutBtn);
        
        panel.add(navBar, BorderLayout.NORTH);

        JPanel buttonPanel = new JPanel(new GridLayout(3, 1, 20, 20));
        
        buttonPanel.setBorder(BorderFactory.createEmptyBorder(150, 300, 150, 300));
        buttonPanel.setBackground(BG_DARK);

        Font btnFont = new Font("Arial", Font.PLAIN, 18);
        
        JButton viewProfileBtn = new JButton("View My Profile & Government IDs");
        viewProfileBtn.setFont(btnFont);
        viewProfileBtn.setBackground(PANEL_DARK);
        viewProfileBtn.setForeground(TEXT_LIGHT);
        
        JButton viewPayrollBtn = new JButton("View My Payslips");
        viewPayrollBtn.setFont(btnFont);
        viewPayrollBtn.setBackground(PANEL_DARK);
        viewPayrollBtn.setForeground(TEXT_LIGHT);

        JButton leaveRequestBtn = new JButton("Apply for Leave Request");
        leaveRequestBtn.setFont(btnFont);
        leaveRequestBtn.setBackground(PANEL_DARK);
        leaveRequestBtn.setForeground(new Color(150, 200, 255)); 

        buttonPanel.add(viewProfileBtn);
        buttonPanel.add(viewPayrollBtn);
        buttonPanel.add(leaveRequestBtn); 

        panel.add(buttonPanel, BorderLayout.CENTER);

        logoutBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                int confirm = JOptionPane.showConfirmDialog(panel, "Are you sure you want to log out?", "Logout", JOptionPane.YES_NO_OPTION);
                if (confirm == JOptionPane.YES_OPTION) {
                    currentLoggedInEmployeeId = ""; 
                    cardLayout.show(mainContainer, "LOGIN");
                }
            }
        });

        viewProfileBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String profileData = getEmployeeProfileString(currentLoggedInEmployeeId);
                showReportWindow("My Profile Details", HTML_HEADER + profileData + HTML_FOOTER);
            }
        });

        viewPayrollBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                handlePayrollFilterRequest(panel, currentLoggedInEmployeeId);
            }
        });

        leaveRequestBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String[] types = {"Sick Leave", "Vacation Leave", "Emergency Leave"};
                String selectedLeave = (String) JOptionPane.showInputDialog(panel, "Select Leave Type:", "Leave Request", 
                        JOptionPane.QUESTION_MESSAGE, null, types, types[0]);
                
                if (selectedLeave != null) {
                    String days = JOptionPane.showInputDialog(panel, "Enter number of days:");
                    if (days != null && !days.trim().isEmpty()) {
                        JOptionPane.showMessageDialog(panel, "Your " + selectedLeave + " request for " + days + " days has been submitted to Admin.", "Request Sent", JOptionPane.INFORMATION_MESSAGE);
                    }
                }
            }
        });

        return panel;
    }

    private static boolean saveNewEmployee(String id, String fName, String lName) {
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

    private static boolean updateEmployeeStatus(String id, String newStatus) {
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

    private static boolean removeEmployeeRecord(String id) {
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

    private void showReportWindow(String title, String htmlContent) {
        JEditorPane editorPane = new JEditorPane("text/html", htmlContent);
        editorPane.setEditable(false);
        editorPane.setBackground(BG_DARK);
        editorPane.setCaretPosition(0); 

        JScrollPane scrollPane = new JScrollPane(editorPane);
        scrollPane.setPreferredSize(new Dimension(1000, 500)); 
        scrollPane.setBorder(BorderFactory.createLineBorder(BTN_ACCENT));

        JOptionPane.showMessageDialog(this, scrollPane, title, JOptionPane.PLAIN_MESSAGE);
    }

    private void handlePayrollFilterRequest(JPanel parentPanel, String empId) {
        String[] options = {"All Time", "Latest Month", "Previous Month", "Custom Date"};
        int choice = JOptionPane.showOptionDialog(parentPanel, "Select the period you want to view:", 
                     "Filter Payslips", JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, 
                     null, options, options[0]);
                     
        String filter = "ALL";
        
        if (choice == 1) {
            filter = "LATEST";
        } else if (choice == 2) {
            filter = "PREVIOUS";
        } else if (choice == 3) {
            filter = JOptionPane.showInputDialog(parentPanel, "Enter Month and Year (Format: MM/YYYY)");
            if (filter == null || filter.trim().isEmpty()) {
                return; 
            }
        } else if (choice != 0) {
            return; 
        }
        
        String report = processPayrollLoop(empId, filter);
        
        String windowTitle = "Payroll for ID: " + empId + " (";
        if (filter.equals("ALL")) {
            windowTitle = windowTitle + "All Records)";
        } else {
            windowTitle = windowTitle + filter + ")";
        }
        
        showReportWindow(windowTitle, HTML_HEADER + report + HTML_FOOTER);
    }

    private void handleBulkPayrollFilterRequest(JPanel parentPanel) {
        String[] options = {"All Time", "Latest Month", "Previous Month", "Custom Date"};
        int choice = JOptionPane.showOptionDialog(parentPanel, "Select the period for the bulk report:", 
                     "Filter Bulk Payslips", JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, 
                     null, options, options[0]);
                     
        String filter = "ALL";
        
        if (choice == 1) {
            filter = "LATEST";
        } else if (choice == 2) {
            filter = "PREVIOUS";
        } else if (choice == 3) {
            filter = JOptionPane.showInputDialog(parentPanel, "Enter Month and Year (Format: MM/YYYY)");
            if (filter == null || filter.trim().isEmpty()) {
                return;
            }
        } else if (choice != 0) {
            return; 
        }
        
        String report = processAllEmployees(filter);
        
        String windowTitle = "Company Bulk Payroll Report (";
        if (filter.equals("ALL")) {
            windowTitle = windowTitle + "All Records)";
        } else {
            windowTitle = windowTitle + filter + ")";
        }
        
        showReportWindow(windowTitle, HTML_HEADER + report + HTML_FOOTER);
    }

    private String processPayrollLoop(String id, String periodFilter) {
        StringBuilder sb = new StringBuilder();
        
        try (FileInputStream fis = new FileInputStream(new File(EXCEL_FILE_PATH));
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            Sheet empSheet = workbook.getSheet("Employee Details");
            Sheet attSheet = workbook.getSheet("Attendance Record");
            
            Row row = null;
            for (Row r : empSheet) {
                if (getCellValueAsString(r.getCell(0)).equals(id)) {
                    row = r;
                    break;
                }
            }

            if (row == null) {
                return "<p style='color:red;'>Employee ID not found.</p>";
            }

            int hourlyColIndex = 18; 
            Row headerRow = empSheet.getRow(0);
            if (headerRow != null) {
                for (Cell cell : headerRow) {
                    if (getCellValueAsString(cell).toLowerCase().contains("hourly")) {
                        hourlyColIndex = cell.getColumnIndex();
                        break;
                    }
                }
            }

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            List<String> activePeriods = getUniquePeriods(attSheet, id);
            
            if (periodFilter.equals("LATEST")) {
                if (activePeriods.size() > 0) {
                    periodFilter = activePeriods.get(activePeriods.size() - 1); 
                }
            } else if (periodFilter.equals("PREVIOUS")) {
                if (activePeriods.size() > 1) {
                    periodFilter = activePeriods.get(activePeriods.size() - 2); 
                } else if (activePeriods.size() > 0) {
                    periodFilter = activePeriods.get(0);
                }
            }

            if (!periodFilter.equals("ALL")) {
                List<String> filteredList = new ArrayList<>();
                if (activePeriods.contains(periodFilter)) {
                    filteredList.add(periodFilter);
                }
                activePeriods = filteredList;
            }
            
            String name = getCellValueAsString(row.getCell(2)) + " " + getCellValueAsString(row.getCell(1));
            
            sb.append("<h2 style='color: #4DA6FF; border-bottom: 1px solid #4DA6FF; padding-bottom: 5px;'>Payroll Summary for: [").append(id).append("] ").append(name).append("</h2>");
            
            if (activePeriods.size() == 0) {
                sb.append("<p>No attendance records found for this employee in the selected period.</p>");
                return sb.toString();
            }

            sb.append("<table border='1' cellpadding='10' cellspacing='0' style='border-collapse: collapse; width: 100%; border-color: #555;'>");
            
            sb.append("<tr style='background-color: #2C2C2C; color: #4DA6FF;'>");
            
            sb.append("<th>Period</th><th>1st Cutoff</th><th>1st NET</th><th>2nd Cutoff</th><th>Deductions</th><th>Tax</th><th>2nd NET</th><th>TOTAL NET PAY</th>");
            sb.append("</tr>");

            for (String period : activePeriods) {
                double h1 = calculateHours(attSheet, id, period, 1, 15);
                double h2 = calculateHours(attSheet, id, period, 16, 31);
                double totalHours = h1 + h2;

                if (totalHours > 0) {
                    double hourlyRate = getNumericSafe(row.getCell(hourlyColIndex), evaluator); 
                    if (hourlyRate > 1000) {
                        hourlyRate = hourlyRate / 160;
                    }

                    double basicSalary = getNumericSafe(row.getCell(13), evaluator); 
                    if (basicSalary <= 0) {
                        basicSalary = hourlyRate * 160;
                    }

                    double gross1 = hourlyRate * h1;
                    double gross2 = hourlyRate * h2;
                    double totalGross = gross1 + gross2;
                    
                    double sss = calculateSSS(basicSalary);
                    double philHealth = basicSalary * 0.025; 
                    if (philHealth > 2500.00) {
                        philHealth = 2500.00; 
                    }
                    double pagIbig = 200.00; 
                    
                    double totalGovtDeductions = sss + philHealth + pagIbig;
                    double taxableIncome = totalGross - totalGovtDeductions;
                    double tax = calculateTax(taxableIncome);
                    
                    double net1 = gross1; 
                    double net2 = gross2 - (totalGovtDeductions + tax);
                    
                    if (net2 < 0) {
                        net1 = net1 + net2; 
                        net2 = 0;           
                    }
                    
                    double totalNetPay = net1 + net2;

                    sb.append("<tr style='text-align: center; background-color: #1E1E1E;'>");
                    sb.append("<td>").append(period).append("</td>");
                    sb.append("<td>").append(String.format("%,.2f", gross1)).append("</td>");
                    sb.append("<td>").append(String.format("%,.2f", net1)).append("</td>");
                    sb.append("<td>").append(String.format("%,.2f", gross2)).append("</td>");
                    
                    sb.append("<td style='color: #FF6666;'>-").append(String.format("%,.2f", totalGovtDeductions)).append("</td>"); 
                    sb.append("<td style='color: #FF6666;'>-").append(String.format("%,.2f", tax)).append("</td>"); 
                    
                    sb.append("<td>").append(String.format("%,.2f", net2)).append("</td>");
                    
                    sb.append("<td style='color: #4CAF50; font-weight: bold;'>").append(String.format("%,.2f", totalNetPay)).append("</td>"); 
                    
                    sb.append("</tr>");
                }
            }
            sb.append("</table><br>");

        } catch (Exception e) {
            sb.append("<p style='color:red;'>Error processing payroll records: ").append(e.getMessage()).append("</p>");
        }
        return sb.toString();
    }

    private String processAllEmployees(String periodFilter) {
        StringBuilder sb = new StringBuilder();
        try (FileInputStream fis = new FileInputStream(new File(EXCEL_FILE_PATH));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheet("Employee Details");
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue; 
                }
                String id = getCellValueAsString(row.getCell(0));
                
                String report = processPayrollLoop(id, periodFilter);
                if (!report.contains("No attendance records found")) {
                    sb.append(report);
                }
            }
            if (sb.toString().trim().isEmpty()) {
                sb.append("<p>No active payroll records found for the selected timeframe.</p>");
            }
        } catch (Exception e) {
            sb.append("<p style='color:red;'>Error reading records from the database.</p>");
        }
        return sb.toString();
    }

    private String getAllEmployeeProfilesString() {
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

    private String getEmployeeProfileString(String searchId) {
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

    private static boolean checkEmployeeExists(String id) {
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

    private static List<String> getUniquePeriods(Sheet sheet, String id) {
        Set<String> periods = new LinkedHashSet<>();
        for (Row row : sheet) {
            if (getCellValueAsString(row.getCell(0)).equals(id)) {
                String date = getCellValueAsString(row.getCell(3));
                try {
                    String[] parts = date.split("/");
                    if (parts.length >= 3) {
                        String mm = parts[0];
                        if (mm.length() == 1) {
                            mm = "0" + mm;
                        }
                        String yyyy = parts[2].split(" ")[0]; 
                        periods.add(mm + "/" + yyyy);
                    }
                } catch (Exception e) {}
            }
        }
        return new ArrayList<>(periods);
    }

    private static double calculateHours(Sheet sheet, String id, String targetPeriod, int startDay, int endDay) {
        double total = 0;
        LocalTime shiftStart = LocalTime.of(8, 0);
        LocalTime shiftEnd = LocalTime.of(17, 0);
        LocalTime grace = LocalTime.of(8, 10);

        for (Row row : sheet) {
            if (getCellValueAsString(row.getCell(0)).equals(id)) {
                String date = getCellValueAsString(row.getCell(3));
                try {
                    String[] parts = date.split("/");
                    if (parts.length >= 3) {
                        String mm = parts[0];
                        if (mm.length() == 1) {
                            mm = "0" + mm;
                        }
                        String yyyy = parts[2].split(" ")[0];
                        
                        if ((mm + "/" + yyyy).equals(targetPeriod)) {
                            int day = Integer.parseInt(parts[1]);
                            if (day >= startDay && day <= endDay) {
                                LocalTime timeIn = parseTime(getCellValueAsString(row.getCell(4)));
                                LocalTime timeOut = parseTime(getCellValueAsString(row.getCell(5)));
                                
                                if (timeIn.isBefore(shiftStart) || (timeIn.isAfter(shiftStart) && timeIn.isBefore(grace))) {
                                    timeIn = shiftStart;
                                }
                                if (timeOut.isAfter(shiftEnd)) {
                                    timeOut = shiftEnd;
                                }
                                
                                if (timeOut.isAfter(timeIn)) {
                                    double duration = Duration.between(timeIn, timeOut).toMinutes() / 60.0;
                                    if (duration > 5) {
                                        duration = duration - 1.0; 
                                    }
                                    total = total + duration;
                                }
                            }
                        }
                    }
                } catch (Exception ex) {}
            }
        }
        return total;
    }

    private static double getNumericSafe(Cell cell, FormulaEvaluator evaluator) {
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

    private static double calculateSSS(double salary) {
        if (salary <= 3250) return 135.00;
        if (salary >= 24750) return 1125.00;
        int steps = (int)((salary - 3250 - 0.01) / 500) + 1;
        return 135.00 + (steps * 22.50);
    }

    private static double calculateTax(double taxableIncome) {
        if (taxableIncome <= 20833) return 0;
        if (taxableIncome <= 33333) return (taxableIncome - 20833) * 0.15;
        if (taxableIncome <= 66667) return 1875 + (taxableIncome - 33333) * 0.20;
        if (taxableIncome <= 166667) return 8541.67 + (taxableIncome - 66667) * 0.25;
        return 33541.67 + (taxableIncome - 166667) * 0.30;
    }

    private static LocalTime parseTime(String t) {
        try {
            t = t.trim().toUpperCase();
            if (t.contains("AM") || t.contains("PM")) {
                return LocalTime.parse(t, DateTimeFormatter.ofPattern("h:mm a"));
            } else {
                return LocalTime.parse(t);
            }
        } catch (Exception e) { 
            return LocalTime.of(8, 0); 
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }
        return FORMATTER.formatCellValue(cell).trim();
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new MotorPH().setVisible(true);
            }
        });
    }
}