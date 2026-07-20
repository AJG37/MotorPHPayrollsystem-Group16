/*
 * Project: MotorPH Payroll System
 * Description: An application for managing employee data and calculating payroll.
 * It features a login portal with access levels for Administrators and Employees.
 * All data is read from and written to a local Excel database.
 */

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import repository.ExcelRepository;
import service.PayrollService;
import javax.swing.*;
import java.awt.*;
import java.awt.Font;  
import java.awt.Color; 
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.time.Duration;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

// Main Swing application window for the MotorPH payroll system.
public class MotorPH extends JFrame {
    
    // Card-based navigation between login/admin/employee views.
    private CardLayout cardLayout;
    private JPanel mainContainer;
    // Tracks the current user session within the UI.
    private String currentLoggedInEmployeeId = "";

    // Static credentials for admin access.
    private static final String ADMIN_ID = "999";
    private static final String ADMIN_PASS = "admin123";
    private static final String EMPLOYEE_ID_PATTERN = "\\d{5}";
    private static final String EMPLOYEE_NAME_PATTERN = "\\p{L}+(?:[ '\\-]\\p{L}+)*";
    private static final String EMPLOYEE_POSITION_PATTERN = "(?=.*\\p{L})[\\p{L} '&\\-]+";
    private static final String PAYROLL_PERIOD_PATTERN = "(0[1-9]|1[0-2])/\\d{4}";
    private static final String[] VALID_EMPLOYEE_STATUSES = {"Regular", "Probationary"};
    // Path to the local Excel file used as the database.
    private static final String EXCEL_FILE_PATH = "MotorPH_EmployeeData.xlsx";
    // UI theme colors.
    private static final Color BG_DARK = new Color(15, 23, 42);
    private static final Color PANEL_DARK = new Color(51, 65, 85);
    private static final Color FIELD_DARK = new Color(17, 24, 39);
    private static final Color TEXT_LIGHT = new Color(248, 250, 252);
    private static final Color BTN_ACCENT = new Color(37, 99, 235);
    private static final Color BORDER_LIGHT = new Color(100, 116, 139);
    private static final Color SUCCESS_TEXT = new Color(134, 239, 172);
    private static final Color WARNING_TEXT = new Color(253, 230, 138);
    private static final Color DANGER_TEXT = new Color(252, 165, 165);
    private static final Color INFO_TEXT = new Color(147, 197, 253);
    
    // HTML wrapper used for report dialogs.
    private static final String HTML_HEADER = "<html><body style='font-family: Arial, sans-serif; background-color: #0F172A; color: #F8FAFC; margin: 10px;'>";
    private static final String HTML_FOOTER = "</body></html>";

    // Small interface used to prepare report content in the background.
    private interface ReportLoader {
        String loadReport();
    }

    // Application constructor that initializes the main window and cards.
    public MotorPH() {
        setTitle("MotorPH Payroll System");
        
        setExtendedState(JFrame.MAXIMIZED_BOTH); 
        
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setResizable(true); 
        
        // Ensure JOptionPane dialogs match the dark theme.
        UIManager.put("OptionPane.background", PANEL_DARK);
        UIManager.put("Panel.background", PANEL_DARK);
        UIManager.put("OptionPane.messageForeground", TEXT_LIGHT);
        UIManager.put("Button.background", BTN_ACCENT);
        UIManager.put("Button.foreground", TEXT_LIGHT);
        UIManager.put("TextField.background", FIELD_DARK);
        UIManager.put("TextField.foreground", TEXT_LIGHT);
        UIManager.put("PasswordField.background", FIELD_DARK);
        UIManager.put("PasswordField.foreground", TEXT_LIGHT);
        UIManager.put("ComboBox.background", FIELD_DARK);
        UIManager.put("ComboBox.foreground", TEXT_LIGHT);
        UIManager.put("ProgressBar.foreground", BTN_ACCENT);

        cardLayout = new CardLayout();
        mainContainer = new JPanel(cardLayout);

        // Register the main screens.
        mainContainer.add(createLoginScreen(), "LOGIN");
        mainContainer.add(createAdminDashboard(), "ADMIN_DASHBOARD");
        mainContainer.add(createEmployeeDashboard(), "EMPLOYEE_DASHBOARD");

        add(mainContainer);
        cardLayout.show(mainContainer, "LOGIN");
    }

    // Stops a file-dependent action and explains the expected Excel format.
    private boolean ensureWorkbookReady(Component parent) {
        String validationError = ExcelRepository.getWorkbookValidationError();
        if (validationError != null) {
            JOptionPane.showMessageDialog(parent, validationError,
                    "Excel Data File Error", JOptionPane.ERROR_MESSAGE);
            return false;
        }
        return true;
    }

    // Employee IDs in the current MotorPH workbook use exactly five digits.
    private static boolean isValidEmployeeId(String employeeId) {
        return employeeId != null && employeeId.matches(EMPLOYEE_ID_PATTERN);
    }

    // Names may contain letters, single spaces, hyphens, and apostrophes.
    private static boolean isValidEmployeeName(String name) {
        return name != null && name.matches(EMPLOYEE_NAME_PATTERN);
    }

    // Positions may use letters, spaces, hyphens, apostrophes, and ampersands.
    private static boolean isValidEmployeePosition(String position) {
        return position != null && position.matches(EMPLOYEE_POSITION_PATTERN);
    }

    // Returns the standard status spelling, or null when the value is not allowed.
    private static String getValidEmployeeStatus(String status) {
        if (status == null) {
            return null;
        }
        for (String validStatus : VALID_EMPLOYEE_STATUSES) {
            if (validStatus.equalsIgnoreCase(status)) {
                return validStatus;
            }
        }
        return null;
    }

    // Prompts for a month/year and explains the required format when invalid.
    private static String promptForPayrollPeriod(Component parent) {
        String period = JOptionPane.showInputDialog(parent,
                "Enter Month and Year (MM/YYYY):");
        if (period == null) {
            return null;
        }

        period = period.trim();
        if (!period.matches(PAYROLL_PERIOD_PATTERN)) {
            JOptionPane.showMessageDialog(parent,
                    "Enter a valid month and year in MM/YYYY format (example: 01/2024).",
                    "Invalid Payroll Period", JOptionPane.WARNING_MESSAGE);
            return null;
        }
        return period;
    }

    // Applies consistent contrast and spacing to dashboard menu buttons.
    private static void styleMenuButton(JButton button, Font font, Color textColor) {
        button.setFont(font);
        button.setBackground(PANEL_DARK);
        button.setForeground(textColor);
        button.setFocusPainted(false);
        button.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        button.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(BORDER_LIGHT),
                BorderFactory.createEmptyBorder(8, 16, 8, 16)));
    }

    // Applies the brighter blue style used for primary actions.
    private static void styleAccentButton(JButton button) {
        button.setBackground(BTN_ACCENT);
        button.setForeground(TEXT_LIGHT);
        button.setFocusPainted(false);
        button.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        button.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(INFO_TEXT),
                BorderFactory.createEmptyBorder(5, 12, 5, 12)));
    }

    // Builds the login screen UI and validation logic.
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
        idField.setBackground(FIELD_DARK);
        idField.setForeground(TEXT_LIGHT);
        idField.setCaretColor(Color.WHITE);
        idField.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(BORDER_LIGHT),
                BorderFactory.createEmptyBorder(5, 7, 5, 7)));
        gbc.gridx = 1;
        panel.add(idField, gbc);

        JLabel passwordLabel = new JLabel("Admin Password:");
        passwordLabel.setForeground(TEXT_LIGHT);
        gbc.gridx = 0; gbc.gridy = 2;
        panel.add(passwordLabel, gbc);

        JPasswordField passwordField = new JPasswordField(15);
        passwordField.setFont(new Font("Arial", Font.PLAIN, 16));
        passwordField.setBackground(FIELD_DARK);
        passwordField.setForeground(TEXT_LIGHT);
        passwordField.setCaretColor(Color.WHITE);
        passwordField.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(BORDER_LIGHT),
                BorderFactory.createEmptyBorder(5, 7, 5, 7)));
        passwordField.setToolTipText("Required only for administrator login");
        gbc.gridx = 1;
        panel.add(passwordField, gbc);

        JButton loginBtn = new JButton("Login");
        loginBtn.setFont(new Font("Arial", Font.BOLD, 14));
        styleAccentButton(loginBtn);
        gbc.gridx = 0; gbc.gridy = 3; gbc.gridwidth = 2;
        panel.add(loginBtn, gbc);

        // Attempt login; route to admin or employee dashboard.
        loginBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                try {
                    String userId = idField.getText().trim();

                    // Basic input validation.
                    if (userId.isEmpty()) {
                        throw new IllegalArgumentException("Employee ID is required.");
                    }

                    // Admin login uses the password field shown below the employee ID.
                    if (userId.equals(ADMIN_ID)) {
                        String pass = new String(passwordField.getPassword());
                        if (pass.equals(ADMIN_PASS)) {
                            // Successful admin login.
                            currentLoggedInEmployeeId = ADMIN_ID;
                            idField.setText("");
                            passwordField.setText("");
                            cardLayout.show(mainContainer, "ADMIN_DASHBOARD");
                        } else {
                            throw new SecurityException("The administrator password is incorrect.");
                        }
                    } else {
                        // Employee login validates ID against the database.
                        if (!ensureWorkbookReady(panel)) {
                            return;
                        }
                        if (ExcelRepository.checkEmployeeExists(userId)) {
                            currentLoggedInEmployeeId = userId;
                            idField.setText("");
                            passwordField.setText("");
                            cardLayout.show(mainContainer, "EMPLOYEE_DASHBOARD");
                        } else {
                            throw new IllegalArgumentException("Employee ID " + userId + " was not found.");
                        }
                    }
                } catch (IllegalArgumentException ex) {
                    JOptionPane.showMessageDialog(panel, ex.getMessage(), "Input Error", JOptionPane.WARNING_MESSAGE);
                } catch (SecurityException ex) {
                    JOptionPane.showMessageDialog(panel, ex.getMessage(), "Access Denied", JOptionPane.ERROR_MESSAGE);
                    passwordField.setText("");
                    passwordField.requestFocusInWindow();
                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(panel, "System error: " + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        // Enter moves admins to the password field; employee IDs log in directly.
        idField.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (idField.getText().trim().equals(ADMIN_ID)) {
                    passwordField.requestFocusInWindow();
                } else {
                    loginBtn.doClick();
                }
            }
        });

        // Pressing Enter in the password field submits the existing Login action.
        passwordField.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                loginBtn.doClick();
            }
        });

        return panel;
    }

    // Builds the administrator dashboard and actions.
    private JPanel createAdminDashboard() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBackground(BG_DARK);

        JPanel navBar = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        navBar.setBackground(PANEL_DARK);
        
        JLabel roleLabel = new JLabel("Role: Administrator | ");
        roleLabel.setForeground(TEXT_LIGHT);
        navBar.add(roleLabel);
        
        JButton logoutBtn = new JButton("Logout");
        styleAccentButton(logoutBtn);
        navBar.add(logoutBtn);
        
        panel.add(navBar, BorderLayout.NORTH);

        JPanel buttonPanel = new JPanel(new GridLayout(8, 1, 15, 15));
        
        buttonPanel.setBorder(BorderFactory.createEmptyBorder(50, 300, 50, 300));
        buttonPanel.setBackground(BG_DARK);

        Font btnFont = new Font("Arial", Font.PLAIN, 18);
        
        JButton viewAllEmpBtn = new JButton("View All Employee Profiles");
        styleMenuButton(viewAllEmpBtn, btnFont, TEXT_LIGHT);
        
        JButton processOneBtn = new JButton("Process Single Employee Payroll");
        styleMenuButton(processOneBtn, btnFont, TEXT_LIGHT);
        
        JButton processAllBtn = new JButton("Process Company Payroll (Bulk)");
        styleMenuButton(processAllBtn, btnFont, INFO_TEXT);

        JButton addEmpBtn = new JButton("Add New Employee");
        styleMenuButton(addEmpBtn, btnFont, SUCCESS_TEXT);
        
        JButton editEmpBtn = new JButton("Update Employee Record");
        styleMenuButton(editEmpBtn, btnFont, WARNING_TEXT);
        
        JButton deleteEmpBtn = new JButton("Delete Employee Record");
        styleMenuButton(deleteEmpBtn, btnFont, DANGER_TEXT);
        
        JButton databaseCheckBtn = new JButton("Check Database Connection");
        styleMenuButton(databaseCheckBtn, btnFont, TEXT_LIGHT);

        buttonPanel.add(viewAllEmpBtn);
        buttonPanel.add(processOneBtn);
        buttonPanel.add(processAllBtn);
        buttonPanel.add(addEmpBtn);    
        buttonPanel.add(editEmpBtn);   
        buttonPanel.add(deleteEmpBtn); 
        buttonPanel.add(databaseCheckBtn);

        panel.add(buttonPanel, BorderLayout.CENTER);
        
        // Clear session and return to login.
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

        // Adds a new employee row to the Excel database.
        addEmpBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (!ensureWorkbookReady(panel)) {
                    return;
                }
                String newId = JOptionPane.showInputDialog(panel, "Enter New Employee ID:");
                if (newId == null) {
                    return;
                }

                newId = newId.trim();
                if (newId.isEmpty()) {
                    JOptionPane.showMessageDialog(panel, "Employee ID is required.",
                            "Invalid Employee ID", JOptionPane.WARNING_MESSAGE);
                    return;
                }
                if (!isValidEmployeeId(newId)) {
                    JOptionPane.showMessageDialog(panel,
                            "Employee ID must contain exactly 5 digits (example: 10035).",
                            "Invalid Employee ID", JOptionPane.WARNING_MESSAGE);
                    return;
                }

                // Prevent duplicate IDs before writing.
                if (ExcelRepository.checkEmployeeExists(newId)) {
                    JOptionPane.showMessageDialog(panel,
                            "Employee ID " + newId + " already exists.",
                            "Duplicate Employee", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                String fName = JOptionPane.showInputDialog(panel, "Enter First Name:");
                if (fName == null) {
                    return;
                }
                fName = fName.trim();
                if (fName.isEmpty()) {
                    JOptionPane.showMessageDialog(panel, "First Name is required.",
                            "Invalid First Name", JOptionPane.WARNING_MESSAGE);
                    return;
                }
                if (!isValidEmployeeName(fName)) {
                    JOptionPane.showMessageDialog(panel,
                            "First Name may contain only letters, spaces, hyphens, and apostrophes.",
                            "Invalid First Name", JOptionPane.WARNING_MESSAGE);
                    return;
                }

                String lName = JOptionPane.showInputDialog(panel, "Enter Last Name:");
                if (lName == null) {
                    return;
                }
                lName = lName.trim();
                if (lName.isEmpty()) {
                    JOptionPane.showMessageDialog(panel, "Last Name is required.",
                            "Invalid Last Name", JOptionPane.WARNING_MESSAGE);
                    return;
                }
                if (!isValidEmployeeName(lName)) {
                    JOptionPane.showMessageDialog(panel,
                            "Last Name may contain only letters, spaces, hyphens, and apostrophes.",
                            "Invalid Last Name", JOptionPane.WARNING_MESSAGE);
                    return;
                }

                boolean success = ExcelRepository.saveNewEmployee(newId, fName, lName);
                if(success) {
                    JOptionPane.showMessageDialog(panel,
                            "Employee " + newId + " was added successfully.",
                            "Employee Added", JOptionPane.INFORMATION_MESSAGE);
                } else {
                    JOptionPane.showMessageDialog(panel,
                            "Unable to add the employee. Close the Excel file if it is open, then try again.",
                            "Add Failed", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        // Updates safe employee text fields without changing payroll values or formulas.
        editEmpBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (!ensureWorkbookReady(panel)) {
                    return;
                }
                String empId = JOptionPane.showInputDialog(panel, "Enter Employee ID to Update:");
                if (empId == null) {
                    return;
                }

                empId = empId.trim();
                if (empId.isEmpty()) {
                    JOptionPane.showMessageDialog(panel, "Employee ID is required.",
                            "Invalid Employee ID", JOptionPane.WARNING_MESSAGE);
                    return;
                }
                if (!isValidEmployeeId(empId)) {
                    JOptionPane.showMessageDialog(panel,
                            "Employee ID must contain exactly 5 digits.",
                            "Invalid Employee ID", JOptionPane.WARNING_MESSAGE);
                    return;
                }
                if (!ExcelRepository.checkEmployeeExists(empId)) {
                    JOptionPane.showMessageDialog(panel,
                            "Employee ID " + empId + " was not found.",
                            "Employee Not Found", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                String[] editableFields = {"First Name", "Last Name", "Status", "Position"};
                String selectedField = (String) JOptionPane.showInputDialog(panel,
                        "Select the employee field to update:",
                        "Update Employee Record",
                        JOptionPane.QUESTION_MESSAGE,
                        null,
                        editableFields,
                        editableFields[0]);
                if (selectedField == null) {
                    return;
                }

                String prompt;
                int columnIndex;
                switch (selectedField) {
                    case "First Name":
                        prompt = "Enter New First Name:";
                        columnIndex = 2;
                        break;
                    case "Last Name":
                        prompt = "Enter New Last Name:";
                        columnIndex = 1;
                        break;
                    case "Status":
                        prompt = "Enter New Status (Regular or Probationary):";
                        columnIndex = 10;
                        break;
                    case "Position":
                        prompt = "Enter New Position:";
                        columnIndex = 11;
                        break;
                    default:
                        return;
                }

                String newValue = JOptionPane.showInputDialog(panel, prompt);
                if (newValue == null) {
                    return;
                }

                newValue = newValue.trim();
                if (newValue.isEmpty()) {
                    JOptionPane.showMessageDialog(panel,
                            selectedField + " is required.",
                            "Invalid " + selectedField, JOptionPane.WARNING_MESSAGE);
                    return;
                }

                if ((selectedField.equals("First Name") || selectedField.equals("Last Name"))
                        && !isValidEmployeeName(newValue)) {
                    JOptionPane.showMessageDialog(panel,
                            selectedField + " may contain only letters, spaces, hyphens, and apostrophes.",
                            "Invalid " + selectedField, JOptionPane.WARNING_MESSAGE);
                    return;
                }

                if (selectedField.equals("Status")) {
                    String validStatus = getValidEmployeeStatus(newValue);
                    if (validStatus == null) {
                        JOptionPane.showMessageDialog(panel,
                                "Employee Status must be Regular or Probationary.",
                                "Invalid Status", JOptionPane.WARNING_MESSAGE);
                        return;
                    }
                    newValue = validStatus;
                }

                if (selectedField.equals("Position") && !isValidEmployeePosition(newValue)) {
                    JOptionPane.showMessageDialog(panel,
                            "Position may contain only letters, spaces, hyphens, apostrophes, and ampersands.",
                            "Invalid Position", JOptionPane.WARNING_MESSAGE);
                    return;
                }

                boolean success = ExcelRepository.updateEmployeeField(empId, columnIndex, newValue);
                if (success) {
                    JOptionPane.showMessageDialog(panel,
                            "Employee " + empId + " " + selectedField
                                    + " updated to " + newValue + ".",
                            "Update Successful", JOptionPane.INFORMATION_MESSAGE);
                } else {
                    JOptionPane.showMessageDialog(panel,
                            "Unable to update the employee record.",
                            "Update Failed", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        // Deletes an employee record after confirmation.
        deleteEmpBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (!ensureWorkbookReady(panel)) {
                    return;
                }
                String empId = JOptionPane.showInputDialog(panel, "Enter Employee ID to Delete:");
                if (empId == null) {
                    return;
                }

                empId = empId.trim();
                if (!isValidEmployeeId(empId)) {
                    JOptionPane.showMessageDialog(panel,
                            "Employee ID must contain exactly 5 digits.",
                            "Invalid Employee ID", JOptionPane.WARNING_MESSAGE);
                    return;
                }

                if (ExcelRepository.checkEmployeeExists(empId)) {
                        int confirm = JOptionPane.showConfirmDialog(panel,
                                "Delete employee " + empId + "?\nThis permanently removes the employee record.",
                                "Confirm Delete", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
                        if (confirm == JOptionPane.YES_OPTION) {
                            boolean success = ExcelRepository.removeEmployeeRecord(empId);
                            if (success) {
                                JOptionPane.showMessageDialog(panel,
                                        "Employee " + empId + " was deleted successfully.",
                                        "Employee Deleted", JOptionPane.INFORMATION_MESSAGE);
                            } else {
                                JOptionPane.showMessageDialog(panel,
                                        "Unable to delete the employee. Close the Excel file if it is open, then try again.",
                                        "Delete Failed", JOptionPane.ERROR_MESSAGE);
                            }
                        }
                } else {
                    JOptionPane.showMessageDialog(panel,
                            "Employee ID " + empId + " was not found.",
                            "Employee Not Found", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        // Generates and displays the full roster report.
        viewAllEmpBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (!ensureWorkbookReady(panel)) {
                    return;
                }
                showReportWithLoading("Loading employee roster...", "All Employee Roster",
                        new ReportLoader() {
                            @Override
                            public String loadReport() {
                                String report = ExcelRepository.getAllEmployeeProfilesString();
                                return HTML_HEADER + report + HTML_FOOTER;
                            }
                        });
            }
        });

        // Generates payroll report for a single employee.
        processOneBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (!ensureWorkbookReady(panel)) {
                    return;
                }
                String empId = JOptionPane.showInputDialog(panel, "Enter Employee ID to Process:");
                if (empId == null) {
                    return;
                }

                empId = empId.trim();
                if (!isValidEmployeeId(empId)) {
                    JOptionPane.showMessageDialog(panel,
                            "Employee ID must contain exactly 5 digits.",
                            "Invalid Employee ID", JOptionPane.WARNING_MESSAGE);
                    return;
                }

                if (ExcelRepository.checkEmployeeExists(empId)) {
                    handlePayrollFilterRequest(panel, empId);
                } else {
                    JOptionPane.showMessageDialog(panel,
                            "Employee ID " + empId + " was not found.",
                            "Employee Not Found", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        // Generates payroll report for all employees.
        processAllBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (!ensureWorkbookReady(panel)) {
                    return;
                }
                handleBulkPayrollFilterRequest(panel);
            }
        });

        // Verifies the Excel file exists and is reachable.
        databaseCheckBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (ensureWorkbookReady(panel)) {
                    JOptionPane.showMessageDialog(panel,
                            "Excel/XLSX employee data file is ready.",
                            "Database Status", JOptionPane.INFORMATION_MESSAGE);
                }
            }
        });

        return panel;
    }

    // Builds the employee self-service dashboard.
    private JPanel createEmployeeDashboard() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBackground(BG_DARK);

        JPanel navBar = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        navBar.setBackground(PANEL_DARK); 
        
        JLabel roleLabel = new JLabel("Role: Employee Self-Service | ");
        roleLabel.setForeground(TEXT_LIGHT);
        navBar.add(roleLabel);
        
        JButton logoutBtn = new JButton("Logout");
        styleAccentButton(logoutBtn);
        navBar.add(logoutBtn);
        
        panel.add(navBar, BorderLayout.NORTH);

        JPanel buttonPanel = new JPanel(new GridLayout(3, 1, 20, 20));
        
        buttonPanel.setBorder(BorderFactory.createEmptyBorder(150, 300, 150, 300));
        buttonPanel.setBackground(BG_DARK);

        Font btnFont = new Font("Arial", Font.PLAIN, 18);
        
        JButton viewProfileBtn = new JButton("View My Profile & Government IDs");
        styleMenuButton(viewProfileBtn, btnFont, TEXT_LIGHT);
        
        JButton viewPayrollBtn = new JButton("View My Payslips");
        styleMenuButton(viewPayrollBtn, btnFont, INFO_TEXT);

        JButton leaveRequestBtn = new JButton("Apply for Leave Request");
        styleMenuButton(leaveRequestBtn, btnFont, INFO_TEXT);

        buttonPanel.add(viewProfileBtn);
        buttonPanel.add(viewPayrollBtn);
        buttonPanel.add(leaveRequestBtn); 

        panel.add(buttonPanel, BorderLayout.CENTER);

        // Clear session and return to login.
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

        // Load and show profile and government IDs.
        viewProfileBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (!ensureWorkbookReady(panel)) {
                    return;
                }
                final String employeeId = currentLoggedInEmployeeId;
                showReportWithLoading("Loading employee profile...", "My Profile Details",
                        new ReportLoader() {
                            @Override
                            public String loadReport() {
                                String profileData = ExcelRepository.getEmployeeProfileString(employeeId);
                                return HTML_HEADER + profileData + HTML_FOOTER;
                            }
                        });
            }
        });

        // Load and show employee payroll records.
        viewPayrollBtn.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (!ensureWorkbookReady(panel)) {
                    return;
                }
                handlePayrollFilterRequest(panel, currentLoggedInEmployeeId);
            }
        });

        // Simple leave request capture (no persistence).
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

    // Displays HTML content in a scrollable dialog.
    private void showReportWindow(String title, String htmlContent) {
        JEditorPane editorPane = new JEditorPane("text/html", htmlContent);
        editorPane.setEditable(false);
        editorPane.setBackground(BG_DARK);
        // Start at the top of the report.
        editorPane.setCaretPosition(0); 

        JScrollPane scrollPane = new JScrollPane(editorPane);
        scrollPane.setPreferredSize(new Dimension(1000, 500)); 
        scrollPane.setBorder(BorderFactory.createLineBorder(BORDER_LIGHT));

        JOptionPane.showMessageDialog(this, scrollPane, title, JOptionPane.PLAIN_MESSAGE);
    }

    // Shows a simple loading message while a report is prepared off the Swing event thread.
    private void showReportWithLoading(String loadingMessage, String reportTitle, ReportLoader loader) {
        final JDialog loadingDialog = new JDialog(this, "Loading", true);
        loadingDialog.setDefaultCloseOperation(JDialog.DO_NOTHING_ON_CLOSE);
        loadingDialog.setResizable(false);

        JPanel loadingPanel = new JPanel(new BorderLayout(10, 10));
        loadingPanel.setBackground(PANEL_DARK);
        loadingPanel.setBorder(BorderFactory.createEmptyBorder(20, 25, 20, 25));

        JLabel loadingLabel = new JLabel(loadingMessage, SwingConstants.CENTER);
        loadingLabel.setForeground(TEXT_LIGHT);
        loadingLabel.setFont(new Font("Arial", Font.PLAIN, 15));

        JProgressBar progressBar = new JProgressBar();
        progressBar.setIndeterminate(true);

        loadingPanel.add(loadingLabel, BorderLayout.CENTER);
        loadingPanel.add(progressBar, BorderLayout.SOUTH);
        loadingDialog.add(loadingPanel);
        loadingDialog.pack();
        loadingDialog.setLocationRelativeTo(this);

        SwingWorker<String, Void> worker = new SwingWorker<String, Void>() {
            @Override
            protected String doInBackground() {
                return loader.loadReport();
            }

            @Override
            protected void done() {
                loadingDialog.dispose();
                try {
                    showReportWindow(reportTitle, get());
                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(MotorPH.this,
                            "Unable to open the report: " + ex.getMessage(),
                            "Report Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        };

        worker.execute();
        loadingDialog.setVisible(true);
    }

    // Prompts for a date filter and shows the payroll report for one employee.
    private void handlePayrollFilterRequest(JPanel parentPanel, String empId) {
        String[] options = {"All Time", "Latest Month", "Previous Month", "Custom Date"};
        int choice = JOptionPane.showOptionDialog(parentPanel, "Select the period you want to view:", 
                     "Filter Payslips", JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, 
                     null, options, options[0]);
                     
        String filter = "ALL";
        
        // Map UI selection to internal filter value.
        if (choice == 1) {
            filter = "LATEST";
        } else if (choice == 2) {
            filter = "PREVIOUS";
        } else if (choice == 3) {
            filter = promptForPayrollPeriod(parentPanel);
            if (filter == null) {
                return;
            }
        } else if (choice != 0) {
            return; 
        }
        
        String windowTitle = "Payroll for ID: " + empId + " (";
        if (filter.equals("ALL")) {
            windowTitle = windowTitle + "All Records)";
        } else {
            windowTitle = windowTitle + filter + ")";
        }
        
        final String selectedFilter = filter;
        showReportWithLoading("Loading employee payroll...", windowTitle,
                new ReportLoader() {
                    @Override
                    public String loadReport() {
                        String report = processPayrollLoop(empId, selectedFilter);
                        return HTML_HEADER + report + HTML_FOOTER;
                    }
                });
    }

    // Prompts for a date filter and shows the payroll report for all employees.
    private void handleBulkPayrollFilterRequest(JPanel parentPanel) {
        String[] options = {"All Time", "Latest Month", "Previous Month", "Custom Date"};
        int choice = JOptionPane.showOptionDialog(parentPanel, "Select the period for the bulk report:", 
                     "Filter Bulk Payslips", JOptionPane.DEFAULT_OPTION, JOptionPane.QUESTION_MESSAGE, 
                     null, options, options[0]);
                     
        String filter = "ALL";
        
        // Map UI selection to internal filter value.
        if (choice == 1) {
            filter = "LATEST";
        } else if (choice == 2) {
            filter = "PREVIOUS";
        } else if (choice == 3) {
            filter = promptForPayrollPeriod(parentPanel);
            if (filter == null) {
                return;
            }
        } else if (choice != 0) {
            return; 
        }
        
        String windowTitle = "Company Bulk Payroll Report (";
        if (filter.equals("ALL")) {
            windowTitle = windowTitle + "All Records)";
        } else {
            windowTitle = windowTitle + filter + ")";
        }
        
        final String selectedFilter = filter;
        showReportWithLoading("Loading company payroll summary...", windowTitle,
                new ReportLoader() {
                    @Override
                    public String loadReport() {
                        String report = processAllEmployees(selectedFilter);
                        return HTML_HEADER + report + HTML_FOOTER;
                    }
                });
    }

    // Generates HTML payroll rows for an employee across one or more periods.
    private String processPayrollLoop(String id, String periodFilter) {
        StringBuilder sb = new StringBuilder();
        
        try (FileInputStream fis = new FileInputStream(new File(EXCEL_FILE_PATH));
             Workbook workbook = new XSSFWorkbook(fis)) {
            
            Sheet empSheet = workbook.getSheet("Employee Details");
            Sheet attSheet = workbook.getSheet("Attendance Record");
            
            // Locate the employee master record.
            Row row = null;
            for (Row r : empSheet) {
                if (ExcelRepository.getCellValueAsString(r.getCell(0)).equals(id)) {
                    row = r;
                    break;
                }
            }

            if (row == null) {
                return "<p style='color:#FCA5A5;'>Employee ID not found.</p>";
            }

            // Determine hourly-rate column from header (fallback to column 18).
            int hourlyColIndex = 18; 
            Row headerRow = empSheet.getRow(0);
            if (headerRow != null) {
                for (Cell cell : headerRow) {
                    if (ExcelRepository.getCellValueAsString(cell).toLowerCase().contains("hourly")) {
                        hourlyColIndex = cell.getColumnIndex();
                        break;
                    }
                }
            }

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            List<String> activePeriods = getUniquePeriods(attSheet, id);
            
            // Translate friendly filters to a single period when needed.
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

            // If a single period is requested, narrow the list.
            if (!periodFilter.equals("ALL")) {
                List<String> filteredList = new ArrayList<>();
                if (activePeriods.contains(periodFilter)) {
                    filteredList.add(periodFilter);
                }
                activePeriods = filteredList;
            }
            
            String name = ExcelRepository.getCellValueAsString(row.getCell(2)) + " " + ExcelRepository.getCellValueAsString(row.getCell(1));
            
            sb.append("<h2 style='color: #93C5FD; border-bottom: 1px solid #93C5FD; padding-bottom: 5px;'>Payroll Summary for: [").append(id).append("] ").append(name).append("</h2>");
            
            if (activePeriods.size() == 0) {
                sb.append("<p>No payroll records are available for this employee during the selected period.</p>");
                return sb.toString();
            }

            sb.append("<table border='1' cellpadding='10' cellspacing='0' style='border-collapse: collapse; width: 100%; border-color: #64748B;'>");
            
            sb.append("<tr style='background-color: #334155; color: #93C5FD;'>");

            sb.append("<th>Period</th><th>1st Cutoff</th><th>1st NET</th><th>2nd Cutoff</th><th>Deductions</th><th>Tax</th><th>2nd NET</th><th>TOTAL NET PAY</th>");
            sb.append("</tr>");

            for (String period : activePeriods) {
                double h1 = calculateHours(attSheet, id, period, 1, 15);
                double h2 = calculateHours(attSheet, id, period, 16, 31);
                double totalHours = h1 + h2;

                if (totalHours > 0) {
                    // If hourly rate is stored as monthly, normalize to hourly.
                    double hourlyRate = ExcelRepository.getNumericSafe(row.getCell(hourlyColIndex), evaluator); 
                    if (hourlyRate > 1000) {
                        hourlyRate = hourlyRate / 160;
                    }

                    // Ensure a basic salary value for deductions and taxes.
                    double basicSalary = ExcelRepository.getNumericSafe(row.getCell(13), evaluator); 
                    if (basicSalary <= 0) {
                        basicSalary = hourlyRate * 160;
                    }

                    double gross1 = hourlyRate * h1;
                    double gross2 = hourlyRate * h2;
                    double totalGross = gross1 + gross2;
                    
                    // Standard government deductions.
                    double sss = PayrollService.calculateSSS(basicSalary);
                    double philHealth = basicSalary * 0.025; 
                    if (philHealth > 2500.00) {
                        philHealth = 2500.00; 
                    }
                    double pagIbig = 200.00; 
                    
                    double totalGovtDeductions = sss + philHealth + pagIbig;
                    double taxableIncome = totalGross - totalGovtDeductions;
                    double tax = PayrollService.calculateTax(taxableIncome);
                    
                    // Apply deductions on second cutoff (simplified business rule).
                    double net1 = gross1; 
                    double net2 = gross2 - (totalGovtDeductions + tax);
                    
                    // If deductions exceed second cutoff, carry back to first.
                    if (net2 < 0) {
                        net1 = net1 + net2; 
                        net2 = 0;           
                    }
                    
                    double totalNetPay = net1 + net2;

                    sb.append("<tr style='text-align: center; background-color: #172033;'>");
                    sb.append("<td>").append(period).append("</td>");
                    sb.append("<td>").append(String.format("%,.2f", gross1)).append("</td>");
                    sb.append("<td>").append(String.format("%,.2f", net1)).append("</td>");
                    sb.append("<td>").append(String.format("%,.2f", gross2)).append("</td>");
                    
                    sb.append("<td style='color: #FCA5A5;'>-").append(String.format("%,.2f", totalGovtDeductions)).append("</td>");
                    sb.append("<td style='color: #FCA5A5;'>-").append(String.format("%,.2f", tax)).append("</td>");
                    
                    sb.append("<td>").append(String.format("%,.2f", net2)).append("</td>");
                    
                    sb.append("<td style='color: #86EFAC; font-weight: bold;'>").append(String.format("%,.2f", totalNetPay)).append("</td>");
                    
                    sb.append("</tr>");
                }
            }
            sb.append("</table><br>");

        } catch (Exception e) {
            sb.append("<p style='color:#FCA5A5;'>Error processing payroll records: ").append(e.getMessage()).append("</p>");
        }
        return sb.toString();
    }

    // Generates a bulk HTML report for all employees in a period.
    private String processAllEmployees(String periodFilter) {
        StringBuilder sb = new StringBuilder();
        StringBuilder employeeRows = new StringBuilder();

        int totalEmployeesProcessed = 0;
        double totalGrossPay = 0.0;
        double totalDeductions = 0.0;
        double totalNetPay = 0.0;

        try (FileInputStream fis = new FileInputStream(new File(EXCEL_FILE_PATH));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet empSheet = workbook.getSheet("Employee Details");
            Sheet attSheet = workbook.getSheet("Attendance Record");

            if (empSheet == null || attSheet == null) {
                return "<p style='color:#FCA5A5;'>Required payroll sheets were not found in the Excel file.</p>";
            }

            // Find the hourly-rate column using the same rule as single payroll.
            int hourlyColIndex = 18;
            Row headerRow = empSheet.getRow(0);
            if (headerRow != null) {
                for (Cell cell : headerRow) {
                    if (ExcelRepository.getCellValueAsString(cell).toLowerCase().contains("hourly")) {
                        hourlyColIndex = cell.getColumnIndex();
                        break;
                    }
                }
            }

            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            for (Row row : empSheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }

                String id = ExcelRepository.getCellValueAsString(row.getCell(0));
                if (id.isEmpty()) {
                    continue;
                }

                List<String> periodsToProcess = getBulkPayrollPeriods(attSheet, id, periodFilter);
                if (periodsToProcess.isEmpty()) {
                    continue;
                }

                String name = ExcelRepository.getCellValueAsString(row.getCell(2)) + " "
                        + ExcelRepository.getCellValueAsString(row.getCell(1));

                double hourlyRate = ExcelRepository.getNumericSafe(row.getCell(hourlyColIndex), evaluator);
                if (hourlyRate > 1000) {
                    hourlyRate = hourlyRate / 160;
                }

                double basicSalary = ExcelRepository.getNumericSafe(row.getCell(13), evaluator);
                if (basicSalary <= 0) {
                    basicSalary = hourlyRate * 160;
                }

                double employeeHours = 0.0;
                double employeeGrossPay = 0.0;
                double employeeDeductions = 0.0;
                double employeeNetPay = 0.0;
                boolean hasPayrollRecord = false;

                for (String period : periodsToProcess) {
                    double h1 = calculateHours(attSheet, id, period, 1, 15);
                    double h2 = calculateHours(attSheet, id, period, 16, 31);
                    double periodHours = h1 + h2;

                    if (periodHours <= 0) {
                        continue;
                    }

                    hasPayrollRecord = true;

                    // Use the same payroll formulas as the single employee report.
                    double gross1 = hourlyRate * h1;
                    double gross2 = hourlyRate * h2;
                    double periodGrossPay = gross1 + gross2;

                    double sss = PayrollService.calculateSSS(basicSalary);
                    double philHealth = basicSalary * 0.025;
                    if (philHealth > 2500.00) {
                        philHealth = 2500.00;
                    }
                    double pagIbig = 200.00;

                    double governmentDeductions = sss + philHealth + pagIbig;
                    double taxableIncome = periodGrossPay - governmentDeductions;
                    double tax = PayrollService.calculateTax(taxableIncome);

                    double net1 = gross1;
                    double net2 = gross2 - (governmentDeductions + tax);
                    if (net2 < 0) {
                        net1 = net1 + net2;
                        net2 = 0;
                    }
                    double periodNetPay = net1 + net2;

                    employeeHours = employeeHours + periodHours;
                    employeeGrossPay = employeeGrossPay + periodGrossPay;
                    employeeDeductions = employeeDeductions + governmentDeductions + tax;
                    employeeNetPay = employeeNetPay + periodNetPay;
                }

                // Employees without usable attendance are not counted as processed.
                if (!hasPayrollRecord) {
                    continue;
                }

                totalEmployeesProcessed = totalEmployeesProcessed + 1;
                totalGrossPay = totalGrossPay + employeeGrossPay;
                totalDeductions = totalDeductions + employeeDeductions;
                totalNetPay = totalNetPay + employeeNetPay;

                employeeRows.append("<tr style='text-align: center; background-color: #172033;'>");
                employeeRows.append("<td>").append(id).append("</td>");
                employeeRows.append("<td style='text-align: left;'>").append(name).append("</td>");
                employeeRows.append("<td>").append(String.format("%,.2f", hourlyRate)).append("</td>");
                employeeRows.append("<td>").append(String.format("%,.2f", employeeHours)).append("</td>");
                employeeRows.append("<td>").append(String.format("%,.2f", employeeGrossPay)).append("</td>");
                employeeRows.append("<td style='color: #FCA5A5;'>").append(String.format("%,.2f", employeeDeductions)).append("</td>");
                employeeRows.append("<td style='color: #86EFAC; font-weight: bold;'>").append(String.format("%,.2f", employeeNetPay)).append("</td>");
                employeeRows.append("</tr>");
            }

            sb.append("<h2 style='color: #93C5FD; border-bottom: 1px solid #93C5FD; padding-bottom: 5px;'>Company-Wide Payroll Summary</h2>");

            if (totalEmployeesProcessed == 0) {
                sb.append("<p>No employees were processed because the selected period has no usable payroll or attendance records.</p>");
                return sb.toString();
            }

            double averageNetPay = totalNetPay / totalEmployeesProcessed;

            sb.append("<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse; width: 100%; border-color: #64748B;'>");
            sb.append("<tr style='background-color: #334155; color: #93C5FD;'>");
            sb.append("<th>Total Employees Processed</th><th>Total Gross Pay</th><th>Total Deductions</th><th>Total Net Pay</th><th>Average Net Pay</th>");
            sb.append("</tr>");
            sb.append("<tr style='text-align: center; background-color: #172033;'>");
            sb.append("<td>").append(totalEmployeesProcessed).append("</td>");
            sb.append("<td>").append(String.format("%,.2f", totalGrossPay)).append("</td>");
            sb.append("<td style='color: #FCA5A5;'>").append(String.format("%,.2f", totalDeductions)).append("</td>");
            sb.append("<td style='color: #86EFAC; font-weight: bold;'>").append(String.format("%,.2f", totalNetPay)).append("</td>");
            sb.append("<td>").append(String.format("%,.2f", averageNetPay)).append("</td>");
            sb.append("</tr></table><br>");

            sb.append("<h3 style='color: #93C5FD;'>Employee Payroll Records</h3>");
            sb.append("<table border='1' cellpadding='8' cellspacing='0' style='border-collapse: collapse; width: 100%; border-color: #64748B;'>");
            sb.append("<tr style='background-color: #334155; color: #93C5FD;'>");
            sb.append("<th>Employee Number</th><th>Name</th><th>Rate</th><th>Hours Worked</th><th>Gross Pay</th><th>Deductions</th><th>Net Pay</th>");
            sb.append("</tr>");
            sb.append(employeeRows);
            sb.append("</table><br>");
        } catch (Exception e) {
            sb.append("<p style='color:#FCA5A5;'>Error processing company payroll: ").append(e.getMessage()).append("</p>");
        }
        return sb.toString();
    }

    // Selects the payroll periods for one employee in the bulk report.
    private static List<String> getBulkPayrollPeriods(Sheet sheet, String id, String periodFilter) {
        List<String> activePeriods = getUniquePeriods(sheet, id);
        List<String> selectedPeriods = new ArrayList<>();
        String selectedFilter = periodFilter == null ? "" : periodFilter.trim();

        if (selectedFilter.equals("ALL")) {
            selectedPeriods.addAll(activePeriods);
        } else if (selectedFilter.equals("LATEST")) {
            if (!activePeriods.isEmpty()) {
                selectedPeriods.add(activePeriods.get(activePeriods.size() - 1));
            }
        } else if (selectedFilter.equals("PREVIOUS")) {
            if (activePeriods.size() > 1) {
                selectedPeriods.add(activePeriods.get(activePeriods.size() - 2));
            } else if (activePeriods.size() == 1) {
                selectedPeriods.add(activePeriods.get(0));
            }
        } else if (activePeriods.contains(selectedFilter)) {
            selectedPeriods.add(selectedFilter);
        }

        return selectedPeriods;
    }

    // Extracts distinct MM/YYYY periods from attendance records.
    private static List<String> getUniquePeriods(Sheet sheet, String id) {
        Set<String> periods = new LinkedHashSet<>();
        for (Row row : sheet) {
            if (ExcelRepository.getCellValueAsString(row.getCell(0)).equals(id)) {
                String date = ExcelRepository.getCellValueAsString(row.getCell(3));
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

    // Totals hours for an employee within a half-month range.
    private static double calculateHours(Sheet sheet, String id, String targetPeriod, int startDay, int endDay) {
        double total = 0;
        LocalTime shiftStart = LocalTime.of(8, 0);
        LocalTime shiftEnd = LocalTime.of(17, 0);
        LocalTime grace = LocalTime.of(8, 10);

        for (Row row : sheet) {
            if (ExcelRepository.getCellValueAsString(row.getCell(0)).equals(id)) {
                String date = ExcelRepository.getCellValueAsString(row.getCell(3));
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
                                LocalTime timeIn = parseTime(ExcelRepository.getCellValueAsString(row.getCell(4)));
                                LocalTime timeOut = parseTime(ExcelRepository.getCellValueAsString(row.getCell(5)));
                                
                                // Apply grace period for time-in.
                                if (timeIn.isBefore(shiftStart) || (timeIn.isAfter(shiftStart) && timeIn.isBefore(grace))) {
                                    timeIn = shiftStart;
                                }
                                // Cap time-out at shift end.
                                if (timeOut.isAfter(shiftEnd)) {
                                    timeOut = shiftEnd;
                                }
                                
                                if (timeOut.isAfter(timeIn)) {
                                    double duration = Duration.between(timeIn, timeOut).toMinutes() / 60.0;
                                    if (duration > 5) {
                                        // Deduct one hour for a standard lunch break.
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

    // Parses time strings with or without AM/PM, falling back to 8:00 AM.
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

    // Application entry point.
    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                // Create and show the main UI on the EDT.
                new MotorPH().setVisible(true);
            }
        });
    }
}
