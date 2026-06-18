package model;

// Stores the calculated payroll totals and deductions for an employee.
public class PayrollResult {
    private String employeeId;
    private String employeeName;
    private double totalHoursWorked;
    private double grossPay;
    private double sssDeduction;
    private double philHealthDeduction;
    private double pagIbigDeduction;
    private double withholdingTax;
    private double netPay;

    public PayrollResult(String employeeId, String employeeName, double totalHoursWorked,
                         double grossPay, double sssDeduction, double philHealthDeduction,
                         double pagIbigDeduction, double withholdingTax, double netPay) {
        this.employeeId = employeeId;
        this.employeeName = employeeName;
        this.totalHoursWorked = totalHoursWorked;
        this.grossPay = grossPay;
        this.sssDeduction = sssDeduction;
        this.philHealthDeduction = philHealthDeduction;
        this.pagIbigDeduction = pagIbigDeduction;
        this.withholdingTax = withholdingTax;
        this.netPay = netPay;
    }

    public String getEmployeeId() {
        return employeeId;
    }

    public String getEmployeeName() {
        return employeeName;
    }

    public double getTotalHoursWorked() {
        return totalHoursWorked;
    }

    public double getGrossPay() {
        return grossPay;
    }

    public double getSssDeduction() {
        return sssDeduction;
    }

    public double getPhilHealthDeduction() {
        return philHealthDeduction;
    }

    public double getPagIbigDeduction() {
        return pagIbigDeduction;
    }

    public double getWithholdingTax() {
        return withholdingTax;
    }

    public double getNetPay() {
        return netPay;
    }
}
