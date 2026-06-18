package service;

// Provides payroll calculations that are separate from the Swing user interface.
public class PayrollService {
    // Calculates the SSS deduction using the system's simplified salary brackets.
    public static double calculateSSS(double salary) {
        if (salary <= 3250) return 135.00;
        if (salary >= 24750) return 1125.00;
        int steps = (int)((salary - 3250 - 0.01) / 500) + 1;
        return 135.00 + (steps * 22.50);
    }

    // Calculates withholding tax using the system's simplified tax brackets.
    public static double calculateTax(double taxableIncome) {
        if (taxableIncome <= 20833) return 0;
        if (taxableIncome <= 33333) return (taxableIncome - 20833) * 0.15;
        if (taxableIncome <= 66667) return 1875 + (taxableIncome - 33333) * 0.20;
        if (taxableIncome <= 166667) return 8541.67 + (taxableIncome - 66667) * 0.25;
        return 33541.67 + (taxableIncome - 166667) * 0.30;
    }
}
