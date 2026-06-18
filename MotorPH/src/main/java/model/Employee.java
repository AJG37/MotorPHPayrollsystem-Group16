package model;

// Stores an employee's basic personal and salary information.
public class Employee {
    private String employeeId;
    private String lastName;
    private String firstName;
    private String birthday;
    private String status;
    private String position;
    private double basicSalary;
    private double hourlyRate;

    public Employee(String employeeId, String lastName, String firstName, String birthday,
                    String status, String position, double basicSalary, double hourlyRate) {
        this.employeeId = employeeId;
        this.lastName = lastName;
        this.firstName = firstName;
        this.birthday = birthday;
        this.status = status;
        this.position = position;
        this.basicSalary = basicSalary;
        this.hourlyRate = hourlyRate;
    }

    public String getEmployeeId() {
        return employeeId;
    }

    public String getLastName() {
        return lastName;
    }

    public String getFirstName() {
        return firstName;
    }

    public String getBirthday() {
        return birthday;
    }

    public String getStatus() {
        return status;
    }

    public String getPosition() {
        return position;
    }

    public double getBasicSalary() {
        return basicSalary;
    }

    public double getHourlyRate() {
        return hourlyRate;
    }

    public String getFullName() {
        return firstName + " " + lastName;
    }
}
