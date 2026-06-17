package model;

import java.time.LocalDate;
import java.time.LocalTime;

public class AttendanceRecord {
    private String employeeId;
    private LocalDate date;
    private LocalTime timeIn;
    private LocalTime timeOut;

    public AttendanceRecord(String employeeId, LocalDate date, LocalTime timeIn, LocalTime timeOut) {
        this.employeeId = employeeId;
        this.date = date;
        this.timeIn = timeIn;
        this.timeOut = timeOut;
    }

    public String getEmployeeId() {
        return employeeId;
    }

    public LocalDate getDate() {
        return date;
    }

    public LocalTime getTimeIn() {
        return timeIn;
    }

    public LocalTime getTimeOut() {
        return timeOut;
    }
}
