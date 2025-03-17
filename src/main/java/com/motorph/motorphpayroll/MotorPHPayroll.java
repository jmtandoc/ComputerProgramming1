/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.motorph.motorphpayroll;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Scanner;

/**
 *
 * @author jhoan
 */

public class MotorPHPayroll {
 public static void main(String[] args) {
        String filePath = "src/MotorPH Employee Data.xlsx";
        Scanner scanner = new Scanner(System.in);
        
        System.out.print("Enter Employee ID: ");
        String employeeID = scanner.nextLine().trim();
        
        System.out.print("Enter month of compensation (e.g., January): ");
        String month = scanner.nextLine().trim();
        
        fetchEmployeeDetails(filePath, employeeID, month);
        scanner.close();
    }

    public static void fetchEmployeeDetails(String filePath, String employeeID, String month) {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet detailsSheet = workbook.getSheet("Employee Details");
            
            String fullName = "Not Found";
            String birthday = "Not Found";
            double basicSalary = 0;
            double riceSubsidy = 0;
            double phoneAllowance = 0;
            double clothingAllowance = 0;
            
            for (Row row : detailsSheet) {
                if (row.getCell(0) != null) {
                    String idFromExcel = getStringValue(row.getCell(0));
                    if (idFromExcel.equals(employeeID)) {
                        fullName = getStringValue(row.getCell(2)) + " " + getStringValue(row.getCell(1));
                        birthday = getFormattedDate(row.getCell(3));
                        basicSalary = getNumericValue(row.getCell(13));
                        riceSubsidy = getNumericValue(row.getCell(14));
                        phoneAllowance = getNumericValue(row.getCell(15));
                        clothingAllowance = getNumericValue(row.getCell(16));
                        break;
                    }
                }
            }
            
            System.out.println("Employee Name: " + fullName);
            System.out.println("Birthday: " + birthday);
            
            double hourlyRate = (basicSalary / 21) / 8; 
            calculateWorkedHours(filePath, employeeID, month, fullName, basicSalary, hourlyRate);
        } catch (IOException e) {
            System.err.println("Error reading the file: " + e.getMessage());
        }
    }
    
    public static void calculateWorkedHours(String filePath, String employeeID, String month, String fullName, double basicSalary, double hourlyRate) {
        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet attendanceSheet = workbook.getSheet("Attendance Record");
            
            Map<Integer, Double> weeklyHours = new HashMap<>();
            double totalMinutes = 0;
            SimpleDateFormat excelDateFormat = new SimpleDateFormat("MM-dd-yyyy");
            Calendar calendar = Calendar.getInstance();
            
            for (Row row : attendanceSheet) {
                if (row.getCell(0) != null) {
                    String idFromExcel = getStringValue(row.getCell(0));
                    Cell dateCell = row.getCell(3);
                    Date date = null;
                    
                    if (idFromExcel.equals(employeeID)) {
                        try {
                            if (dateCell.getCellType() == CellType.NUMERIC) {
                                date = dateCell.getDateCellValue();
                            } else {
                                date = excelDateFormat.parse(dateCell.getStringCellValue());
                            }
                        } catch (Exception e) {
                            continue;
                        }
                        
                        calendar.setTime(date);
                        int weekNumber = calendar.get(Calendar.WEEK_OF_MONTH);
                        String monthFromDate = new SimpleDateFormat("MMMM").format(date);
                        
                        if (monthFromDate.equalsIgnoreCase(month)) {
                            Cell loginCell = row.getCell(4);
                            Cell logoutCell = row.getCell(5);
                            
                            LocalTime loginTime = parseTime(loginCell);
                            LocalTime logoutTime = parseTime(logoutCell);
                            
                            if (loginTime != null && logoutTime != null) {
                                if (Duration.between(LocalTime.of(8, 0), loginTime).toMinutes() <= 10) {
                                    loginTime = LocalTime.of(8, 0);
                                }
                                
                                Duration duration = Duration.between(loginTime, logoutTime);
                                double hoursWorked = duration.toMinutes() / 60.0;
                                totalMinutes += duration.toMinutes();
                                weeklyHours.put(weekNumber, weeklyHours.getOrDefault(weekNumber, 0.0) + hoursWorked);
                            }
                        }
                    }
                }
            }
            
            double totalHours = totalMinutes / 60.0;
            double totalGrossPay = totalHours * hourlyRate;
            double sssContribution = calculateSSSContribution(basicSalary);
            double philhealthContribution = basicSalary * 0.03;
            double pagibigContribution = basicSalary < 1500 ? basicSalary * 0.01 : basicSalary * 0.02;
            double taxableIncome = totalGrossPay - (sssContribution + philhealthContribution + pagibigContribution);
            double withholdingTax = calculateWithholdingTax(taxableIncome);
            double netPay = (totalGrossPay - withholdingTax);
            double weeklyNetPay = netPay/4;
            System.out.println("Month: " + month);
            
            for (int week = 1; week <= 4; week++) {
                double weeklyHoursWorked = weeklyHours.getOrDefault(week, 0.0);
                double weeklyGrossPay = weeklyHoursWorked * hourlyRate;
                
            System.out.printf("Week %d Worked Hours: %.2f hours | Gross Pay: %.2f\n", week, weeklyHoursWorked, weeklyGrossPay);}
            System.out.printf("Total Monthly Worked Hours: %.2f hours | Total Gross Pay: %.2f\n", totalHours, totalGrossPay);   
            System.out.printf("SSS Contribution: %.2f\n", sssContribution);
            System.out.printf("PhilHealth Contribution: %.2f\n", philhealthContribution);
            System.out.printf("Pag-IBIG Contribution: %.2f\n", pagibigContribution);
            System.out.printf("Withholding Tax: %.2f\n", withholdingTax);
            System.out.printf("NetPay: %.2f\n", netPay);
            System.out.printf("Estimated Weekly NetPay: %.2f\n", weeklyNetPay);

        } catch (IOException e) {
            System.err.println("Error reading the file: " + e.getMessage());
        }
    }
    
    private static String getStringValue(Cell cell) {
        if (cell == null) return "";
        return cell.getCellType() == CellType.STRING ? cell.getStringCellValue().trim() : String.valueOf((int) cell.getNumericCellValue());
    }

    private static double getNumericValue(Cell cell) {
        if (cell == null) return 0;
        return cell.getCellType() == CellType.NUMERIC ? cell.getNumericCellValue() : 0;
    }
    
    private static String getFormattedDate(Cell cell) {
        if (cell == null || cell.getCellType() != CellType.NUMERIC) return "Not Found";
        Date date = cell.getDateCellValue();
        SimpleDateFormat sdf = new SimpleDateFormat("dd-MMM-yyyy");
        return sdf.format(date);
    }
    
    private static LocalTime parseTime(Cell cell) {
        if (cell == null) return null;
        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                double excelTime = cell.getNumericCellValue();
                int hours = (int) (excelTime * 24);
                int minutes = (int) ((excelTime * 1440) % 60);
                return LocalTime.of(hours, minutes);
            } else if (cell.getCellType() == CellType.STRING) {
                return LocalTime.parse(cell.getStringCellValue().trim(), DateTimeFormatter.ofPattern("hh:mm a"));
            }
        } catch (Exception e) {
            System.err.println("Error parsing time from cell: " + cell.toString());
            return null;
        }
        return null;
    }

    private static double calculateSSSContribution(double basicSalary) {
         if (basicSalary < 3250) return 135.00;
        else if (basicSalary < 3750) return 157.50;
        else if (basicSalary < 4250) return 180.00;
        else if (basicSalary < 4750) return 202.50;
        else if (basicSalary < 5250) return 225.00;
        else if (basicSalary < 5750) return 247.50;
        else if (basicSalary < 6250) return 270.00;
        else if (basicSalary < 6750) return 292.50;
        else if (basicSalary < 7250) return 315.00;
        else if (basicSalary < 7750) return 337.50;
        else if (basicSalary < 8250) return 360.00;
        else if (basicSalary < 8750) return 382.50;
        else if (basicSalary < 9250) return 405.00;
        else if (basicSalary < 9750) return 427.50;
        else if (basicSalary < 10250) return 450.00;
        else if (basicSalary < 10750) return 472.50;
        else if (basicSalary < 11250) return 495.00;
        else if (basicSalary < 11750) return 517.50;
        else if (basicSalary < 12250) return 540.00;
        else if (basicSalary < 12750) return 562.50;
        else if (basicSalary < 13250) return 585.00;
        else if (basicSalary < 13750) return 607.50;
        else if (basicSalary < 14250) return 630.00;
        else if (basicSalary < 14750) return 652.50;
        else if (basicSalary < 15250) return 675.00;
        else if (basicSalary < 15750) return 697.50;
        else if (basicSalary < 16250) return 720.00;
        else if (basicSalary < 16750) return 742.50;
        else if (basicSalary < 16250) return 765.00;
        else if (basicSalary < 16750) return 787.50;
        else if (basicSalary < 17250) return 810.00;
        else if (basicSalary < 17750) return 832.50;
        else if (basicSalary < 18250) return 855.00;
        else if (basicSalary < 18750) return 877.50;
        else if (basicSalary < 19250) return 900.00;
        else if (basicSalary < 19750) return 922.50;
        else if (basicSalary < 20250) return 945.00;
        else if (basicSalary < 20750) return 967.50;
        else if (basicSalary < 21250) return 990.00;
        else if (basicSalary < 21750) return 1012.50;
        else if (basicSalary < 22250) return 1035.00;
        else if (basicSalary < 22750) return 1057.50;
        else if (basicSalary < 23250) return 1080.00;
        else if (basicSalary < 23750) return 1102.50;
        else return 1125.00;
    }

    private static double calculateWithholdingTax(double taxableIncome) {
        if (taxableIncome < 20832) return 0;
        else if (taxableIncome < 33333) return (taxableIncome - 20832) * 0.20;
        else if (taxableIncome < 66667) return (taxableIncome - 33333) * 0.25 + 2500;
        else if (taxableIncome < 166667) return (taxableIncome - 66667) * 0.30 + 10833;
        else if (taxableIncome < 666667) return (taxableIncome - 166667) * 0.32 + 40833.33;
        else return (taxableIncome - 166667) * 0.35 + 200833.33;
    }
}