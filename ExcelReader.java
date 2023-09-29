import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class ExcelReader {

    public static void main(String[] args) {
        try {
            // Load the Excel file
            FileInputStream excelFile = new FileInputStream(new File("Assignment_Timecard.xlsx"));
            Workbook workbook = new XSSFWorkbook(excelFile);

            // Get the first sheet in the Excel file
            Sheet sheet = workbook.getSheetAt(0);

            Set<String> printedEmployees = new HashSet<>();
            Set<String> employess = new HashSet<>();
            Map<String, Date> employeeShifts = new HashMap<>();

            // if worked for 7 consecutive days
            // Iterate through rows
            System.out.println("-------------------------------(1)---------------------");
            System.out.println("Employess who have worked for 7 consecutive days:");
            for (Row row : sheet) {
                // Skip the header row
                if (row.getRowNum() == 0) {
                    continue;
                }

                //Skip row with blank cells
                if(row.getCell(2).getCellType()==Cell.CELL_TYPE_STRING ){
                    continue;
                }
                if(row.getCell(3).getCellType()==Cell.CELL_TYPE_STRING ){
                    continue;
                }

                // Get the date cell
                Cell dateCell = row.getCell(2); // Assuming the date is in the first column
                Cell dateCell2 = row.getCell(3);
                Cell n=row.getCell(7);//for name cell


                if (DateUtil.isCellDateFormatted(dateCell)) {

                    Date timeIn = dateCell.getDateCellValue();
                    Date timeOut = dateCell2.getDateCellValue();

                    DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
                    String nameEmp=n.getStringCellValue();


                    boolean consecutiveDays = checkConsecutiveDays(sheet, row.getRowNum(), timeIn,nameEmp);

                    if (consecutiveDays) {
                        // Print the name and position
                        String name = row.getCell(7).getStringCellValue();
                        String position = row.getCell(0).getStringCellValue();

                        String employeeKey = name + "-" + position;
                        if (!printedEmployees.contains(employeeKey)) {

                            System.out.println("Employee Name: " + name +"||"+" Position "+position);


                            // Add the employee to the set to indicate it has been printed
                            printedEmployees.add(employeeKey);
                        }

                    }
                }
            }
            System.out.println("-----------------------(2)----------------");
            System.out.println("Employees who have less than 10 hours of time between shifts but greater than 1 hour:");
            for (Row row : sheet) {
                // Skip the header row
                if (row.getRowNum() == 0) {
                    continue;
                }
                //Skip row with blank cells
                if(row.getCell(2).getCellType()==Cell.CELL_TYPE_STRING ){
                    continue;
                }
                if(row.getCell(3).getCellType()==Cell.CELL_TYPE_STRING ){
                    continue;
                }

                // Get the date cell
                Cell timeInCell = row.getCell(2); // Assuming the date is in the first column
                Cell timeOutCell = row.getCell(3);
                Cell nameCell = row.getCell(7);//for name cell

                if (nameCell != null && timeInCell != null && timeOutCell != null ) {
                    String name = nameCell.getStringCellValue();

                    Date timeIn = timeInCell.getDateCellValue();
                    Date timeOut = timeOutCell.getDateCellValue();

                    if (employeeShifts.containsKey(name)) {
                        Date previousShiftTimeOut = employeeShifts.get(name);
                        long hoursBetweenShifts = (timeIn.getTime() - previousShiftTimeOut.getTime()) / (60 * 60 * 1000);

                        if (hoursBetweenShifts > 1 && hoursBetweenShifts < 10) {
                            System.out.println("Employee Name: " + name + " || Position: " + row.getCell(0).getStringCellValue());
                        }
                    }

                    // Store the current shift's time out for the next iteration
                    employeeShifts.put(name, timeOut);
                }
            }

            System.out.println("-----------------------(3)----------------");
            System.out.println("Employess who have worked for more than 14 hours in a single shift:");
            //Who has worked for more than 14 hours in a single shift
            for (Row row : sheet) {
                // Skip the header row
                if (row.getRowNum() == 0) {
                    continue;
                }
                //Skip row with blank cells
                if(row.getCell(2).getCellType()==Cell.CELL_TYPE_STRING ){
                    continue;
                }
                if(row.getCell(3).getCellType()==Cell.CELL_TYPE_STRING ){
                    continue;
                }
                // Get the date cell
                Cell dateCell = row.getCell(2); // Assuming the date is in the first column
                Cell dateCell2 = row.getCell(3);
                Cell n = row.getCell(7);//for name cell


                if (DateUtil.isCellDateFormatted(dateCell)) {
//                    Date date = dateCell.getDateCellValue();
                    Date timeIn = dateCell.getDateCellValue();
                    Date timeOut = dateCell2.getDateCellValue();

                    DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
                    String nameEmp = n.getStringCellValue();
                    boolean fourteenHrsShift=checkIfFourteenHrs(timeIn,timeOut);
                    //Who has worked for more than 14 hours in a single shift
                    if(fourteenHrsShift){

                        String name = row.getCell(7).getStringCellValue();
                        String employeeKey1 = name;
                        String position = row.getCell(0).getStringCellValue();

                        if(!employess.contains(employeeKey1)){

                            System.out.println("Employee Name: " + name +" || Position: "+position);
                            employess.add(employeeKey1);
                        }
                    }
                }
            }





                // Close the workbook and file
            excelFile.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }



    private static boolean checkIfFourteenHrs(Date timeIn, Date timeOut) {

        long hoursWorked = (timeOut.getTime() - timeIn.getTime()) / (60 * 60 * 1000);

        if (hoursWorked > 14) {
            return true;
        }
        return false;
    }


    //check if worked for 7 consecutive days
    private static boolean checkConsecutiveDays(Sheet sheet, int rowNum, Date startDate, String nameEmp) {
        int consecutiveDaysCount = 1; // Initialize with 1 as we already have the first day\
        for(int i=rowNum+1;i<sheet.getLastRowNum();i++){
            Cell emp=sheet.getRow(i).getCell(7);
            String empName=emp.getStringCellValue();

            if(nameEmp==empName){
                consecutiveDaysCount++;
            }
            if(consecutiveDaysCount==7){
                return true;
            }
        }
    return false;

    }
    }



