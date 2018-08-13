package org.apache.fineract.cn.datamigration.util;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.fineract.cn.office.api.v1.domain.ContactDetail;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.stream.IntStream;

public class EmployeeSheetGenerator {
  public static String identifier =RandomStringUtils.randomAlphanumeric(10);
  public static String givenName=RandomStringUtils.randomAlphanumeric(10);
  public static String middleName=RandomStringUtils.randomAlphanumeric(9);
  public static String surname=RandomStringUtils.randomAlphanumeric(9);
  public static String assignedOffice=OfficeSheetGenerator.identifier;
  public static String type=ContactDetail.Type.MOBILE.name();
  public static String group=ContactDetail.Group.PRIVATE.name();
  public static  String value=RandomStringUtils.randomAlphabetic(12);
  public static String preferenceLevel=String.valueOf(1);


  //generating office excel sheet with some data
  public static void employeeSheetDownload(){

    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Employees");

    int startRowIndex = 0;
    int startColIndex = 0;

    Font font = worksheet.getWorkbook().createFont();
    XSSFCellStyle headerCellStyle = worksheet.getWorkbook().createCellStyle();

    headerCellStyle.setWrapText(true);
    headerCellStyle.setFont(font);
    XSSFRow rowHeader = worksheet.createRow((short) startRowIndex);
    XSSFRow firstRow = worksheet.createRow( 1);
    rowHeader.setHeight((short) 500);

    XSSFCell cell1 = rowHeader.createCell(startColIndex+0);
    cell1.setCellValue("Identifier");
    cell1.setCellStyle(headerCellStyle);

    XSSFCell cell2 = rowHeader.createCell(startColIndex+1);
    cell2.setCellValue("Given Name");
    cell2.setCellStyle(headerCellStyle);


    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Middle Name");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Surname");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5 = rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Assigned Office ");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6 = rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Type ");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7 = rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Group ");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8 = rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Value ");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9= rowHeader.createCell(startColIndex+8);
    cell9.setCellValue("Preference Level");
    cell9.setCellStyle(headerCellStyle);

    XSSFCell cell11 = firstRow.createCell(startColIndex+0);
    cell11.setCellValue(identifier);

    XSSFCell cell22 = firstRow.createCell(startColIndex+1);
    cell22.setCellValue(givenName);

    XSSFCell cell33 = firstRow.createCell(startColIndex+2);
    cell33.setCellValue(middleName);

    XSSFCell cell44 = firstRow.createCell(startColIndex+3);
    cell44.setCellValue(surname);

    XSSFCell cell55= firstRow.createCell(startColIndex+4);
    cell55.setCellValue(assignedOffice);

    XSSFCell cell66= firstRow.createCell(startColIndex+5);
    cell66.setCellValue(type);

    XSSFCell cell77= firstRow.createCell(startColIndex+6);
    cell77.setCellValue(group);

    XSSFCell cell88= firstRow.createCell(startColIndex+7);
    cell88.setCellValue(value);

    XSSFCell cell99= firstRow.createCell(startColIndex+8);
    cell99.setCellValue(preferenceLevel);

    IntStream.range(0, 9).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));


    try {
      // Retrieve the output stream
      FileOutputStream fos = new FileOutputStream(new File("/home/le/Desktop/Employees.xlsx"));
      worksheet.getWorkbook().write(fos);
      fos.flush();
      fos.close();
    } catch (Exception e) {
      System.out.println("Unable to write report to the output stream");
    }
  }
}
