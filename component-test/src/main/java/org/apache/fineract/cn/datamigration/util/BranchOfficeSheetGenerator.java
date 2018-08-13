package org.apache.fineract.cn.datamigration.util;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.stream.IntStream;

public class BranchOfficeSheetGenerator {
  public static String identifier =RandomStringUtils.randomAlphanumeric(10);
  public static String parentIdentifier =OfficeSheetGenerator.identifier;
  public static String name=RandomStringUtils.randomAlphanumeric(9);
  public static String description=RandomStringUtils.randomAlphanumeric(10);
  public static String street=RandomStringUtils.randomAlphanumeric(10);
  public static String city=RandomStringUtils.randomAlphanumeric(10);
  public static String region=RandomStringUtils.randomAlphanumeric(16);
  public static String postalCode=RandomStringUtils.randomAlphanumeric(6);
  public static  String countryCode=RandomStringUtils.randomAlphabetic(2);
  public static String country=RandomStringUtils.randomAlphanumeric(10);

  public static void branchSheetDownload(){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Branch_Office");

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
    cell2.setCellValue("Parent Identifier");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Name");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Description");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5= rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Street");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6= rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("City");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7= rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Region");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8= rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Postal Code");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9= rowHeader.createCell(startColIndex+8);
    cell9.setCellValue("Country Code");
    cell9.setCellStyle(headerCellStyle);

    XSSFCell cell10= rowHeader.createCell(startColIndex+9);
    cell10.setCellValue("Country");
    cell10.setCellStyle(headerCellStyle);

    XSSFCell cell11 = firstRow.createCell(startColIndex+0);
    cell11.setCellValue(identifier);

    XSSFCell cell22 = firstRow.createCell(startColIndex+1);
    cell22.setCellValue(parentIdentifier);

    XSSFCell cell33 = firstRow.createCell(startColIndex+2);
    cell33.setCellValue(name);

    XSSFCell cell44 = firstRow.createCell(startColIndex+3);
    cell44.setCellValue(description);

    XSSFCell cell55= firstRow.createCell(startColIndex+4);
    cell55.setCellValue(street);

    XSSFCell cell66= firstRow.createCell(startColIndex+5);
    cell66.setCellValue(city);

    XSSFCell cell77= firstRow.createCell(startColIndex+6);
    cell77.setCellValue(region);

    XSSFCell cell88= firstRow.createCell(startColIndex+7);
    cell88.setCellValue(postalCode);

    XSSFCell cell99= firstRow.createCell(startColIndex+8);
    cell99.setCellValue(countryCode);

    XSSFCell cell111= firstRow.createCell(startColIndex+9);
    cell111.setCellValue(country);

    IntStream.range(0, 10).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

    try {
      // Retrieve the output stream
      FileOutputStream fos = new FileOutputStream(new File("/home/le/Desktop/Branch_Office.xlsx"));
      worksheet.getWorkbook().write(fos);
      fos.flush();
      fos.close();
    } catch (Exception e) {
      System.out.println("Unable to write report to the output stream");
    }

  }

}
