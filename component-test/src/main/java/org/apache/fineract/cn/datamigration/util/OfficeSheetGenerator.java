package org.apache.fineract.cn.datamigration.util;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.util.stream.IntStream;

public class OfficeSheetGenerator {


   private OfficeSheetGenerator(){
     super();
   }
  public static String identifier =RandomStringUtils.randomAlphanumeric(10);
  public static String name=RandomStringUtils.randomAlphanumeric(9);
  public static String description=RandomStringUtils.randomAlphanumeric(10);
  public static String street=RandomStringUtils.randomAlphanumeric(10);
  public static String city=RandomStringUtils.randomAlphanumeric(10);
  public static String region=RandomStringUtils.randomAlphanumeric(10);
  public static String postalCode=RandomStringUtils.randomAlphanumeric(6);
  public static  String countryCode=RandomStringUtils.randomAlphabetic(2);
  public static String country=RandomStringUtils.randomAlphanumeric(10);


  //generating office excel sheet with some data
  public static void CreateofficeSheet(){
    HttpServletResponse response=null;
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Offices");

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
    cell2.setCellValue("Name ");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Description");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Street");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5= rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("City");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6= rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Region");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7= rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Postal Code");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8= rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Country Code");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9= rowHeader.createCell(startColIndex+8);
    cell9.setCellValue("Country");
    cell9.setCellStyle(headerCellStyle);


//gerenating data to populate the sheet
    XSSFCell cell11 = firstRow.createCell(startColIndex+0);
    cell11.setCellValue(identifier);

    XSSFCell cell22 = firstRow.createCell(startColIndex+1);
    cell22.setCellValue(name);

    XSSFCell cell33 = firstRow.createCell(startColIndex+2);
    cell33.setCellValue(description);

    XSSFCell cell44 = firstRow.createCell(startColIndex+3);
    cell44.setCellValue(street);

    XSSFCell cell55= firstRow.createCell(startColIndex+4);
    cell55.setCellValue(city);

    XSSFCell cell66= firstRow.createCell(startColIndex+5);
    cell66.setCellValue(region);

    XSSFCell cell77= firstRow.createCell(startColIndex+6);
    cell77.setCellValue(postalCode);

    XSSFCell cell88= firstRow.createCell(startColIndex+7);
    cell88.setCellValue(countryCode);

    XSSFCell cell99= firstRow.createCell(startColIndex+8);
    cell99.setCellValue(country);

    IntStream.range(0, 9).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    try {
      // Retrieve the output stream
      FileOutputStream fos = new FileOutputStream(new File("/home/le/Desktop/office.xlsx"));
      worksheet.getWorkbook().write(fos);
      fos.flush();
      fos.close();

    } catch (Exception e) {
      System.out.println("Unable to write report to the output stream");
    }

  }

}
