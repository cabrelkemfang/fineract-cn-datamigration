package org.apache.fineract.cn.datamigration.util;


import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.stream.IntStream;

public class GroupSheetGenerator {

  public static String identifier =RandomStringUtils.randomAlphanumeric(15);
  public static String groupDefinitionIdentifier =GroupDefinitionGenerator.identifier;
  public static String name =RandomStringUtils.randomAlphanumeric(15);
  public static String leaders =RandomStringUtils.randomAlphanumeric(15);
  public static String members =RandomStringUtils.randomAlphanumeric(15);
  public static String office =RandomStringUtils.randomAlphanumeric(15);
  public static String assignedEmployee =RandomStringUtils.randomAlphanumeric(15);
  public static String weekdays =RandomStringUtils.randomAlphanumeric(15);
  public static String status =RandomStringUtils.randomAlphanumeric(15);
  public static String street =RandomStringUtils.randomAlphanumeric(15);
  public static String city =RandomStringUtils.randomAlphanumeric(15);
  public static String region =RandomStringUtils.randomAlphanumeric(15);
  public static String postalCode =RandomStringUtils.randomAlphanumeric(15);
  public static String countryCode =RandomStringUtils.randomAlphanumeric(15);
  public static String country =RandomStringUtils.randomAlphanumeric(15);

  public static void groupSheetDownload() {
      XSSFWorkbook workbook = new XSSFWorkbook();
      XSSFSheet worksheet = workbook.createSheet("Group");

      int startRowIndex = 0;
      int startColIndex = 0;

      Font font = worksheet.getWorkbook().createFont();
      XSSFCellStyle headerCellStyle = worksheet.getWorkbook().createCellStyle();


      headerCellStyle.setWrapText(true);
      headerCellStyle.setFont(font);
      XSSFRow rowHeader = worksheet.createRow((short) startRowIndex);
      XSSFRow firstRow = worksheet.createRow(1);
      rowHeader.setHeight((short) 500);

      XSSFCell cell1 = rowHeader.createCell(startColIndex + 0);
      cell1.setCellValue("Identifier");
      cell1.setCellStyle(headerCellStyle);

      XSSFCell cell2 = rowHeader.createCell(startColIndex + 1);
      cell2.setCellValue("Group Definition Identifier");
      cell2.setCellStyle(headerCellStyle);

      XSSFCell cell3 = rowHeader.createCell(startColIndex + 2);
      cell3.setCellValue("Name");
      cell3.setCellStyle(headerCellStyle);

      XSSFCell cell4 = rowHeader.createCell(startColIndex + 3);
      cell4.setCellValue("Leaders");
      cell4.setCellStyle(headerCellStyle);

      XSSFCell cell5 = rowHeader.createCell(startColIndex + 4);
      cell5.setCellValue("Members");
      cell5.setCellStyle(headerCellStyle);

      XSSFCell cell6 = rowHeader.createCell(startColIndex + 5);
      cell6.setCellValue("Office");
      cell6.setCellStyle(headerCellStyle);

      XSSFCell cell7 = rowHeader.createCell(startColIndex + 6);
      cell7.setCellValue("Assigned Employee");
      cell7.setCellStyle(headerCellStyle);

      XSSFCell cell8 = rowHeader.createCell(startColIndex + 7);
      cell8.setCellValue("Weekday");
      cell8.setCellStyle(headerCellStyle);

      XSSFCell cell9 = rowHeader.createCell(startColIndex + 8);
      cell9.setCellValue("Status");
      cell9.setCellStyle(headerCellStyle);

      XSSFCell cell10 = rowHeader.createCell(startColIndex + 9);
      cell10.setCellValue("Street");
      cell10.setCellStyle(headerCellStyle);

      XSSFCell cell11 = rowHeader.createCell(startColIndex + 10);
      cell11.setCellValue("City");
      cell11.setCellStyle(headerCellStyle);

      XSSFCell cell12 = rowHeader.createCell(startColIndex + 11);
      cell12.setCellValue("Region");
      cell12.setCellStyle(headerCellStyle);

      XSSFCell cell13 = rowHeader.createCell(startColIndex + 12);
      cell13.setCellValue("Postal Code");
      cell13.setCellStyle(headerCellStyle);

      XSSFCell cell14 = rowHeader.createCell(startColIndex + 13);
      cell14.setCellValue("Country Code");
      cell14.setCellStyle(headerCellStyle);

      XSSFCell cell15 = rowHeader.createCell(startColIndex + 14);
      cell15.setCellValue("Country ");
      cell15.setCellStyle(headerCellStyle);


      XSSFCell cellx1 = firstRow.createCell(startColIndex + 0);
      cellx1.setCellValue(identifier);

      XSSFCell cellx2 = firstRow.createCell(startColIndex + 1);
      cellx2.setCellValue(groupDefinitionIdentifier);

      XSSFCell cellx3 = firstRow.createCell(startColIndex + 2);
      cellx3.setCellValue(name);

      XSSFCell cellx4 = firstRow.createCell(startColIndex + 3);
      cellx4.setCellValue(leaders);

      XSSFCell cellx5 = firstRow.createCell(startColIndex + 4);
      cellx5.setCellValue(members);

      XSSFCell cellx6 = firstRow.createCell(startColIndex + 5);
      cellx6.setCellValue(office);

      XSSFCell cellx7 = firstRow.createCell(startColIndex + 6);
      cellx7.setCellValue(assignedEmployee);

      XSSFCell cellx8 = firstRow.createCell(startColIndex + 7);
      cellx8.setCellValue(weekdays);

      XSSFCell cellx9 = firstRow.createCell(startColIndex + 8);
      cellx9.setCellValue(status);

      XSSFCell cellx10 = firstRow.createCell(startColIndex + 9);
      cellx10.setCellValue(street);

      XSSFCell cellx11 = firstRow.createCell(startColIndex + 10);
      cellx11.setCellValue(city);

      XSSFCell cellx12 = firstRow.createCell(startColIndex + 11);
      cellx12.setCellValue(region);

      XSSFCell cellx13 = firstRow.createCell(startColIndex + 12);
      cellx13.setCellValue(postalCode);

      XSSFCell cellx14 = firstRow.createCell(startColIndex + 13);
      cellx14.setCellValue(countryCode);

      XSSFCell cellx15 = firstRow.createCell(startColIndex + 14);
      cellx15.setCellValue(country);

      IntStream.range(0, 15).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

      try {
        // Retrieve the output stream
        FileOutputStream fos = new FileOutputStream(new File("/home/le/Desktop/Group.xlsx"));
        worksheet.getWorkbook().write(fos);
        fos.flush();
        fos.close();
      } catch (Exception e) {
        System.out.println("Unable to write report to the output stream");
      }

    }

}
