package org.apache.fineract.cn.datamigration.util;


import org.apache.commons.lang3.RandomStringUtils;
import org.apache.fineract.cn.accounting.api.v1.domain.AccountType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.stream.IntStream;

public class SubLedgerSheetGenerator {
  public static String type =AccountType.ASSET.name();
  public static String identifier =RandomStringUtils.randomAlphanumeric(8);
  public static String name=RandomStringUtils.randomAlphanumeric(9);
  public static String description=RandomStringUtils.randomAlphanumeric(10);
  public static Boolean showAccountsInChart=Boolean.FALSE;

  public static void subLedgerSheetDownload(){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("SubLedgers");

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
    cell1.setCellValue("Parent Ledger Identifier");
    cell1.setCellStyle(headerCellStyle);

    XSSFCell cell2 = rowHeader.createCell(startColIndex+1);
    cell2.setCellValue("Type");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Identifier");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Name");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5 = rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Description");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6= rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Show Accounts In Chart");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell11 = firstRow.createCell(startColIndex+0);
    cell11.setCellValue(LedgerSheetGenerator.identifier);

    XSSFCell cell22 = firstRow.createCell(startColIndex+1);
    cell22.setCellValue(type);

    XSSFCell cell33 = firstRow.createCell(startColIndex+2);
    cell33.setCellValue(identifier);

    XSSFCell cell44 = firstRow.createCell(startColIndex+3);
    cell44.setCellValue(name);

    XSSFCell cell55= firstRow.createCell(startColIndex+4);
    cell55.setCellValue(description);

    XSSFCell cell66= firstRow.createCell(startColIndex+5);
    cell66.setCellValue(showAccountsInChart);

    IntStream.range(0, 6).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

    try {
      // Retrieve the output stream
      FileOutputStream fos = new FileOutputStream(new File("/home/le/Desktop/SubLedgers.xlsx"));
      worksheet.getWorkbook().write(fos);
      fos.flush();
      fos.close();
    } catch (Exception e) {
      System.out.println("Unable to write report to the output stream");
    }

  }

}
