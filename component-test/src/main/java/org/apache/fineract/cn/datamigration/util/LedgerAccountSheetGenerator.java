package org.apache.fineract.cn.datamigration.util;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.fineract.cn.accounting.api.v1.domain.AccountType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.stream.IntStream;

public class LedgerAccountSheetGenerator {

  public static String type =AccountType.ASSET.name();
  public static String identifier=RandomStringUtils.randomAlphanumeric(9);
  public static String name=RandomStringUtils.randomAlphanumeric(10);
  public static String holders=RandomStringUtils.randomAlphanumeric(10);
  public static String signatureAuthorities=RandomStringUtils.randomAlphanumeric(10);
  public static Double balance=0.00;
  public static String ledger=LedgerSheetGenerator.identifier;


  public static void accountSheetDownload( ){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Accounts");

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
    cell1.setCellValue("Type");
    cell1.setCellStyle(headerCellStyle);

    XSSFCell cell2 = rowHeader.createCell(startColIndex+1);
    cell2.setCellValue("Identifier");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Name");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Holders");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5= rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Signature Authorities");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6= rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Balance");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7= rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Ledger");
    cell7.setCellStyle(headerCellStyle);


    XSSFCell cell11 = firstRow.createCell(startColIndex+0);
    cell11.setCellValue(type);

    XSSFCell cell22 = firstRow.createCell(startColIndex+1);
    cell22.setCellValue(identifier);

    XSSFCell cell33 = firstRow.createCell(startColIndex+2);
    cell33.setCellValue(name);

    XSSFCell cell44 = firstRow.createCell(startColIndex+3);
    cell44.setCellValue(holders);

    XSSFCell cell55= firstRow.createCell(startColIndex+4);
    cell55.setCellValue(signatureAuthorities);

    XSSFCell cell66= firstRow.createCell(startColIndex+5);
    cell66.setCellValue(balance);

    XSSFCell cell77= firstRow.createCell(startColIndex+6);
    cell77.setCellValue(ledger);

    IntStream.range(0, 7).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

    try {
      // Retrieve the output stream
      FileOutputStream fos = new FileOutputStream(new File("/home/le/Desktop/Accounts.xlsx"));
      worksheet.getWorkbook().write(fos);
      fos.flush();
      fos.close();
    } catch (Exception e) {
      System.out.println("Unable to write report to the output stream");
    }

  }
}
