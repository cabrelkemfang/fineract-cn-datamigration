package org.apache.fineract.cn.datamigration.util;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.time.Clock;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.stream.IntStream;

public class JournalEntrySheetGenerator {
  public static String transactionIdentifier =RandomStringUtils.randomAlphanumeric(8);
  public static String transactionDate=ZonedDateTime.now(Clock.systemUTC()).format(DateTimeFormatter.ISO_ZONED_DATE_TIME);
  public static String transactionType=RandomStringUtils.randomAlphanumeric(4);
  public static String clerk="clark";
  public static String note=RandomStringUtils.randomAlphanumeric(10);
  public static String accountNumberDebtor=RandomStringUtils.randomAlphanumeric(10);
  public static String amountDebtor=RandomStringUtils.randomAlphanumeric(10);
  public static String accountNumberCreditor=RandomStringUtils.randomAlphanumeric(6);
  public static  String amountCreditor=RandomStringUtils.randomAlphabetic(2);
  public static String state=RandomStringUtils.randomAlphanumeric(10);
  public static String message=RandomStringUtils.randomAlphanumeric(25);

  public void journalEntryDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("JournalEntry");

    int startRowIndex = 0;
    int startColIndex = 0;

    Font font = worksheet.getWorkbook().createFont();
    XSSFCellStyle headerCellStyle = worksheet.getWorkbook().createCellStyle();

    headerCellStyle.setWrapText(true);
    headerCellStyle.setFont(font);
    XSSFRow rowHeader = worksheet.createRow((short) startRowIndex);
    rowHeader.setHeight((short) 500);

    XSSFCell cell1 = rowHeader.createCell(startColIndex+0);
    cell1.setCellValue("Transaction Identifier");
    cell1.setCellStyle(headerCellStyle);

    XSSFCell cell2 = rowHeader.createCell(startColIndex+1);
    cell2.setCellValue("Transaction Date");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Transaction Type");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Clerk");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5= rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Note");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6= rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Debtor Account Number ");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7= rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Debtor Amount ");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8= rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Creditor Account Number ");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9= rowHeader.createCell(startColIndex+8);
    cell9.setCellValue("Creditor Amount ");
    cell9.setCellStyle(headerCellStyle);

    XSSFCell cell10= rowHeader.createCell(startColIndex+9);
    cell10.setCellValue("State");
    cell10.setCellStyle(headerCellStyle);

    XSSFCell cell11= rowHeader.createCell(startColIndex+10);
    cell10.setCellValue("Message");
    cell10.setCellStyle(headerCellStyle);

    IntStream.range(0, 10).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

    try {
      // Retrieve the output stream
      FileOutputStream fos = new FileOutputStream(new File("/home/le/Desktop/JournalEntry.xlsx"));
      worksheet.getWorkbook().write(fos);
      fos.flush();
      fos.close();
    } catch (Exception e) {
      System.out.println("Unable to write report to the output stream");
    }

  }
}
