package org.apache.fineract.cn.datamigration.util;


import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.stream.IntStream;

public class TellersSheetGenerator {

  public static String officeIdentifier =OfficeSheetGenerator.identifier;
  public static String code =RandomStringUtils.randomAlphanumeric(8);
  public static String password =RandomStringUtils.randomAlphanumeric(32);
  public static BigDecimal cashdrawLimit =BigDecimal.valueOf(10000L);
  public static String tellerAccountIdentifier =RandomStringUtils.randomAlphanumeric(8);
  public static String vaultAccountIdentifier =RandomStringUtils.randomAlphanumeric(8);
  public static String chequesReceivableAccount =RandomStringUtils.randomAlphanumeric(8);
  public static String cashOverShortAccount =RandomStringUtils.randomAlphanumeric(8);
  public static Boolean denominationRequired =Boolean.FALSE;
  public static String assignedEmployee =RandomStringUtils.randomAlphanumeric(8);
  public static String state ="CLOSED";

  public static void tellerSheetDownload( ){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Tellers");

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
    cell1.setCellValue("Office Identifier");
    cell1.setCellStyle(headerCellStyle);

    XSSFCell cell2 = rowHeader.createCell(startColIndex+1);
    cell2.setCellValue("code");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("password");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Cashdraw Limit");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5 = rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Teller Account Identifier");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6 = rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Vault Account Identifierault ");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7 = rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Cheques Receivable Account ");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8 = rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Cash Over Short Account ");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9 = rowHeader.createCell(startColIndex+8);
    cell9.setCellValue("Denomination Required ");
    cell9.setCellStyle(headerCellStyle);

    XSSFCell cell10= rowHeader.createCell(startColIndex+9);
    cell10.setCellValue("Assigned Employee");
    cell10.setCellStyle(headerCellStyle);

    XSSFCell cell11= rowHeader.createCell(startColIndex+10);
    cell11.setCellValue("State");
    cell11.setCellStyle(headerCellStyle);



    XSSFCell cellx1 = firstRow.createCell(startColIndex+0);
    cellx1.setCellValue(officeIdentifier);

    XSSFCell cellx2 = firstRow.createCell(startColIndex+1);
    cellx2.setCellValue(code);

    XSSFCell cellx3 = firstRow.createCell(startColIndex+2);
    cellx3.setCellValue(password);

    XSSFCell cellx4 = firstRow.createCell(startColIndex+3);
    cellx4.setCellValue(String.valueOf(cashdrawLimit));

    XSSFCell cellx5 = firstRow.createCell(startColIndex+4);
    cellx5.setCellValue(tellerAccountIdentifier);

    XSSFCell cellx6 = firstRow.createCell(startColIndex+5);
    cellx6.setCellValue(vaultAccountIdentifier);

    XSSFCell cellx7 = firstRow.createCell(startColIndex+6);
    cellx7.setCellValue(chequesReceivableAccount);

    XSSFCell cellx8 = firstRow.createCell(startColIndex+7);
    cellx8.setCellValue(cashOverShortAccount);

    XSSFCell cellx9= firstRow.createCell(startColIndex+8);
    cellx9.setCellValue(denominationRequired);

    XSSFCell cellx10= firstRow.createCell(startColIndex+9);
    cellx10.setCellValue(assignedEmployee);

    XSSFCell cellx11= firstRow.createCell(startColIndex+10);
    cellx11.setCellValue(state);

    IntStream.range(0, 11).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    try {
      // Retrieve the output stream
      FileOutputStream fos = new FileOutputStream(new File("/home/le/Desktop/Tellers.xlsx"));
      worksheet.getWorkbook().write(fos);
      fos.flush();
      fos.close();
    } catch (Exception e) {
    }

  }
}
