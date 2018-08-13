package org.apache.fineract.cn.datamigration.util;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.stream.IntStream;

public class ProductInstanceSheetGenerator  {
  public static String customerIdentifier =RandomStringUtils.randomAlphanumeric(8);
  public static String productIdentifier =ProductDefinitionSheetGenerator.identifier;
  public static String accountIdentifier=RandomStringUtils.randomAlphanumeric(9);
  public static String alternativeAccountNumber=RandomStringUtils.randomAlphanumeric(10);
  public static String beneficiaries=RandomStringUtils.randomAlphanumeric(10);
  public static String state="PENDING";
  public static Double balance=0.00;

  public static void productInstanceSheetDownload( ){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Product_Instance");

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
    cell1.setCellValue("Customer Identifier");
    cell1.setCellStyle(headerCellStyle);

    XSSFCell cell2 = rowHeader.createCell(startColIndex+1);
    cell2.setCellValue("Product Identifier");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Account Identifier");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Alternative Account Number ");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5 = rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Beneficiaries ");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6 = rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("State ");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7 = rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Balance");
    cell7.setCellStyle(headerCellStyle);


    XSSFCell cellx1 = firstRow.createCell(startColIndex+0);
    cellx1.setCellValue(customerIdentifier);

    XSSFCell cellx2 = firstRow.createCell(startColIndex+1);
    cellx2.setCellValue(productIdentifier);

    XSSFCell cellx3 = firstRow.createCell(startColIndex+2);
    cellx3.setCellValue(accountIdentifier);

    XSSFCell cellx4 = firstRow.createCell(startColIndex+3);
    cellx4.setCellValue(alternativeAccountNumber);

    XSSFCell cellx5 = firstRow.createCell(startColIndex+4);
    cellx5.setCellValue(beneficiaries);

    XSSFCell cellx6 = firstRow.createCell(startColIndex+5);
    cellx6.setCellValue(state);

    XSSFCell cellx7 = firstRow.createCell(startColIndex+6);
    cellx7.setCellValue(balance);

    IntStream.range(0, 7).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));


    try {
      // Retrieve the output stream
      FileOutputStream fos = new FileOutputStream(new File("/home/le/Desktop/Product_Instance.xlsx"));
      worksheet.getWorkbook().write(fos);
      fos.flush();
      fos.close();
    } catch (Exception e) {
    }

  }
}
