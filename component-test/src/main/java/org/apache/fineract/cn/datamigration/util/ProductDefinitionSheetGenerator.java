package org.apache.fineract.cn.datamigration.util;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.fineract.cn.deposit.api.v1.domain.InterestPayable;
import org.apache.fineract.cn.deposit.api.v1.domain.TimeUnit;
import org.apache.fineract.cn.deposit.api.v1.domain.Type;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.stream.IntStream;

public class ProductDefinitionSheetGenerator {
  public static String identifier =RandomStringUtils.randomAlphanumeric(8);
  public static String type =Type.SAVINGS.name();
  public static String name =RandomStringUtils.randomAlphanumeric(10);
  public static String description =RandomStringUtils.randomAlphanumeric(15);
  public static String currencyCode ="USD ";
  public static String currencyName ="US Dollar";
  public static String currencySign ="$";
  public static Integer currencyScale =3;
  public static Double minimumBalance =50.00;
  public static String equityLedgerIdentifier ="91xx";
  public static String cashAccountIdentifier ="76xx";
  public static String expenseAccountIdentifier ="38xx";
  public static String accrueAccountIdentifier ="82xx";
  public static Double interest =1.25D;
  public static Integer termPeriod =12;
  public static String termTimeUnit =TimeUnit.MONTH.name();
  public static String termInterestPayable =InterestPayable.MATURITY.name();
  public static String chargeActionIdentifier ="Open";
  public static String chargeIncomeAccountIdentifier ="10123";
  public static String chargeName ="Closing Account Fee";
  public static String chargeDescription ="chargeDescription";
  public static Boolean chargeProportional =Boolean.FALSE;
  public static Double chargeAmount =2.00D;
  public static Boolean flexible =Boolean.FALSE;
  public static Boolean active =Boolean.FALSE;


  public static void productDefinitionSheetDownload( ){
    XSSFWorkbook workbook = new XSSFWorkbook();

    XSSFSheet worksheet = workbook.createSheet("Product_Definitions");

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
    cell4.setCellValue("Description");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5 = rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Currency Code ");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6 = rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Currency Name");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7 = rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Currency Sign");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8 = rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Currency Scale ");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9= rowHeader.createCell(startColIndex+8);
    cell9.setCellValue(" Minimum Balance ");
    cell9.setCellStyle(headerCellStyle);

    XSSFCell cell10= rowHeader.createCell(startColIndex+9);
    cell10.setCellValue(" Equity Ledger Identifier");
    cell10.setCellStyle(headerCellStyle);

    XSSFCell cell11= rowHeader.createCell(startColIndex+10);
    cell11.setCellValue("Cash Account Identifier  ");
    cell11.setCellStyle(headerCellStyle);

    XSSFCell cell12= rowHeader.createCell(startColIndex+11);
    cell12.setCellValue("Expense Account Identifier");
    cell12.setCellStyle(headerCellStyle);

    XSSFCell cell13= rowHeader.createCell(startColIndex+12);
    cell13.setCellValue("Accrue Account Identifier ");
    cell13.setCellStyle(headerCellStyle);

    XSSFCell cell14= rowHeader.createCell(startColIndex+13);
    cell14.setCellValue("Interest ");
    cell14.setCellStyle(headerCellStyle);

    XSSFCell cell15= rowHeader.createCell(startColIndex+14);
    cell15.setCellValue(" Term Period  ");
    cell15.setCellStyle(headerCellStyle);

    XSSFCell cell16= rowHeader.createCell(startColIndex+15);
    cell16.setCellValue(" Term Time Unit ");
    cell16.setCellStyle(headerCellStyle);

    XSSFCell cell17= rowHeader.createCell(startColIndex+16);
    cell17.setCellValue(" Interest Payable ");
    cell17.setCellStyle(headerCellStyle);

    XSSFCell cell18= rowHeader.createCell(startColIndex+17);
    cell18.setCellValue(" Action Identifier ");
    cell18.setCellStyle(headerCellStyle);

    XSSFCell cell19= rowHeader.createCell(startColIndex+18);
    cell19.setCellValue(" Income Account Identifier ");
    cell19.setCellStyle(headerCellStyle);

    XSSFCell cell20= rowHeader.createCell(startColIndex+19);
    cell20.setCellValue(" Charge Name ");
    cell20.setCellStyle(headerCellStyle);

    XSSFCell cell21= rowHeader.createCell(startColIndex+20);
    cell21.setCellValue(" Charge Description ");
    cell21.setCellStyle(headerCellStyle);

    XSSFCell cell22= rowHeader.createCell(startColIndex+21);
    cell22.setCellValue(" Charge Proportional ");
    cell22.setCellStyle(headerCellStyle);

    XSSFCell cell23= rowHeader.createCell(startColIndex+22);
    cell23.setCellValue(" Charge Amount ");
    cell23.setCellStyle(headerCellStyle);

    XSSFCell cell24= rowHeader.createCell(startColIndex+23);
    cell24.setCellValue("Flexible");
    cell24.setCellStyle(headerCellStyle);

    XSSFCell cell25= rowHeader.createCell(startColIndex+24);
    cell25.setCellValue(" Active");
    cell25.setCellStyle(headerCellStyle);



    XSSFCell cellx1 = firstRow.createCell(startColIndex+0);
    cellx1.setCellValue(type);

    XSSFCell cellx2 = firstRow.createCell(startColIndex+1);
    cellx2.setCellValue(identifier);

    XSSFCell cellx3 = firstRow.createCell(startColIndex+2);
    cellx3.setCellValue(name);

    XSSFCell cellx4 = firstRow.createCell(startColIndex+3);
    cellx4.setCellValue(description);

    XSSFCell cellx5 = firstRow.createCell(startColIndex+4);
    cellx5.setCellValue(currencyCode);

    XSSFCell cellx6 = firstRow.createCell(startColIndex+5);
    cellx6.setCellValue(currencyName);

    XSSFCell cellx7 = firstRow.createCell(startColIndex+6);
    cellx7.setCellValue(currencySign);

    XSSFCell cellx8 = firstRow.createCell(startColIndex+7);
    cellx8.setCellValue(currencyScale);

    XSSFCell cellx9= firstRow.createCell(startColIndex+8);
    cellx9.setCellValue(minimumBalance);

    XSSFCell cellx10= firstRow.createCell(startColIndex+9);
    cellx10.setCellValue(equityLedgerIdentifier);

    XSSFCell cellx11= firstRow.createCell(startColIndex+10);
    cellx11.setCellValue(cashAccountIdentifier);

    XSSFCell cellx12= firstRow.createCell(startColIndex+11);
    cellx12.setCellValue(expenseAccountIdentifier);

    XSSFCell cellx13= firstRow.createCell(startColIndex+12);
    cellx13.setCellValue(accrueAccountIdentifier);

    XSSFCell cellx14= firstRow.createCell(startColIndex+13);
    cellx14.setCellValue(interest);

    XSSFCell cellx15= firstRow.createCell(startColIndex+14);
    cellx15.setCellValue(termPeriod);

    XSSFCell cellx16= firstRow.createCell(startColIndex+15);
    cellx16.setCellValue(termTimeUnit);

    XSSFCell cellx17= firstRow.createCell(startColIndex+16);
    cellx17.setCellValue(termInterestPayable);

    XSSFCell cellx18= firstRow.createCell(startColIndex+17);
    cellx18.setCellValue(chargeActionIdentifier);

    XSSFCell cellx19= firstRow.createCell(startColIndex+18);
    cellx19.setCellValue(chargeIncomeAccountIdentifier);

    XSSFCell cellx20= firstRow.createCell(startColIndex+19);
    cellx20.setCellValue(chargeName);

    XSSFCell cellx21= firstRow.createCell(startColIndex+20);
    cellx21.setCellValue(chargeDescription);

    XSSFCell cellx22= firstRow.createCell(startColIndex+21);
    cellx22.setCellValue(chargeProportional);

    XSSFCell cellx23= firstRow.createCell(startColIndex+22);
    cellx23.setCellValue(chargeAmount);

    XSSFCell cellx24= firstRow.createCell(startColIndex+23);
    cellx24.setCellValue(flexible);

    XSSFCell cellx25= firstRow.createCell(startColIndex+24);
    cellx25.setCellValue(active);

    IntStream.range(0, 25).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

    try {
      // Retrieve the output stream
      FileOutputStream fos = new FileOutputStream(new File("/home/le/Desktop/Product_Definitions.xlsx"));
      worksheet.getWorkbook().write(fos);
      fos.flush();
      fos.close();
    } catch (Exception e) {
    }

  }
}
