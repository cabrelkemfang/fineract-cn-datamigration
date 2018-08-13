package org.apache.fineract.cn.datamigration.util;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.stream.IntStream;

public class ActionSheetGenerator {
  public static String identifier =RandomStringUtils.randomAlphanumeric(10);
  public static String name =RandomStringUtils.randomAlphanumeric(9);
  public static String description=RandomStringUtils.randomAlphanumeric(15);
  public static String transactionType="ACCO";


  public static void actionSheetDownload( ){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Action");

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
    cell1.setCellValue(" Identifier");
    cell1.setCellStyle(headerCellStyle);

    XSSFCell cell2 = rowHeader.createCell(startColIndex+1);
    cell2.setCellValue("Name");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Description ");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Transaction Type");
    cell4.setCellStyle(headerCellStyle);


    XSSFCell cellx1 = firstRow.createCell(startColIndex+0);
    cellx1.setCellValue(identifier);

    XSSFCell cellx2 = firstRow.createCell(startColIndex+1);
    cellx2.setCellValue(name);

    XSSFCell cellx3 = firstRow.createCell(startColIndex+2);
    cellx3.setCellValue(description);

    XSSFCell cellx4 = firstRow.createCell(startColIndex+3);
    cellx4.setCellValue(transactionType);


    IntStream.range(0, 4).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

    try {
      // Retrieve the output stream
      FileOutputStream fos = new FileOutputStream(new File("/home/le/Desktop/Action.xlsx"));
      worksheet.getWorkbook().write(fos);
      fos.flush();
      fos.close();
    } catch (Exception e) {
    }
  }

}
