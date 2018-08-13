package org.apache.fineract.cn.datamigration.util;


import org.apache.commons.lang3.RandomStringUtils;
import org.apache.fineract.cn.group.api.v1.domain.Cycle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.stream.IntStream;

public class GroupDefinitionGenerator {
  public static String identifier =RandomStringUtils.randomAlphanumeric(8);
  public static String description =RandomStringUtils.randomAlphanumeric(8);
  public static Integer minimalSize =10;
  public static Integer maximalSize =30;
  public static Integer numberOfMeetings =30;
  public static String frequency =Cycle.Frequency.WEEKLY.name();
  public static String adjustment =Cycle.Adjustment.NEXT_BUSINESS_DAY.name();

  public static void groupDefinitionSheetDownload( ){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Group_definition");

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
    cell2.setCellValue("Description");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Minimal Size ");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Maximal Size ");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5 = rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Number Of Meetings ");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6 = rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Frequency");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7 = rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Adjustment");
    cell7.setCellStyle(headerCellStyle);


    XSSFCell cellx1 = firstRow.createCell(startColIndex+0);
    cellx1.setCellValue(identifier);

    XSSFCell cellx2 = firstRow.createCell(startColIndex+1);
    cellx2.setCellValue(identifier);

    XSSFCell cellx3 = firstRow.createCell(startColIndex+2);
    cellx3.setCellValue(description);

    XSSFCell cellx4 = firstRow.createCell(startColIndex+3);
    cellx4.setCellValue(minimalSize);

    XSSFCell cellx5 = firstRow.createCell(startColIndex+4);
    cellx5.setCellValue(maximalSize);

    XSSFCell cellx6 = firstRow.createCell(startColIndex+5);
    cellx6.setCellValue(numberOfMeetings);

    XSSFCell cellx7 = firstRow.createCell(startColIndex+6);
    cellx7.setCellValue(frequency);

    XSSFCell cellx8 = firstRow.createCell(startColIndex+7);
    cellx8.setCellValue(adjustment);

    IntStream.range(0, 7).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

    try {
      FileOutputStream fos = new FileOutputStream(new File("/home/le/Desktop/Group_definition.xlsx"));
      worksheet.getWorkbook().write(fos);
      fos.flush();
      fos.close();
    } catch (Exception e) {
    }
  }

}
