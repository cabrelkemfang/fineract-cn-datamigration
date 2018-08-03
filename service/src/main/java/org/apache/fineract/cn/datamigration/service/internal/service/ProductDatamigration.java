package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.datamigration.service.connector.UserManagement;
import org.apache.fineract.cn.portfolio.api.v1.client.PortfolioManager;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.util.stream.IntStream;

public class ProductDatamigration {

  private final Logger logger;
  private  final PortfolioManager portfolioManager;
  private final UserManagement userManagement;

  @Autowired
  public ProductDatamigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                                final PortfolioManager portfolioManager,
                                final UserManagement userManagement) {
    super();
    this.logger = logger;
    this.portfolioManager = portfolioManager;
    this.userManagement = userManagement;
  }

  public static void productSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Customers");

    Datavalidator.validator(worksheet,"PERSON","BUSINESS",1);
    Datavalidator.validator(worksheet,"CURRENT_BALANCE","BEGINNING_BALANCE",9);
    Datavalidator.validator(worksheet,"TRUE","FALSE",12);
    Datavalidator.validatorState(worksheet,"PENDING","ACTIVE","LOCKED","CLOSED",24);

    Datavalidator.validator(worksheet,"BUSINESS","PRIVATE",20);
    Datavalidator.validatorType(worksheet,"EMAIL","PHONE","MOBILE",19);

    int startRowIndex = 0;
    int startColIndex = 0;

    Font font = worksheet.getWorkbook().createFont();
    XSSFCellStyle headerCellStyle = worksheet.getWorkbook().createCellStyle();

    headerCellStyle.setWrapText(true);
    headerCellStyle.setFont(font);
    XSSFRow rowHeader = worksheet.createRow((short) startRowIndex);
    rowHeader.setHeight((short) 500);


    XSSFCell cell1 = rowHeader.createCell(startColIndex+0);
    cell1.setCellValue("Identifier");
    cell1.setCellStyle(headerCellStyle);

    XSSFCell cell2 = rowHeader.createCell(startColIndex+1);
    cell2.setCellValue("Name");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Temporal Unit");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Term Range Minimum;");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5 = rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Term Range Maximum ");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6 = rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Balance Range Minimum ");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7 = rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Balance Range Maximum  ");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8 = rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Interest Range Minimum");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9= rowHeader.createCell(startColIndex+8);
    cell9.setCellValue("Interest Range Maximum");
    cell9.setCellStyle(headerCellStyle);

    XSSFCell cell10= rowHeader.createCell(startColIndex+9);
    cell10.setCellValue("Interest Basis ");
    cell10.setCellStyle(headerCellStyle);

    XSSFCell cell11= rowHeader.createCell(startColIndex+10);
    cell11.setCellValue("Pattern Package ");
    cell11.setCellStyle(headerCellStyle);

    XSSFCell cell12= rowHeader.createCell(startColIndex+11);
    cell12.setCellValue("Description ");
    cell12.setCellStyle(headerCellStyle);

    XSSFCell cell13= rowHeader.createCell(startColIndex+12);
    cell13.setCellValue("Enabled ");
    cell13.setCellStyle(headerCellStyle);

    XSSFCell cell14= rowHeader.createCell(startColIndex+13);
    cell14.setCellValue("Currency Code");
    cell14.setCellStyle(headerCellStyle);

    XSSFCell cell15= rowHeader.createCell(startColIndex+14);
    cell15.setCellValue("Minor Currency Unit Digits");
    cell15.setCellStyle(headerCellStyle);

    XSSFCell cell16= rowHeader.createCell(startColIndex+15);
    cell16.setCellValue("Parameters");
    cell16.setCellStyle(headerCellStyle);

    XSSFCell cell17= rowHeader.createCell(startColIndex+16);
    cell17.setCellValue("Postal Code");
    cell17.setCellStyle(headerCellStyle);

    XSSFCell cell18= rowHeader.createCell(startColIndex+17);
    cell18.setCellValue("Country Code");
    cell18.setCellStyle(headerCellStyle);

    XSSFCell cell19= rowHeader.createCell(startColIndex+18);
    cell19.setCellValue("Country");
    cell19.setCellStyle(headerCellStyle);

    //contactDetail
    XSSFCell cell20= rowHeader.createCell(startColIndex+19);
    cell20.setCellValue("Type");
    cell20.setCellStyle(headerCellStyle);

    XSSFCell cell21= rowHeader.createCell(startColIndex+20);
    cell21.setCellValue("Group");
    cell21.setCellStyle(headerCellStyle);

    XSSFCell cell22= rowHeader.createCell(startColIndex+21);
    cell22.setCellValue("Value");
    cell22.setCellStyle(headerCellStyle);

    XSSFCell cell23= rowHeader.createCell(startColIndex+22);
    cell23.setCellValue("Preference Level");
    cell23.setCellStyle(headerCellStyle);

    XSSFCell cell24= rowHeader.createCell(startColIndex+23);
    cell24.setCellValue("Validated");
    cell24.setCellStyle(headerCellStyle);

    XSSFCell cell25= rowHeader.createCell(startColIndex+24);
    cell25.setCellValue("Current State");
    cell25.setCellStyle(headerCellStyle);

    XSSFCell cell26= rowHeader.createCell(startColIndex+25);
    cell26.setCellValue("Application Date");
    cell26.setCellStyle(headerCellStyle);

    IntStream.range(0, 25).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    response.setHeader("Content-Disposition", "inline; filename=Customer.xlsx");
    response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    try {
      // Retrieve the output stream
      ServletOutputStream outputStream = response.getOutputStream();
      // Write to the output stream
      worksheet.getWorkbook().write(outputStream);
      // Flush the stream
      outputStream.flush();

      outputStream.close();
    } catch (Exception e) {
    }

  }

}
