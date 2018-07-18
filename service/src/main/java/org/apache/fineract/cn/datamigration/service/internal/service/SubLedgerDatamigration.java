package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.accounting.api.v1.client.LedgerManager;
import org.apache.fineract.cn.accounting.api.v1.domain.Ledger;
import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.datamigration.service.connector.UserManagement;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.slf4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartException;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.stream.IntStream;

@Service
public class SubLedgerDatamigration {
  private final Logger logger;
  private final LedgerManager ledgerManager;
  private final UserManagement userManagement;

  @Autowired
  public SubLedgerDatamigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                             final LedgerManager ledgerManager,
                             final UserManagement userManagement) {
    super();
    this.logger = logger;
    this.ledgerManager = ledgerManager;
    this.userManagement = userManagement;
  }

  public void subLedgerSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("SubLedgers");

    Datavalidator.validatorLedger(worksheet,"ASSET","LIABILITY","EQUITY","REVENUE","EXPENSE",0);

    int startRowIndex = 0;
    int startColIndex = 0;

    Font font = worksheet.getWorkbook().createFont();
    XSSFCellStyle headerCellStyle = worksheet.getWorkbook().createCellStyle();

    headerCellStyle.setWrapText(true);
    headerCellStyle.setFont(font);
    XSSFRow rowHeader = worksheet.createRow((short) startRowIndex);
    rowHeader.setHeight((short) 500);

    XSSFCell cell1 = rowHeader.createCell(startColIndex+0);
    cell1.setCellValue("Ledger Identifier");
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


    IntStream.range(0, 4).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    response.setHeader("Content-Disposition", "inline; filename=SubLedgers.xlsx");
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
      System.out.println("Unable to write report to the output stream");
    }

  }

  public void subLedgerSheetUpload(MultipartFile file){
    if (!file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
      throw new MultipartException("Only excel files accepted!");
    }
    try {
      XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
      Sheet firstSheet = workbook.getSheetAt(0);
      int rowCount = firstSheet.getLastRowNum() + 1;
      Row row;
      String ledgerIdentifier=null;
      String type = null;
      String identifier = null;
      String name = null;
      String description = null;
      Boolean showAccountsInChart = false;

      SimpleDateFormat date=new SimpleDateFormat("yyyy-MM-dd");

      for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
        row = firstSheet.getRow(rowIndex);

        if (row.getCell(0) == null) {
          ledgerIdentifier = null;
        } else {
          switch (row.getCell(0) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              ledgerIdentifier = row.getCell(0).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(0))) {

                ledgerIdentifier =  date.format(row.getCell(0).getDateCellValue());
              } else {
                ledgerIdentifier =  String.valueOf(row.getCell(0).getNumericCellValue());
              }
              break;
          }
        }


        if (row.getCell(1) == null) {
          type = null;
        } else {
          switch (row.getCell(1) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              type = row.getCell(1).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(1))) {

                type =  date.format(row.getCell(1).getDateCellValue());
              } else {
                type =  String.valueOf(row.getCell(1).getNumericCellValue());
              }
              break;
          }
        }
        if (row.getCell(2) == null) {
          identifier = null;
        } else {
          switch (row.getCell(2) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              identifier = row.getCell(2).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(2))) {

                identifier =   date.format(row.getCell(2).getDateCellValue());
              } else {
                identifier =   String.valueOf(((Double)row.getCell(2).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(3) == null) {
          name = null;
        } else {
          switch (row.getCell(3) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              name = row.getCell(3).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(3))) {

                name =  date.format(row.getCell(3).getDateCellValue());
              } else {
                name =  String.valueOf(((Double)row.getCell(3).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(4) == null) {
          description = null;
        } else {
          switch (row.getCell(4) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              description = row.getCell(4).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(4))) {

                description =   date.format(row.getCell(4).getDateCellValue());
              } else {
                description =   String.valueOf(((Double)row.getCell(4).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(5) == null) {
          showAccountsInChart = false;
        } else {
          switch (row.getCell(5) .getCellType()) {

            case Cell.CELL_TYPE_NUMERIC:

              if(((Double)row.getCell(5).getNumericCellValue()).intValue()==0){
                showAccountsInChart = Boolean.parseBoolean("false");
              }else{
                showAccountsInChart = Boolean.parseBoolean("true");
              }
              break;
          }
        }

        Ledger ledger = new Ledger();

        ledger.setType(String.valueOf(type));
        ledger.setIdentifier(String.valueOf(identifier));
        ledger.setName(String.valueOf(name));
        ledger.setDescription(String.valueOf(description));
        ledger.setShowAccountsInChart(showAccountsInChart);

        this.userManagement.authenticate();
        this.ledgerManager.addSubLedger(ledgerIdentifier,ledger);
      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

}
