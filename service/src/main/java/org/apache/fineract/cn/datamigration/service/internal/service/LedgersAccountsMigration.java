package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.accounting.api.v1.client.LedgerManager;
import org.apache.fineract.cn.accounting.api.v1.domain.Account;
import org.apache.fineract.cn.accounting.api.v1.domain.LedgerPage;
import org.apache.fineract.cn.datamigration.service.ServiceConstants;
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
import java.util.HashSet;
import java.util.Set;
import java.util.stream.IntStream;

@Service
public class LedgersAccountsMigration {

  private final Logger logger;
  private final LedgerManager ledgerManager;

  @Autowired
  public LedgersAccountsMigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                                  final LedgerManager ledgerManager) {
    super();
    this.logger = logger;
    this.ledgerManager = ledgerManager;
  }
  public void accountSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Accounts");
    final LedgerPage currentLedgerPage = this.ledgerManager.fetchLedgers(true, null, null, null, null, null, null);
    int sizeOfLedger=currentLedgerPage.getLedgers().size();

    String [] ledgerIdentifier = new String[sizeOfLedger];
    for(int i=0;i<sizeOfLedger;i++){
      ledgerIdentifier[i]=currentLedgerPage.getLedgers().get(i).getIdentifier();
    }

    Datavalidator.validatorLedger(worksheet,"ASSET","LIABILITY","EQUITY","REVENUE","EXPENSE",0);
    Datavalidator.validatorString(worksheet,ledgerIdentifier,6);
    int startRowIndex = 0;
    int startColIndex = 0;

    Font font = worksheet.getWorkbook().createFont();
    XSSFCellStyle headerCellStyle = worksheet.getWorkbook().createCellStyle();

    headerCellStyle.setWrapText(true);
    headerCellStyle.setFont(font);
    XSSFRow rowHeader = worksheet.createRow((short) startRowIndex);
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


    IntStream.range(0, 7).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    response.setHeader("Content-Disposition", "inline; filename=Accounts.xlsx");
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

  public void accountSheetUpload(MultipartFile file){
    if (!file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
      throw new MultipartException("Only excel files accepted!");
    }
    try {
      XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
      Sheet firstSheet = workbook.getSheetAt(0);
      int rowCount = firstSheet.getLastRowNum() + 1;
      Row row;
      String type = null;
      String identifier = null;
      String name = null;
      String holders = null;
      String signatureAuthorities = null;
      Double balance = null;
      String ledger = null;

      for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
        row = firstSheet.getRow(rowIndex);
        if (row.getCell(0) == null) {
          type = null;
        } else {
          switch (row.getCell(0) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              type = row.getCell(0).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                type =   String.valueOf(((Double)row.getCell(0).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(1) == null) {
          identifier = null;
        } else {
          switch (row.getCell(1) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              identifier = row.getCell(1).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                identifier =   String.valueOf(((Double)row.getCell(1).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(2) == null) {
          name = null;
        } else {
          switch (row.getCell(2) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              name = row.getCell(2).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                name =  String.valueOf(((Double)row.getCell(2).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(3) == null) {
          holders = null;
        } else {
          switch (row.getCell(3) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              holders = row.getCell(3).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                holders =   String.valueOf(((Double)row.getCell(3).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(4) == null) {
          signatureAuthorities = null;
        } else {
          switch (row.getCell(4) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              signatureAuthorities = row.getCell(4).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                signatureAuthorities =   String.valueOf(((Double)row.getCell(4).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(5) == null) {
          balance = null;
        } else {
          switch (row.getCell(5) .getCellType()) {

            case Cell.CELL_TYPE_NUMERIC:
                balance = (Double)row.getCell(5).getNumericCellValue();
              break;
          }
        }

        if (row.getCell(6) == null) {
          ledger = null;
        } else {
          switch (row.getCell(6) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              ledger = row.getCell(6).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                ledger =   String.valueOf(((Double)row.getCell(6).getNumericCellValue()).intValue());
              break;
          }
        }

        Set<String> holder = new HashSet<>();
        holder.add(holders);

        Set<String> signatureAuthoritie = new HashSet<>();
        signatureAuthoritie.add(signatureAuthorities);

        Account account = new Account();
        account.setType(String.valueOf(type));
        account.setIdentifier(String.valueOf(identifier));
        account.setName(String.valueOf(name));
        account.setHolders(holder);
        account.setSignatureAuthorities(signatureAuthoritie);
        account.setBalance(balance);
        account.setLedger(String.valueOf(ledger));

        this.ledgerManager.createAccount(account);
      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

}
