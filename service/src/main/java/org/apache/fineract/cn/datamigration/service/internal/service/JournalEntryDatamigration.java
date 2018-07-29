package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.accounting.api.v1.client.LedgerManager;
import org.apache.fineract.cn.accounting.api.v1.domain.Creditor;
import org.apache.fineract.cn.accounting.api.v1.domain.Debtor;
import org.apache.fineract.cn.accounting.api.v1.domain.JournalEntry;
import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.datamigration.service.connector.UserManagement;
import org.apache.fineract.cn.office.api.v1.domain.Address;
import org.apache.fineract.cn.office.api.v1.domain.Office;
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
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.IntStream;

@Service
public class JournalEntryDatamigration {
  private final Logger logger;
  private final LedgerManager ledgerManager;
  private final UserManagement userManagement;

  @Autowired
  public JournalEntryDatamigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                             final LedgerManager ledgerManager,
                             final UserManagement userManagement) {
    super();
    this.logger = logger;
    this.ledgerManager = ledgerManager;
    this.userManagement = userManagement;
  }

  public void journalEntryDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("JournalEntry");
    Datavalidator.validator(worksheet,"PENDING","PROCESSED",9);

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
    cell6.setCellValue("Account Number (Debtor)");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7= rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Amount (Debtor)");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8= rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Account Number (Creditor)");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9= rowHeader.createCell(startColIndex+8);
    cell9.setCellValue("Amount (Creditor)");
    cell9.setCellStyle(headerCellStyle);

    XSSFCell cell10= rowHeader.createCell(startColIndex+9);
    cell10.setCellValue("State");
    cell10.setCellStyle(headerCellStyle);

    XSSFCell cell11= rowHeader.createCell(startColIndex+10);
    cell10.setCellValue("Message");
    cell10.setCellStyle(headerCellStyle);

    IntStream.range(0, 10).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    response.setHeader("Content-Disposition", "inline; filename=JournalENtry.xlsx");
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

  public void journalEntrySheetUpload(MultipartFile file){
    if (!file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
      throw new MultipartException("Only excel files accepted!");
    }
    try {
      XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
      Sheet firstSheet = workbook.getSheetAt(0);
      int rowCount = firstSheet.getLastRowNum() + 1;
      Row row;

      String transactionIdentifier =null ;
      String transactionDate  =null ;
      String transactionType  =null;
      String clerk  =null ;
      String note   =null;
      String accountNumberDebtor   =null;
      String amountDebtor   =null;
      String accountNumberCreditor  =null;
      String amountCreditor  =null;
      String state   =null;
      String message   =null;

      for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
        row = firstSheet.getRow(rowIndex);
        if (row.getCell(0) == null) {
          transactionIdentifier = null;
        } else {
          switch (row.getCell(0) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              transactionIdentifier = row.getCell(0).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(0))) {

                transactionIdentifier = String.valueOf( row.getCell(0).getStringCellValue());
              } else {
                transactionIdentifier =  String.valueOf(row.getCell(0).getNumericCellValue());
              }
              break;
          }
        }

        if (row.getCell(1) == null) {
          transactionDate = null;
        } else {
          switch (row.getCell(1) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              transactionDate = row.getCell(1).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(1))) {

                transactionDate =  String.valueOf(row.getCell(1).getStringCellValue());
              } else {
                transactionDate =   String.valueOf(((Double)row.getCell(1).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(2) == null) {
          transactionType = null;
        } else {
          switch (row.getCell(2) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              transactionType = row.getCell(2).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(2))) {

                transactionType =  String.valueOf(row.getCell(2).getStringCellValue());
              } else {
                transactionType =  String.valueOf(((Double)row.getCell(2).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(3) == null) {
          clerk = null;
        } else {
          switch (row.getCell(3) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              clerk = row.getCell(3).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(3))) {

                clerk =  String.valueOf(row.getCell(3).getStringCellValue());
              } else {
                clerk =   String.valueOf(((Double)row.getCell(3).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(4) == null) {
          note = null;
        } else {
          switch (row.getCell(4) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              note = row.getCell(4).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(4))) {

                note =  String.valueOf(row.getCell(4).getStringCellValue());
              } else {
                note =   String.valueOf(((Double)row.getCell(4).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(5) == null) {
          accountNumberDebtor = null;
        } else {
          switch (row.getCell(5) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              accountNumberDebtor = row.getCell(5).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(5))) {

                accountNumberDebtor =  String.valueOf(row.getCell(5).getStringCellValue());
              } else {
                accountNumberDebtor =  String.valueOf(((Double)row.getCell(5).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(6) == null) {
          amountDebtor = null;
        } else {
          switch (row.getCell(6) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              amountDebtor = row.getCell(6).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(6))) {

                amountDebtor = String.valueOf( row.getCell(6).getStringCellValue());
              } else {
                amountDebtor =   String.valueOf(((Double)row.getCell(6).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(7) == null) {
          accountNumberCreditor = null;
        } else {
          switch (row.getCell(7) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              accountNumberCreditor = String.valueOf(row.getCell(7).getStringCellValue());
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(7))) {

                accountNumberCreditor =  String.valueOf(row.getCell(7).getStringCellValue());
              } else {
                accountNumberCreditor =  String.valueOf(((Double)row.getCell(7).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(8) == null) {
          amountCreditor = null;
        } else {
          switch (row.getCell(8) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              amountCreditor = String.valueOf(row.getCell(8).getStringCellValue());
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(8))) {

                amountCreditor =  String.valueOf(row.getCell(8).getStringCellValue());
              } else {
                amountCreditor =  String.valueOf(((Double)row.getCell(8).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(9) == null) {
          state = null;
        } else {
          switch (row.getCell(9) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              state = row.getCell(9).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(9))) {

                state = String.valueOf( row.getCell(9).getStringCellValue());
              } else {
                state =  String.valueOf(((Double)row.getCell(9).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(10) == null) {
          message = null;
        } else {
          switch (row.getCell(10) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              message = row.getCell(10).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(10))) {

                message = String.valueOf( row.getCell(10).getStringCellValue());
              } else {
                message =  String.valueOf(((Double)row.getCell(10).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        Creditor creditor =new Creditor();
        creditor.setAccountNumber(String.valueOf(accountNumberCreditor));
        creditor.setAmount(String.valueOf(amountCreditor));
        Set<Creditor> creditors = new HashSet<>();
        creditors.add(creditor);

        Debtor debtor = new Debtor();
        debtor.setAccountNumber(String.valueOf(accountNumberDebtor));
        debtor.setAmount(String.valueOf(amountDebtor));
        Set<Debtor> debtors = new HashSet<>();
        debtors.add(debtor);

        JournalEntry journalEntry=new JournalEntry();
        journalEntry.setTransactionIdentifier(String.valueOf(transactionIdentifier));
        journalEntry.setTransactionDate(String.valueOf(transactionDate));
        journalEntry.setTransactionType(String.valueOf(transactionType));
        journalEntry.setClerk(String.valueOf(clerk));
        journalEntry.setNote(String.valueOf(note));
        journalEntry.setCreditors(creditors);
        journalEntry.setDebtors(debtors);
        journalEntry.setMessage(String.valueOf(message));
        journalEntry.setState(String.valueOf(state));

        this.userManagement.authenticate();
        this.ledgerManager.createJournalEntry(journalEntry);
      }

    } catch (IOException e) {
      e.printStackTrace();
    }
  }


}
