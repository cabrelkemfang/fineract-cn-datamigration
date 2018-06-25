package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.datamigration.service.connector.UserManagement;
import org.apache.fineract.cn.teller.api.v1.client.TellerManager;
import org.apache.fineract.cn.teller.api.v1.domain.Teller;
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
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import java.util.stream.IntStream;
@Service
public class TellerDatamigration {
  private final Logger logger;
  private final TellerManager tellerManager;
  private final UserManagement userManagement;


  @Autowired
  public TellerDatamigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                                      final TellerManager tellerManager,
                                      final UserManagement userManagement) {
    super();
    this.logger = logger;
    this.tellerManager = tellerManager;
    this.userManagement = userManagement;
  }

  public static void tellerSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Tellers");

    int startRowIndex = 0;
    int startColIndex = 0;

    Font font = worksheet.getWorkbook().createFont();
    XSSFCellStyle headerCellStyle = worksheet.getWorkbook().createCellStyle();

    headerCellStyle.setWrapText(true);
    headerCellStyle.setFont(font);
    XSSFRow rowHeader = worksheet.createRow((short) startRowIndex);
    rowHeader.setHeight((short) 500);


    XSSFCell cell1 = rowHeader.createCell(startColIndex+0);
    cell1.setCellValue("Office Identifier");
    cell1.setCellStyle(headerCellStyle);

    XSSFCell cell2 = rowHeader.createCell(startColIndex+0);
    cell2.setCellValue("code");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+1);
    cell3.setCellValue("password");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+2);
    cell4.setCellValue("Cashdraw Limit");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5 = rowHeader.createCell(startColIndex+3);
    cell5.setCellValue("Teller Account Identifier");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6 = rowHeader.createCell(startColIndex+4);
    cell6.setCellValue("Vault Account Identifierault ");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7 = rowHeader.createCell(startColIndex+5);
    cell7.setCellValue("Cheques Receivable Account ");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8 = rowHeader.createCell(startColIndex+6);
    cell8.setCellValue("Cash Over Short Account ");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9 = rowHeader.createCell(startColIndex+7);
    cell9.setCellValue("Denomination Required ");
    cell9.setCellStyle(headerCellStyle);

    XSSFCell cell10= rowHeader.createCell(startColIndex+8);
    cell10.setCellValue("Assigned Employee");
    cell10.setCellStyle(headerCellStyle);

    XSSFCell cell11= rowHeader.createCell(startColIndex+9);
    cell11.setCellValue("state");
    cell11.setCellStyle(headerCellStyle);

    XSSFCell cell12= rowHeader.createCell(startColIndex+10);
    cell12.setCellValue("Created By");
    cell12.setCellStyle(headerCellStyle);

    XSSFCell cell13= rowHeader.createCell(startColIndex+11);
    cell13.setCellValue("Created On");
    cell13.setCellStyle(headerCellStyle);

    XSSFCell cell14= rowHeader.createCell(startColIndex+12);
    cell14.setCellValue("Last Modified By");
    cell14.setCellStyle(headerCellStyle);

    XSSFCell cell15= rowHeader.createCell(startColIndex+13);
    cell15.setCellValue("Last Modified On");
    cell15.setCellStyle(headerCellStyle);

    XSSFCell cell16= rowHeader.createCell(startColIndex+14);
    cell16.setCellValue("Last Opened By");
    cell16.setCellStyle(headerCellStyle);

    XSSFCell cell17= rowHeader.createCell(startColIndex+15);
    cell17.setCellValue("Last Opened On");
    cell17.setCellStyle(headerCellStyle);



    IntStream.range(0, 16).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

    response.setHeader("Content-Disposition", "inline; filename=teller.xlsx");
    response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    try {
      // Retrieve the output stream
      ServletOutputStream outputStream = response.getOutputStream();
      // Write to the output stream
      worksheet.getWorkbook().write(outputStream);
      // Flush the stream
      outputStream.flush();
    } catch (Exception e) {
    }

  }

  public void tellerSheetUpload(MultipartFile file){
    if (!file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
      throw new MultipartException("Only excel files accepted!");
    }
    try {

      XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
      Sheet firstSheet = workbook.getSheetAt(0);
      int offset = 0;
      int currentPosition = 0;

      for (Row nextRow : firstSheet) {
        int column = 0;
        Iterator<Cell> cellIterator = nextRow.cellIterator();
        if ((currentPosition++ > offset)) {
          String officeIdentifier = null;
          String code = null;
          String password = null;
          String cashdrawLimit = null;
          String tellerAccountIdentifier = null;
          String vaultAccountIdentifier = null;
          String chequesReceivableAccount = null;
          String cashOverShortAccount = null;
          Boolean denominationRequired = false;
          String assignedEmployee = null;
          String state = null;
          String createdBy = null;
          String createdOn = null;
          String lastModifiedBy = null;
          String lastModifiedOn = null;
          String lastOpenedBy = null;
          String lastOpenedOn = null;
          SimpleDateFormat date=new SimpleDateFormat("yyyy-MM-dd");

          while ((cellIterator.hasNext())) {
            XSSFCell cell = (XSSFCell) cellIterator.next();
            switch (cell.getCellType()) { // stop if blank field found
              case Cell.CELL_TYPE_BLANK:
                break;
            }
            if (column == 0) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  officeIdentifier = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    officeIdentifier =date.format(cell.getDateCellValue());
                  } else {
                    code=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  officeIdentifier = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 1) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  code = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    code =date.format(cell.getDateCellValue());
                  } else {
                    code=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  code = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 2) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  password = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    password =date.format(cell.getDateCellValue());
                  } else {
                    password=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  password = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 3) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  cashdrawLimit = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    cashdrawLimit =date.format(cell.getDateCellValue());
                  } else {
                    cashdrawLimit =Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  cashdrawLimit = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 4) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  tellerAccountIdentifier = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    tellerAccountIdentifier =date.format(cell.getDateCellValue());
                  } else {
                    tellerAccountIdentifier=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  tellerAccountIdentifier = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 5) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  vaultAccountIdentifier = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {
                    vaultAccountIdentifier =date.format(cell.getDateCellValue());
                  } else {
                    vaultAccountIdentifier=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  vaultAccountIdentifier = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 6) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  chequesReceivableAccount = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    chequesReceivableAccount =date.format(cell.getDateCellValue());
                  } else {
                    chequesReceivableAccount=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  chequesReceivableAccount = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 7) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  cashOverShortAccount = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    cashOverShortAccount =date.format(cell.getDateCellValue());
                  } else {
                    cashOverShortAccount=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  cashOverShortAccount = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 8) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                  denominationRequired = cell.getBooleanCellValue();
                  break;
              }
            }
            if (column == 9) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  assignedEmployee = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    assignedEmployee =date.format(cell.getDateCellValue());
                  } else {
                    assignedEmployee=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  assignedEmployee = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 10){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  state = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    state =date.format(cell.getDateCellValue());
                  } else {
                    state=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  state = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 11){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  createdBy = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    createdBy =date.format(cell.getDateCellValue());
                  } else {
                    createdBy=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  createdBy = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 12){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  createdOn = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    createdOn =date.format(cell.getDateCellValue());
                  } else {
                    createdOn=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  createdOn = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 13){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  assignedEmployee = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    assignedEmployee =date.format(cell.getDateCellValue());
                  } else {
                    assignedEmployee=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  assignedEmployee = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 14){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  lastModifiedBy = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    lastModifiedBy =date.format(cell.getDateCellValue());
                  } else {
                    lastModifiedBy=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  lastModifiedBy = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 15){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  lastModifiedOn = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    lastModifiedOn =date.format(cell.getDateCellValue());
                  } else {
                    lastModifiedOn=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  lastModifiedOn = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 16){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  lastOpenedBy = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    lastOpenedBy =date.format(cell.getDateCellValue());
                  } else {
                    lastOpenedBy=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  lastOpenedBy = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 17){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  lastOpenedOn = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    lastOpenedOn =date.format(cell.getDateCellValue());
                  } else {
                    lastOpenedOn=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  lastOpenedOn = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            column++;
          }
          Teller teller= new Teller();
          teller.setCode(String.valueOf(code));
          teller.setPassword(String.valueOf(password));
          BigDecimal cashdraw = new BigDecimal(cashdrawLimit);
          teller.setCashdrawLimit(cashdraw);

          teller.setTellerAccountIdentifier(String.valueOf(tellerAccountIdentifier));
          teller.setVaultAccountIdentifier(String.valueOf(vaultAccountIdentifier));
          teller.setChequesReceivableAccount(chequesReceivableAccount);
          teller.setCashOverShortAccount(cashOverShortAccount);
          teller.setDenominationRequired(denominationRequired);
          teller.setAssignedEmployee(assignedEmployee);
          teller.setState(state);
          teller.setCreatedBy(String.valueOf(createdBy));
          teller.setCreatedOn(String.valueOf(createdOn));
          teller.setLastModifiedBy(String.valueOf(lastModifiedBy));
          teller.setLastModifiedOn(String.valueOf(lastModifiedOn));
          teller.setLastOpenedBy(String.valueOf(lastOpenedBy));
          teller.setLastOpenedOn(String.valueOf(lastOpenedOn));

          this.userManagement.authenticate();
          this.tellerManager.create(officeIdentifier,teller);
        }

      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

}
