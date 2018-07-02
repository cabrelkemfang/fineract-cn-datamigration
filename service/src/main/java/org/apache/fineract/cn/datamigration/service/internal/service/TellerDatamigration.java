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
    Datavalidator.validator(worksheet,"TRUE","FALSE",8);
    Datavalidator.validatorState(worksheet,"ACTIVE","CLOSED","OPEN","PAUSED",10);

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

    IntStream.range(0, 11).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

    response.setHeader("Content-Disposition", "inline; filename=teller.xlsx");
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

  public void tellerSheetUpload(MultipartFile file){
    if (!file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
      throw new MultipartException("Only excel files accepted!");
    }
    try {
      XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
      Sheet firstSheet = workbook.getSheetAt(0);
      int rowCount = firstSheet.getLastRowNum() + 1;
      Row row;
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

      SimpleDateFormat date=new SimpleDateFormat("yyyy-MM-dd");

      for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
        row = firstSheet.getRow(rowIndex);
        if (row.getCell(0) == null) {
          officeIdentifier = null;
        } else {
          switch (row.getCell(0) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              officeIdentifier = row.getCell(0).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(0))) {

                officeIdentifier =  date.format(row.getCell(0).getDateCellValue());
              } else {
                officeIdentifier =  String.valueOf(row.getCell(0).getNumericCellValue());
              }
              break;
          }
        }

        if (row.getCell(1) == null) {
          code = null;
        } else {
          switch (row.getCell(1) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              code = row.getCell(1).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(1))) {

                code =   date.format(row.getCell(1).getDateCellValue());
              } else {
                code =   String.valueOf(((Double)row.getCell(1).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(2) == null) {
          password = null;
        } else {
          switch (row.getCell(2) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              password = row.getCell(2).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(2))) {

                password =  date.format(row.getCell(2).getDateCellValue());
              } else {
                password =  String.valueOf(((Double)row.getCell(2).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(3) == null) {
          cashdrawLimit = null;
        } else {
          switch (row.getCell(3) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              cashdrawLimit = row.getCell(3).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(3))) {

                cashdrawLimit =   date.format(row.getCell(3).getDateCellValue());
              } else {
                cashdrawLimit =   String.valueOf(((Double)row.getCell(3).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(4) == null) {
          tellerAccountIdentifier = null;
        } else {
          switch (row.getCell(4) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              tellerAccountIdentifier = row.getCell(4).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(4))) {

                tellerAccountIdentifier =   date.format(row.getCell(4).getDateCellValue());
              } else {
                tellerAccountIdentifier =   String.valueOf(((Double)row.getCell(4).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(5) == null) {
          vaultAccountIdentifier = null;
        } else {
          switch (row.getCell(5) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              vaultAccountIdentifier = row.getCell(5).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(5))) {

                vaultAccountIdentifier =   date.format(row.getCell(5).getDateCellValue());
              } else {
                vaultAccountIdentifier =  String.valueOf(((Double)row.getCell(5).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(6) == null) {
          chequesReceivableAccount = null;
        } else {
          switch (row.getCell(6) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              chequesReceivableAccount = row.getCell(6).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(6))) {

                chequesReceivableAccount =  date.format(row.getCell(6).getDateCellValue());
              } else {
                chequesReceivableAccount =   String.valueOf(((Double)row.getCell(6).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(7) == null) {
          cashOverShortAccount = null;
        } else {
          switch (row.getCell(7) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              cashOverShortAccount = String.valueOf(row.getCell(7).getStringCellValue());
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(7))) {

                cashOverShortAccount =  date.format(row.getCell(7).getDateCellValue());
              } else {
                cashOverShortAccount =  String.valueOf(((Double)row.getCell(7).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(8) == null) {
          denominationRequired = false;
        } else {
          switch (row.getCell(8) .getCellType()) {

            case Cell.CELL_TYPE_NUMERIC:

              if(((Double)row.getCell(8).getNumericCellValue()).intValue()==0){
                denominationRequired = Boolean.parseBoolean("false");
              }else{
                denominationRequired = Boolean.parseBoolean("true");
              }
              break;
          }
        }

        if (row.getCell(9) == null) {
          assignedEmployee = null;
        } else {
          switch (row.getCell(9) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              assignedEmployee = row.getCell(9).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(9))) {

                assignedEmployee =  date.format(row.getCell(9).getDateCellValue());
              } else {
                assignedEmployee =  String.valueOf(((Double)row.getCell(9).getNumericCellValue()).intValue());
              }
              break;
          }
        }
        if (row.getCell(10) == null) {
          state = null;
        } else {
          switch (row.getCell(10) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              state = row.getCell(10).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(10))) {

                state =  date.format(row.getCell(10).getDateCellValue());
              } else {
                state =  String.valueOf(((Double)row.getCell(10).getNumericCellValue()).intValue());
              }
              break;
          }
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

          this.userManagement.authenticate();
          this.tellerManager.create(officeIdentifier,teller);
        }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }

}
