package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.portfolio.api.v1.client.PortfolioManager;
import org.apache.fineract.cn.portfolio.api.v1.domain.ChargeDefinition;
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
import java.time.temporal.ChronoUnit;
import java.util.stream.IntStream;

@Service
public class ChargeDefinitionMigration {

  private final Logger logger;
  private  final PortfolioManager portfolioManager;

  @Autowired
  public ChargeDefinitionMigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                              final PortfolioManager portfolioManager) {
    super();
    this.logger = logger;
    this.portfolioManager = portfolioManager;
  }

  public void chargedefinitionSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Charge_Definition");

    Datavalidator.validator(worksheet,"TRUE","FALSE",12);
    Datavalidator.validator(worksheet,"TRUE","FALSE",16);
    Datavalidator.validatorType(worksheet,"FIXED","PROPORTIONAL","INTEREST",6);

    int startRowIndex = 0;
    int startColIndex = 0;

    Font font = worksheet.getWorkbook().createFont();
    XSSFCellStyle headerCellStyle = worksheet.getWorkbook().createCellStyle();

    headerCellStyle.setWrapText(true);
    headerCellStyle.setFont(font);
    XSSFRow rowHeader = worksheet.createRow((short) startRowIndex);
    rowHeader.setHeight((short) 500);

    XSSFCell cell1 = rowHeader.createCell(startColIndex+0);
    cell1.setCellValue(" Identifier");
    cell1.setCellStyle(headerCellStyle);

    XSSFCell cell2 = rowHeader.createCell(startColIndex+1);
    cell2.setCellValue("Name ");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Description");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Accrue Action");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5= rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Charge Action");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6= rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Amount");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7= rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Charge Method ");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8= rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Proportional To");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9= rowHeader.createCell(startColIndex+8);
    cell9.setCellValue("From Account Designator");
    cell9.setCellStyle(headerCellStyle);

    XSSFCell cell10= rowHeader.createCell(startColIndex+9);
    cell10.setCellValue("Accrual Account Designator");
    cell10.setCellStyle(headerCellStyle);

    XSSFCell cell11= rowHeader.createCell(startColIndex+10);
    cell11.setCellValue("To Account Designator");
    cell11.setCellStyle(headerCellStyle);

    XSSFCell cell12= rowHeader.createCell(startColIndex+11);
    cell12.setCellValue("For Cycle Size Unit");
    cell12.setCellStyle(headerCellStyle);

    XSSFCell cell13= rowHeader.createCell(startColIndex+12);
    cell13.setCellValue("Read Only");
    cell13.setCellStyle(headerCellStyle);

    XSSFCell cell14= rowHeader.createCell(startColIndex+13);
    cell14.setCellValue("For Segment Set");
    cell14.setCellStyle(headerCellStyle);

    XSSFCell cell15= rowHeader.createCell(startColIndex+14);
    cell15.setCellValue("From Segment");
    cell15.setCellStyle(headerCellStyle);

    XSSFCell cell16= rowHeader.createCell(startColIndex+15);
    cell16.setCellValue("To Segment");
    cell16.setCellStyle(headerCellStyle);

    XSSFCell cell17= rowHeader.createCell(startColIndex+16);
    cell17.setCellValue("Charge On Top");
    cell17.setCellStyle(headerCellStyle);


    IntStream.range(0, 11).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    response.setHeader("Content-Disposition", "attachment; filename=Charge_Definition+.xlsx");
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

  public void chargedefinitionSheetUpload(MultipartFile file) throws IOException {
    if (!file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
      throw new MultipartException("Only excel files accepted!");
    }
    try {
      XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
      Sheet firstSheet = workbook.getSheetAt(0);
      int rowCount = firstSheet.getLastRowNum() + 1;
      Row row;
      String identifier = null;
      String name = null;
      String description = null;
      String accrueAction = null;
      String chargeAction = null;
      Double amount = null;
      String chargeMethod = null;
      String proportionalTo = null;
      String fromAccountDesignator = null;
      String accrualAccountDesignator = null;
      String toAccountDesignator = null;
      String forCycleSizeUnit = null;
      Boolean readOnly = false;
      String forSegmentSet = null;
      String fromSegment = null;
      String toSegment = null;
      Boolean chargeOnTop = false;


      for (int rowIndex = 1; rowIndex < rowCount; rowIndex++) {
        row = firstSheet.getRow(rowIndex);

        if (row.getCell(0) == null) {
          identifier = null;
        } else {
          switch (row.getCell(0) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              identifier = row.getCell(0).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              identifier =  String.valueOf(row.getCell(0).getNumericCellValue());
              break;
          }
        }

        if (row.getCell(1) == null) {
          name = null;
        } else {
          switch (row.getCell(1) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              name = row.getCell(1).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              name =   String.valueOf(((Double)row.getCell(1).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(2) == null) {
          description = null;
        } else {
          switch (row.getCell(2) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              description = row.getCell(2).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              description =  String.valueOf(((Double)row.getCell(2).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(3) == null) {
          accrueAction = null;
        } else {
          switch (row.getCell(3) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              accrueAction = row.getCell(3).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              accrueAction =   String.valueOf(((Double)row.getCell(3).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(4) == null) {
          chargeAction = null;
        } else {
          switch (row.getCell(4) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              chargeAction = row.getCell(4).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              chargeAction =   String.valueOf(((Double)row.getCell(4).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(5) == null) {
          amount = null;
        } else {
          switch (row.getCell(5) .getCellType()) {

            case Cell.CELL_TYPE_NUMERIC:
              amount =  (Double)row.getCell(5).getNumericCellValue();
              break;
          }
        }

        if (row.getCell(6) == null) {
          chargeMethod = null;
        } else {
          switch (row.getCell(6) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              chargeMethod = row.getCell(6).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              chargeMethod =   String.valueOf(((Double)row.getCell(6).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(7) == null) {
          proportionalTo = null;
        } else {
          switch (row.getCell(7) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              proportionalTo = String.valueOf(row.getCell(7).getStringCellValue());
              break;

            case Cell.CELL_TYPE_NUMERIC:
              proportionalTo =  String.valueOf(((Double)row.getCell(7).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(8) == null) {
          fromAccountDesignator = null;
        } else {
          switch (row.getCell(8) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              fromAccountDesignator = String.valueOf(row.getCell(8).getStringCellValue());
              break;

            case Cell.CELL_TYPE_NUMERIC:
              fromAccountDesignator =  String.valueOf(((Double)row.getCell(8).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(9) == null) {
          accrualAccountDesignator = null;
        } else {
          switch (row.getCell(9) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              accrualAccountDesignator = row.getCell(9).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              accrualAccountDesignator =  String.valueOf(((Double)row.getCell(9).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(10) == null) {
          toAccountDesignator = null;
        } else {
          switch (row.getCell(10) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              toAccountDesignator = row.getCell(10).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              toAccountDesignator =  String.valueOf(((Double)row.getCell(10).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(11) == null) {
          forCycleSizeUnit = null;
        } else {
          switch (row.getCell(11) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              forCycleSizeUnit = row.getCell(11).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              forCycleSizeUnit =  String.valueOf(((Double)row.getCell(11).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(12) == null) {
          readOnly = false;
        } else {
          switch (row.getCell(12).getCellType()){
            case Cell.CELL_TYPE_STRING:
              readOnly = Boolean.parseBoolean(String.valueOf(row.getCell(12).getStringCellValue()));
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (((Double) row.getCell(12).getNumericCellValue()).intValue() == 0) {
                readOnly = Boolean.parseBoolean("false");
              } else {
                readOnly = Boolean.parseBoolean("true");
              }
          }
        }
        if (row.getCell(13) == null) {
          forSegmentSet = null;
        } else {
          switch (row.getCell(13) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              forSegmentSet = row.getCell(13).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              forSegmentSet =  String.valueOf(((Double)row.getCell(13).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(14) == null) {
          fromSegment = null;
        } else {
          switch (row.getCell(14) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              fromSegment = row.getCell(14).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              fromSegment =  String.valueOf(((Double)row.getCell(14).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(15) == null) {
          toSegment = null;
        } else {
          switch (row.getCell(15) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              toSegment = row.getCell(15).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              toSegment =  String.valueOf(((Double)row.getCell(15).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(16) == null) {
          chargeOnTop = false;
        } else {
          switch (row.getCell(16) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              chargeOnTop = Boolean.parseBoolean( String.valueOf(row.getCell(16).getStringCellValue()));
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if(((Double)row.getCell(16).getNumericCellValue()).intValue()==0){
                chargeOnTop = Boolean.parseBoolean("false");
              }else{
                chargeOnTop = Boolean.parseBoolean("true");
              }
          }
        }

        ChargeDefinition chargeDefinition = new ChargeDefinition();
        chargeDefinition.setIdentifier(String.valueOf(identifier));
        chargeDefinition.setName(String.valueOf(name));
        chargeDefinition.setDescription(String.valueOf(description));
        chargeDefinition.setAccrueAction(String.valueOf(accrueAction));
        chargeDefinition.setChargeAction(String.valueOf(chargeAction));
        chargeDefinition.setAmount(BigDecimal.valueOf(amount));
        chargeDefinition.setChargeMethod(ChargeDefinition.ChargeMethod.valueOf(chargeMethod));
        chargeDefinition.setProportionalTo(String.valueOf(proportionalTo));
        chargeDefinition.setFromAccountDesignator(String.valueOf(fromAccountDesignator));
        chargeDefinition.setAccrualAccountDesignator(String.valueOf(accrualAccountDesignator));
        chargeDefinition.setToAccountDesignator(String.valueOf(toAccountDesignator));
        chargeDefinition.setForCycleSizeUnit(ChronoUnit.valueOf(forCycleSizeUnit));
        chargeDefinition.setReadOnly(readOnly);
        chargeDefinition.setForSegmentSet(String.valueOf(forSegmentSet));
        chargeDefinition.setFromSegment(String.valueOf(fromSegment));
        chargeDefinition.setToSegment(String.valueOf(toSegment));
        chargeDefinition.setChargeOnTop(chargeOnTop);

        //this.portfolioManager.createChargeDefinition();
      }

    } catch (IOException e) {
      e.printStackTrace();
    }
  }


}
