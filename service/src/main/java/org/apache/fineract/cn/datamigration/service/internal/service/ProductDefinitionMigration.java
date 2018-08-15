package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.deposit.api.v1.client.DepositAccountManager;
import org.apache.fineract.cn.deposit.api.v1.definition.domain.Charge;
import org.apache.fineract.cn.deposit.api.v1.definition.domain.Currency;
import org.apache.fineract.cn.deposit.api.v1.definition.domain.ProductDefinition;
import org.apache.fineract.cn.deposit.api.v1.definition.domain.Term;
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
public class ProductDefinitionMigration {

  private final Logger logger;
  private final DepositAccountManager depositAccountManager;


  @Autowired
  public ProductDefinitionMigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                                      final DepositAccountManager depositAccountManager) {
    super();
    this.logger = logger;
    this.depositAccountManager = depositAccountManager;
  }

  public static void productDefinitionSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();

    XSSFSheet worksheet = workbook.createSheet("Product_Definitions");
    Datavalidator.validator(worksheet,"MONTH","YEAR",15);
    Datavalidator.validatorType(worksheet,"CHECKING","SAVINGS","SHARE",0);
    Datavalidator.validatorState(worksheet,"MATURITY","ANNUALLY","MONTHLY","QUARTERLY",16);
    Datavalidator.validator(worksheet,"TRUE","FALSE",21);
    Datavalidator.validator(worksheet,"TRUE","FALSE",23);
    Datavalidator.validator(worksheet,"TRUE","FALSE",24);

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
    cell19.setCellValue("  Income Account Identifier ");
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

    IntStream.range(0, 25).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    response.setHeader("Content-Disposition", "inline; filename=Product_Definitions.xlsx");
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

  public void productDefinitionSheetUpload(MultipartFile file){
    if (!file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
      throw new MultipartException("Only excel files accepted!");
    }
    try {
      XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
      Sheet firstSheet = workbook.getSheetAt(0);
      int rowCount = firstSheet.getLastRowNum() + 1;
      Row row;
      String identifier = null;
      String type = null;
      String name = null;
      String description = null;
      String currencyCode = null;
      String currencyName = null;
      String currencySign = null;
      String currencyScale = null;
      Double minimumBalance=null;
      String equityLedgerIdentifier = null;
      String cashAccountIdentifier = null;
      String expenseAccountIdentifier = null;
      String accrueAccountIdentifier = null;
      Double interest = null;
      String termPeriod= null;
      String termTimeUnit= null;
      String termInterestPayable = null;
      String chargeActionIdentifier = null;
      String chargeIncomeAccountIdentifier = null;
      String chargeName = null;
      String chargeDescription = null;
      Boolean chargeProportional = false;
      Double chargeAmount = null;
      Boolean flexible = false;
      Boolean active = false;

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
              type =  String.valueOf(row.getCell(0).getNumericCellValue());
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
          description = null;
        } else {
          switch (row.getCell(3) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              description = row.getCell(3).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              description =   String.valueOf(((Double)row.getCell(3).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(4) == null) {
          currencyCode = null;
        } else {
          switch (row.getCell(4) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              currencyCode = row.getCell(4).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              currencyCode =   String.valueOf(((Double)row.getCell(4).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(5) == null) {
          currencyName = null;
        } else {
          switch (row.getCell(5) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              currencyName = row.getCell(5).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              currencyName =  String.valueOf(((Double)row.getCell(5).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(6) == null) {
          currencySign = null;
        } else {
          switch (row.getCell(6) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              currencySign = row.getCell(6).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              currencySign =   String.valueOf(((Double)row.getCell(6).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(7) == null) {
          currencyScale = null;
        } else {
          switch (row.getCell(7) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              currencyScale = String.valueOf(row.getCell(7).getStringCellValue());
              break;

            case Cell.CELL_TYPE_NUMERIC:
              currencyScale =  String.valueOf(((Double)row.getCell(7).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(8) == null) {
          minimumBalance = null;
        } else {
          switch (row.getCell(8) .getCellType()) {

            case Cell.CELL_TYPE_NUMERIC:
              minimumBalance = ((Double)row.getCell(8).getNumericCellValue());
              break;
          }
        }

        if (row.getCell(9) == null) {
          equityLedgerIdentifier = null;
        } else {
          switch (row.getCell(9) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              equityLedgerIdentifier = row.getCell(9).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              equityLedgerIdentifier =  String.valueOf(((Double)row.getCell(9).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(10) == null) {
          cashAccountIdentifier = null;
        } else {
          switch (row.getCell(10) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              cashAccountIdentifier = row.getCell(10).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              cashAccountIdentifier =  String.valueOf(((Double)row.getCell(10).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(11) == null) {
          expenseAccountIdentifier = null;
        } else {
          switch (row.getCell(11) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              expenseAccountIdentifier = row.getCell(11).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              expenseAccountIdentifier =  String.valueOf(((Double)row.getCell(11).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(12) == null) {
          accrueAccountIdentifier = null;
        } else {
          switch (row.getCell(12) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              accrueAccountIdentifier = row.getCell(12).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              accrueAccountIdentifier =  String.valueOf(((Double)row.getCell(12).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(13) == null) {
          interest = null;
        } else {
          switch (row.getCell(13) .getCellType()) {

            case Cell.CELL_TYPE_NUMERIC:
              interest =  (Double)row.getCell(13).getNumericCellValue();
              break;
          }
        }
        if (row.getCell(14) == null) {
          termPeriod = null;
        } else {
          switch (row.getCell(14) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              termPeriod = row.getCell(14).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              termPeriod =  String.valueOf(((Double)row.getCell(14).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(15) == null) {
          termTimeUnit = null;
        } else {
          switch (row.getCell(15) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              termTimeUnit = row.getCell(15).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              termTimeUnit =  String.valueOf(((Double)row.getCell(15).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(16) == null) {
          termInterestPayable = null;
        } else {
          switch (row.getCell(16) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              termInterestPayable = row.getCell(16).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              termInterestPayable =  String.valueOf(((Double)row.getCell(16).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(17) == null) {
          chargeActionIdentifier = null;
        } else {
          switch (row.getCell(17) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              chargeActionIdentifier = row.getCell(17).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              chargeActionIdentifier =  String.valueOf(((Double)row.getCell(17).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(18) == null) {
          chargeIncomeAccountIdentifier = null;
        } else {
          switch (row.getCell(18) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              chargeIncomeAccountIdentifier = row.getCell(18).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              chargeIncomeAccountIdentifier =  String.valueOf(((Double)row.getCell(18).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(19) == null) {
          chargeName = null;
        } else {
          switch (row.getCell(19) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              chargeName = row.getCell(19).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              chargeName =  String.valueOf(((Double)row.getCell(19).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(20) == null) {
          chargeDescription = null;
        } else {
          switch (row.getCell(20) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              chargeDescription = row.getCell(20).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              chargeDescription =  String.valueOf(((Double)row.getCell(20).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(21) == null) {
          chargeProportional = false;
        } else {
          switch (row.getCell(21) .getCellType()) {
           case Cell.CELL_TYPE_STRING:
            chargeProportional = Boolean.parseBoolean( String.valueOf(row.getCell(22).getStringCellValue()));
            break;

          case Cell.CELL_TYPE_NUMERIC:
            if(((Double)row.getCell(21).getNumericCellValue()).intValue()==0){
              chargeProportional = Boolean.parseBoolean("false");
            }else{
              chargeProportional = Boolean.parseBoolean("true");
            }

            break;
          }
        }

        if (row.getCell(22) == null) {
          chargeAmount = null;
        } else {
          switch (row.getCell(22) .getCellType()) {

            case Cell.CELL_TYPE_NUMERIC:
              chargeAmount = (Double)row.getCell(22).getNumericCellValue();
              break;
          }
        }
        if (row.getCell(23) == null) {
          flexible = false;
        } else {
          switch (row.getCell(23) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              flexible = Boolean.parseBoolean( String.valueOf(row.getCell(24).getStringCellValue()));
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if(((Double)row.getCell(23).getNumericCellValue()).intValue()==0){
                flexible = Boolean.parseBoolean("false");
              }else{
                flexible = Boolean.parseBoolean("true");
              }

              break;
          }
        }
        if (row.getCell(24) == null) {
          active = false;
        } else {
          switch (row.getCell(24) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              active = Boolean.parseBoolean( String.valueOf(row.getCell(25).getStringCellValue()));
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if(((Double)row.getCell(24).getNumericCellValue()).intValue()==0){
                active = Boolean.parseBoolean("false");
              }else{
                active = Boolean.parseBoolean("true");
              }

              break;
          }
        }

        Currency currency = new Currency();
        currency.setCode(currencyCode);
        currency.setName(currencyName);
        currency.setSign(currencySign);
        currency.setScale(Integer.valueOf(currencyScale));

        Term term =new Term();
        term.setPeriod(Integer.valueOf(termPeriod));
        term.setTimeUnit(String.valueOf(termTimeUnit));
        term.setInterestPayable(String.valueOf(termInterestPayable));

        Charge charge = new Charge();
        charge.setActionIdentifier(String.valueOf(chargeActionIdentifier));
        charge.setIncomeAccountIdentifier(String.valueOf(chargeIncomeAccountIdentifier));
        charge.setName(String.valueOf(chargeName));
        charge.setDescription(String.valueOf(chargeDescription));
        charge.setProportional(chargeProportional);
        charge.setAmount(chargeAmount);
        Set<Charge> charges = new HashSet<>();
        charges.add(charge);

        ProductDefinition productDefinition = new ProductDefinition();
        productDefinition.setType(String.valueOf(type));
        productDefinition.setIdentifier(String.valueOf(identifier));
        productDefinition.setName(String.valueOf(name));
        productDefinition.setDescription(String.valueOf(description));
        productDefinition.setCurrency(currency);
        productDefinition.setMinimumBalance(minimumBalance);
        productDefinition.setEquityLedgerIdentifier(String.valueOf(equityLedgerIdentifier));
        productDefinition.setCashAccountIdentifier(String.valueOf(cashAccountIdentifier));
        productDefinition.setExpenseAccountIdentifier(String.valueOf(expenseAccountIdentifier));
        productDefinition.setAccrueAccountIdentifier(String.valueOf(accrueAccountIdentifier));
        productDefinition.setInterest(interest);
        productDefinition.setTerm(term);
        productDefinition.setCharges(charges);
        productDefinition.setFlexible(flexible);
        productDefinition.setActive(active);

        this.depositAccountManager.create(productDefinition);
      }

    } catch (IOException e) {
      e.printStackTrace();
    }
  }

}
