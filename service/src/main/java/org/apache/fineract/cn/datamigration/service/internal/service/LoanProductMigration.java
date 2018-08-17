package org.apache.fineract.cn.datamigration.service.internal.service;


import org.apache.fineract.cn.accounting.api.v1.client.LedgerManager;
import org.apache.fineract.cn.accounting.api.v1.domain.LedgerPage;
import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.portfolio.api.v1.client.PortfolioManager;
import org.apache.fineract.cn.portfolio.api.v1.domain.*;
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
import java.util.HashSet;
import java.util.Set;
import java.util.stream.IntStream;

@Service
public class LoanProductMigration {

  private final Logger logger;
  private  final PortfolioManager portfolioManager;
  private final LedgerManager ledgerManager;

  @Autowired
  public LoanProductMigration(@Qualifier(ServiceConstants.LOGGER_NAME)
                                final Logger logger,
                                final PortfolioManager portfolioManager,
                                final LedgerManager ledgerManager) {
    super();
    this.logger = logger;
    this.portfolioManager = portfolioManager;
    this.ledgerManager = ledgerManager;
  }

  public  void productSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Products");

    final LedgerPage currentLedgerPage = this.ledgerManager.fetchLedgers(false, null, null, null, null, null, null);
    int sizeOfLedger=currentLedgerPage.getLedgers().size();

    String [] ledgerIdentifier = new String[sizeOfLedger];
    for(int i=0;i<sizeOfLedger;i++){
      ledgerIdentifier[i]=currentLedgerPage.getLedgers().get(i).getIdentifier();
    }

    Datavalidator.validator(worksheet,"CURRENT_BALANCE","BEGINNING_BALANCE",9);
    Datavalidator.validator(worksheet,"TRUE","FALSE",12);
    Datavalidator.validatorString(worksheet,ledgerIdentifier,18);
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
    cell17.setCellValue("Account Assignment Designator");
    cell17.setCellStyle(headerCellStyle);

    XSSFCell cell18= rowHeader.createCell(startColIndex+17);
    cell18.setCellValue("Account Identifier");
    cell18.setCellStyle(headerCellStyle);

    XSSFCell cell19= rowHeader.createCell(startColIndex+18);
    cell19.setCellValue("Ledger Identifier");
    cell19.setCellStyle(headerCellStyle);

    XSSFCell cell20= rowHeader.createCell(startColIndex+19);
    cell20.setCellValue("Alternative Account Number");
    cell20.setCellStyle(headerCellStyle);




    IntStream.range(0, 20).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    response.setHeader("Content-Disposition", "inline; filename=Products.xlsx");
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

  public void productSheetUpload(MultipartFile file){
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
      String termRangeTemporalUnit = null;
      String termRangeMinimum = null;
      String termRangeMaximum = null;
      Double balanceRangeMinimum = null;
      Double balanceRangeMaximum = null;
      Double interestRangeMinimum = null;
      Double interestRangeMaximum = null;
      String interestBasis = null;
      String patternPackage = null;
      String description = null;
      Boolean enabled=false;
      String currencyCode = null;
      String minorCurrencyUnitDigits=null;
      String designator=null;
      String accountIdentifier = null;
      String ledgerIdentifier = null;
      String alternativeAccountNumber = null;
      String parameters = null;
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
          termRangeTemporalUnit = null;
        } else {
          switch (row.getCell(2) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              termRangeTemporalUnit = row.getCell(2).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                termRangeTemporalUnit =  String.valueOf(((Double)row.getCell(2).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(3) == null) {
          termRangeMinimum = null;
        } else {
          switch (row.getCell(3) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              termRangeMinimum = row.getCell(3).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              termRangeMinimum =  String.valueOf(((Double)row.getCell(3).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(4) == null) {
          termRangeMaximum = null;
        } else {
          switch (row.getCell(4) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              termRangeMaximum = row.getCell(4).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              termRangeMaximum =  String.valueOf(((Double)row.getCell(4).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(5) == null) {
          balanceRangeMinimum = null;
        } else {
          switch (row.getCell(5) .getCellType()) {

            case Cell.CELL_TYPE_NUMERIC:
                balanceRangeMinimum =  (Double)row.getCell(5).getNumericCellValue();
              break;
          }
        }

        if (row.getCell(6) == null) {
          balanceRangeMaximum = null;
        } else {
          switch (row.getCell(6) .getCellType()) {

            case Cell.CELL_TYPE_NUMERIC:
                balanceRangeMaximum = (Double)row.getCell(6).getNumericCellValue();
              break;
          }
        }

        if (row.getCell(7) == null) {
          interestRangeMinimum = null;
        } else {
          switch (row.getCell(7) .getCellType()) {

            case Cell.CELL_TYPE_NUMERIC:
                interestRangeMinimum = (Double)row.getCell(7).getNumericCellValue();
              break;
          }
        }

        if (row.getCell(8) == null) {
          interestRangeMaximum = null;
        } else {
          switch (row.getCell(8) .getCellType()) {

            case Cell.CELL_TYPE_NUMERIC:
                interestRangeMaximum =(Double)row.getCell(8).getNumericCellValue();
              break;
          }
        }

        if (row.getCell(9) == null) {
          interestBasis = null;
        } else {
          switch (row.getCell(9) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              interestBasis = row.getCell(9).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                interestBasis =  String.valueOf(((Double)row.getCell(9).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(10) == null) {
          patternPackage = null;
        } else {
          switch (row.getCell(10) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              patternPackage = row.getCell(10).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                patternPackage =  String.valueOf(((Double)row.getCell(10).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(11) == null) {
          description = null;
        } else {
          switch (row.getCell(11) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              description = row.getCell(11).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                description =  String.valueOf(((Double)row.getCell(11).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(12) == null) {
          enabled = false;
        } else {
          switch (row.getCell(12) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              enabled = Boolean.parseBoolean( String.valueOf(row.getCell(12).getStringCellValue()));
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if(((Double)row.getCell(12).getNumericCellValue()).intValue()==0){
                enabled = Boolean.parseBoolean("false");
              }else{
                enabled = Boolean.parseBoolean("true");
              }

              break;
          }
        }
        if (row.getCell(13) == null) {
          currencyCode = null;
        } else {
          switch (row.getCell(13) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              currencyCode = row.getCell(13).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                currencyCode =  String.valueOf(((Double)row.getCell(13).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(14) == null) {
          minorCurrencyUnitDigits = null;
        } else {
          switch (row.getCell(14) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              minorCurrencyUnitDigits = row.getCell(14).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                minorCurrencyUnitDigits =  String.valueOf(((Double)row.getCell(14).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(15) == null) {
          designator = null;
        } else {
          switch (row.getCell(15) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              designator = row.getCell(15).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                designator =  String.valueOf(((Double)row.getCell(15).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(16) == null) {
          accountIdentifier = null;
        } else {
          switch (row.getCell(16) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              accountIdentifier = row.getCell(16).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                accountIdentifier =  String.valueOf(((Double)row.getCell(16).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(17) == null) {
          ledgerIdentifier = null;
        } else {
          switch (row.getCell(17) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              ledgerIdentifier = row.getCell(17).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                ledgerIdentifier =  String.valueOf(((Double)row.getCell(17).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(18) == null) {
          alternativeAccountNumber = null;
        } else {
          switch (row.getCell(18) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              alternativeAccountNumber = row.getCell(18).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              alternativeAccountNumber =  String.valueOf(((Double)row.getCell(18).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(19) == null) {
          parameters = null;
        } else {
          switch (row.getCell(19) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              parameters = row.getCell(19).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                parameters =  String.valueOf(((Double)row.getCell(19).getNumericCellValue()).intValue());
              break;
          }
        }

        TermRange termRange= new TermRange();
        termRange.setMaximum(Integer.valueOf(termRangeMinimum));
        termRange.setMaximum(Integer.valueOf(termRangeMaximum));
        termRange.setTemporalUnit(ChronoUnit.valueOf(termRangeTemporalUnit));
        BalanceRange balanceRange =new BalanceRange();
        balanceRange.setMaximum(BigDecimal.valueOf(balanceRangeMaximum));
        balanceRange.setMinimum(BigDecimal.valueOf(balanceRangeMinimum));

        InterestRange interestRange = new InterestRange();
        interestRange.setMaximum(BigDecimal.valueOf(interestRangeMaximum));
        interestRange.setMinimum(BigDecimal.valueOf(interestRangeMinimum));

        AccountAssignment accountAssignment = new AccountAssignment();
        accountAssignment.setAccountIdentifier(String.valueOf(accountIdentifier));
        accountAssignment.setAlternativeAccountNumber(String.valueOf(alternativeAccountNumber));
        accountAssignment.setDesignator(String.valueOf(designator));
        accountAssignment.setLedgerIdentifier(String.valueOf(ledgerIdentifier));
        Set<AccountAssignment> accountAssignments = new HashSet<>();
        accountAssignments.add(accountAssignment);

        Product product = new Product();
        product.setIdentifier(String.valueOf(identifier));
        product.setName(String.valueOf(name));
        product.setTermRange(termRange);
        product.setBalanceRange(balanceRange);
        product.setInterestBasis(InterestBasis.valueOf(interestBasis));
        product.setInterestRange(interestRange);
        product.setPatternPackage(String.valueOf(patternPackage));
        product.setDescription(String.valueOf(description));
        product.setEnabled(enabled);
        product.setCurrencyCode(String.valueOf(currencyCode));
        product.setMinorCurrencyUnitDigits(Integer.valueOf(minorCurrencyUnitDigits));
        product.setAccountAssignments(accountAssignments);
        product.setParameters(parameters);

        this.portfolioManager.createProduct(product);
      }

    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}

