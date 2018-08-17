package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.office.api.v1.client.OrganizationManager;
import org.apache.fineract.cn.office.api.v1.domain.Address;
import org.apache.fineract.cn.office.api.v1.domain.Office;
import org.apache.fineract.cn.office.api.v1.domain.OfficePage;
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
import java.util.stream.IntStream;


@Service
public class OfficeBranchMigration {
  private final Logger logger;
  private final OrganizationManager organizationManager;



  @Autowired
  public OfficeBranchMigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                               final OrganizationManager organizationManager) {
    super();
    this.logger = logger;
    this.organizationManager = organizationManager;
  }
  public void branchSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Branch_Office");

    int startRowIndex = 0;
    int startColIndex = 0;
   // Datavalidator.validatorString(worksheet,officeIdentifier,1);

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
    cell2.setCellValue("Parent Identifier");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Name");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Description");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5= rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Street");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6= rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("City");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7= rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Region");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8= rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Postal Code");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9= rowHeader.createCell(startColIndex+8);
    cell9.setCellValue("Country Code");
    cell9.setCellStyle(headerCellStyle);

    XSSFCell cell10= rowHeader.createCell(startColIndex+9);
    cell10.setCellValue("Country");
    cell10.setCellStyle(headerCellStyle);

    IntStream.range(0, 10).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    response.setHeader("Content-Disposition", "inline; filename=Branch.xlsx");
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


  public void branchSheetUpload(MultipartFile file){
    if (!file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
      throw new MultipartException("Only excel files accepted!");
    }
    try {
      XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
      Sheet firstSheet = workbook.getSheetAt(0);
      int rowCount = firstSheet.getLastRowNum() + 1;
      Row row;
      String identifier =null ;
      String parentIdentifier  =null ;
      String name  =null;
      String description  =null ;
      String street   =null;
      String city   =null;
      String region   =null;
      String postalCode   =null;
      String countryCode   =null;
      String country   =null;
      boolean externalReferences=false;

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
          parentIdentifier = null;
        } else {
          switch (row.getCell(1) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              parentIdentifier = row.getCell(1).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                parentIdentifier =   String.valueOf(((Double)row.getCell(1).getNumericCellValue()).intValue());
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
          street = null;
        } else {
          switch (row.getCell(4) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              street = row.getCell(4).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                street =   String.valueOf(((Double)row.getCell(4).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(5) == null) {
          city = null;
        } else {
          switch (row.getCell(5) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              city = row.getCell(5).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                city =  String.valueOf(((Double)row.getCell(5).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(6) == null) {
          region = null;
        } else {
          switch (row.getCell(6) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              region = row.getCell(6).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                region =   String.valueOf(((Double)row.getCell(6).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(7) == null) {
          postalCode = null;
        } else {
          switch (row.getCell(7) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              postalCode = String.valueOf(row.getCell(7).getStringCellValue());
              break;

            case Cell.CELL_TYPE_NUMERIC:
                postalCode =  String.valueOf(((Double)row.getCell(7).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(8) == null) {
          countryCode = null;
        } else {
          switch (row.getCell(8) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              countryCode = String.valueOf(row.getCell(8).getStringCellValue());
              break;

            case Cell.CELL_TYPE_NUMERIC:
                countryCode =  String.valueOf(((Double)row.getCell(8).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(9) == null) {
          country = null;
        } else {
          switch (row.getCell(9) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              country = row.getCell(9).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                country =  String.valueOf(((Double)row.getCell(9).getNumericCellValue()).intValue());
              break;
          }
        }
        Address address = new Address();
        address.setStreet(String.valueOf(street));
        address.setCity(String.valueOf(city));
        address.setRegion(String.valueOf(region));
        address.setPostalCode(String.valueOf(postalCode));
        address.setCountryCode(String.valueOf(countryCode));
        address.setCountry(String.valueOf(country));
        Office office = new Office();
        office.setIdentifier(String.valueOf(identifier));
        office.setParentIdentifier(String.valueOf(parentIdentifier));
        office.setName(String.valueOf(name));
        office.setDescription(String.valueOf(description));
        office.setAddress(address);
        office.setExternalReferences(externalReferences);

        this.organizationManager.addBranch(parentIdentifier,office);
      }

    } catch (IOException e) {
      e.printStackTrace();
    }
  }

}
