package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.datamigration.service.connector.UserManagement;

import org.apache.fineract.cn.office.api.v1.client.OrganizationManager;
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
import java.text.SimpleDateFormat;
import java.util.Iterator;
import java.util.stream.IntStream;

@Service
public class OfficeDatamigrationService {
  private final Logger logger;
  private final OrganizationManager organizationManager;
  private final UserManagement userManagement;

  @Autowired
  public OfficeDatamigrationService(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                                      final OrganizationManager organizationManager,
                                      final UserManagement userManagement) {
    super();
    this.logger = logger;
    this.organizationManager = organizationManager;
    this.userManagement = userManagement;
  }

  public void officeSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Offices");

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

    XSSFCell cell11= rowHeader.createCell(startColIndex+10);
    cell11.setCellValue("External References");
    cell11.setCellStyle(headerCellStyle);

    IntStream.range(0, 11).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    response.setHeader("Content-Disposition", "inline; filename=Offices.xlsx");
    response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    try {
      // Retrieve the output stream
      ServletOutputStream outputStream = response.getOutputStream();
      // Write to the output stream
      worksheet.getWorkbook().write(outputStream);
      // Flush the stream
      outputStream.flush();
    } catch (Exception e) {
      System.out.println("Unable to write report to the output stream");
    }

  }


  public void officeSheetUpload(MultipartFile file){
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
          String identifier = null;
          String parentIdentifier = null;
          String name = null;
          String description = null;

          String street = null;
          String city = null;
          String region = null;
          String postalCode = null;
          String countryCode = null;
          String country = null;

          Boolean externalReferences = false;

          SimpleDateFormat date = new SimpleDateFormat("yyyy-MM-dd");
          while ((cellIterator.hasNext())) {
            XSSFCell cell = (XSSFCell) cellIterator.next();
            switch (cell.getCellType()) { // stop if blank field found
              case Cell.CELL_TYPE_BLANK:
                break;
            }
            if (column == 0) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  identifier = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    identifier = date.format(cell.getDateCellValue());
                  } else {
                    identifier = Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  identifier = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 1) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  parentIdentifier = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    parentIdentifier = date.format(cell.getDateCellValue());
                  } else {
                    parentIdentifier = Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  parentIdentifier = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 2) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  name = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    name = date.format(cell.getDateCellValue());
                  } else {
                    name = Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  name = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 3) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  description = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    description = date.format(cell.getDateCellValue());
                  } else {
                    description = Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  description = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 4) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  street = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    street = date.format(cell.getDateCellValue());
                  } else {
                    street = Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  street = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 5) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  city = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    city = date.format(cell.getDateCellValue());
                  } else {
                    city = Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  city = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 6) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  region = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    region = date.format(cell.getDateCellValue());
                  } else {
                    region = Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  region = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 7) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  postalCode = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    postalCode = date.format(cell.getDateCellValue());
                  } else {
                    postalCode = Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  postalCode = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 8) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  countryCode = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    countryCode = date.format(cell.getDateCellValue());
                  } else {
                    countryCode = Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  countryCode = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 9) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  country = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    country = date.format(cell.getDateCellValue());
                  } else {
                    country = Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  country = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 10) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_BOOLEAN:
                  externalReferences = cell.getBooleanCellValue();
                  break;
              }
            }

            column++;
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

          this.userManagement.authenticate();
          this.organizationManager.createOffice(office);
        }
      }

    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}
