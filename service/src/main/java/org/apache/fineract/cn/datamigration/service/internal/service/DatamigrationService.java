/*
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.customer.api.v1.client.CustomerManager;
import org.apache.fineract.cn.customer.api.v1.domain.Address;
import org.apache.fineract.cn.customer.api.v1.domain.ContactDetail;
import org.apache.fineract.cn.customer.api.v1.domain.Customer;
import org.apache.fineract.cn.customer.catalog.api.v1.domain.Value;
import org.apache.fineract.cn.datamigration.service.connector.UserManagement;
import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.lang.DateOfBirth;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.slf4j.Logger;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartException;
import org.springframework.web.multipart.MultipartFile;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.stream.IntStream;


@Service
public class DatamigrationService {
  private final Logger logger;
  private final CustomerManager customerManager;
  private final UserManagement userManagement;


  @Autowired
  public DatamigrationService(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                                final CustomerManager customerManager,
                                final UserManagement userManagement) {
     super();
     this.logger = logger;
     this.customerManager = customerManager;
     this.userManagement = userManagement;
  }

  public final ByteArrayInputStream customersFormDownload(){

     ByteArrayOutputStream outByteStream = new ByteArrayOutputStream();
     XSSFWorkbook workbook = new XSSFWorkbook();
     XSSFSheet worksheet = workbook.createSheet("customers");

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
     cell2.setCellValue("Type");
     cell2.setCellStyle(headerCellStyle);


     XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
     cell3.setCellValue("Given Name");
     cell3.setCellStyle(headerCellStyle);

     XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
     cell4.setCellValue("Middle Name");
     cell4.setCellStyle(headerCellStyle);

     XSSFCell cell5 = rowHeader.createCell(startColIndex+4);
     cell5.setCellValue("Surname ");
     cell5.setCellStyle(headerCellStyle);

     //dateOfBirth
     XSSFCell cell6 = rowHeader.createCell(startColIndex+5);
     cell6.setCellValue("Year ");
     cell6.setCellStyle(headerCellStyle);

     XSSFCell cell7 = rowHeader.createCell(startColIndex+6);
     cell7.setCellValue("Month ");
     cell7.setCellStyle(headerCellStyle);

     XSSFCell cell8 = rowHeader.createCell(startColIndex+7);
     cell8.setCellValue("Day ");
     cell8.setCellStyle(headerCellStyle);

     XSSFCell cell9= rowHeader.createCell(startColIndex+8);
     cell9.setCellValue("Member");
     cell9.setCellStyle(headerCellStyle);

     XSSFCell cell10= rowHeader.createCell(startColIndex+9);
     cell10.setCellValue("Account Beneficiary");
     cell10.setCellStyle(headerCellStyle);

     XSSFCell cell11= rowHeader.createCell(startColIndex+10);
     cell11.setCellValue("Reference Customer");
     cell11.setCellStyle(headerCellStyle);

     XSSFCell cell12= rowHeader.createCell(startColIndex+11);
     cell12.setCellValue("Assigned Office");
     cell12.setCellStyle(headerCellStyle);

     XSSFCell cell13= rowHeader.createCell(startColIndex+12);
     cell13.setCellValue("Assigned Employee");
     cell13.setCellStyle(headerCellStyle);

     //address
     XSSFCell cell14= rowHeader.createCell(startColIndex+13);
     cell14.setCellValue("Street");
     cell14.setCellStyle(headerCellStyle);

     XSSFCell cell15= rowHeader.createCell(startColIndex+14);
     cell15.setCellValue("City");
     cell15.setCellStyle(headerCellStyle);

     XSSFCell cell16= rowHeader.createCell(startColIndex+15);
     cell16.setCellValue("Region");
     cell16.setCellStyle(headerCellStyle);

     XSSFCell cell17= rowHeader.createCell(startColIndex+16);
     cell17.setCellValue("Postal Code");
     cell17.setCellStyle(headerCellStyle);

     XSSFCell cell18= rowHeader.createCell(startColIndex+17);
     cell18.setCellValue("Country Code");
     cell18.setCellStyle(headerCellStyle);

     XSSFCell cell19= rowHeader.createCell(startColIndex+18);
     cell19.setCellValue("Country");
     cell19.setCellStyle(headerCellStyle);

     //contactDetail
     XSSFCell cell20= rowHeader.createCell(startColIndex+19);
     cell20.setCellValue("Type");
     cell20.setCellStyle(headerCellStyle);

     XSSFCell cell21= rowHeader.createCell(startColIndex+20);
     cell21.setCellValue("Group");
     cell21.setCellStyle(headerCellStyle);

     XSSFCell cell22= rowHeader.createCell(startColIndex+21);
     cell22.setCellValue("Value");
     cell22.setCellStyle(headerCellStyle);

     XSSFCell cell23= rowHeader.createCell(startColIndex+22);
     cell23.setCellValue("Preference Level");
     cell23.setCellStyle(headerCellStyle);

     XSSFCell cell24= rowHeader.createCell(startColIndex+23);
     cell24.setCellValue("Validated");
     cell24.setCellStyle(headerCellStyle);


     XSSFCell cell25= rowHeader.createCell(startColIndex+24);
     cell25.setCellValue("Current State");
     cell25.setCellStyle(headerCellStyle);

     XSSFCell cell26= rowHeader.createCell(startColIndex+25);
     cell26.setCellValue("Application Date");
     cell26.setCellStyle(headerCellStyle);

     //value
     XSSFCell cell27= rowHeader.createCell(startColIndex+26);
     cell27.setCellValue("Catalog Identifier");
     cell27.setCellStyle(headerCellStyle);

     XSSFCell cell28= rowHeader.createCell(startColIndex+27);
     cell28.setCellValue("Field Identifier");
     cell28.setCellStyle(headerCellStyle);

     XSSFCell cell29= rowHeader.createCell(startColIndex+28);
     cell29.setCellValue("Value");
     cell29.setCellStyle(headerCellStyle);

     XSSFCell cell30= rowHeader.createCell(startColIndex+29);
     cell30.setCellValue("Created By");
     cell30.setCellStyle(headerCellStyle);

     XSSFCell cell31= rowHeader.createCell(startColIndex+30);
     cell31.setCellValue("Created On");
     cell31.setCellStyle(headerCellStyle);

     XSSFCell cell32= rowHeader.createCell(startColIndex+31);
     cell32.setCellValue("LastModified By");
     cell32.setCellStyle(headerCellStyle);

     XSSFCell cell33= rowHeader.createCell(startColIndex+32);
     cell33.setCellValue("LastModified On");
     cell33.setCellStyle(headerCellStyle);

     IntStream.range(0, 33).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

     try {
        worksheet.getWorkbook().write(outByteStream);
        // Flush the stream
        outByteStream.flush();
     } catch (Exception e) {
        System.out.println("Unable to write report to the output stream");
     }

     return new ByteArrayInputStream(outByteStream.toByteArray());

  }

  public void customersFormUpload(MultipartFile file){
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
          String type = null;
          String givenName = null;
          String middleName = null;
          String surname = null;

          String year = null;
          String month = null;
          String day = null;

          String member=null;
          String accountBeneficiary = null;
          String referenceCustomer = null;
          String assignedOffice = null;
          String assignedEmployee = null;

          String street = null;
          String city = null;
          String region = null;
          String postalCode = null;
          String countryCode = null;
          String country = null;

          String typecontactDetail = null;
          String group = null;
          String value = null;
          String preferenceLevel = null;
          String validated = null;

          String currentState = null;
          String applicationDate = null;

          String catalogIdentifier = null;
          String fieldIdentifier = null;
          String value2 = null;

          String createdBy = null;
          String createdOn = null;
          String lastModifiedBy = null;
          String lastModifiedOn = null;


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
                  identifier = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  identifier = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 1) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  type = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  type = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  type = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 2) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  givenName = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  givenName = Integer.toString((int) cell.getNumericCellValue());
                  break;
                case Cell.CELL_TYPE_BOOLEAN:
                  givenName = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 3) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  middleName = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  middleName = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  middleName = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 5) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  surname = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  surname = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  surname = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 6) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  year = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  year = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  year = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 7) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  month = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  month = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  month = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 8) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  day = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  day = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  day = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 9) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  member = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  member = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  member = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 10){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  accountBeneficiary = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  accountBeneficiary = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  accountBeneficiary = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 11){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  referenceCustomer = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  referenceCustomer = Integer.toString((int) cell.getNumericCellValue());
                  break;
                case Cell.CELL_TYPE_BOOLEAN:
                  referenceCustomer = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 12){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  assignedOffice = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  assignedOffice = Integer.toString((int) cell.getNumericCellValue());
                  break;
                case Cell.CELL_TYPE_BOOLEAN:
                  assignedOffice = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 13){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  assignedEmployee = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  assignedEmployee = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  assignedEmployee = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 14){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  street = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  street = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  street = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 15){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  city = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  city = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  city = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 16){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  region = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  region = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  region = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 17){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  postalCode = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  postalCode = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  postalCode = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 18){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  countryCode = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  countryCode = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  countryCode = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 19){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  country = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  country = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  country = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 20){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  typecontactDetail = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  typecontactDetail = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  typecontactDetail = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 21){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  group = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  group = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  group = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 22){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  value = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  value = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  value = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 23){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  preferenceLevel = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  preferenceLevel = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  preferenceLevel = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 24){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  validated = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  //validated = Integer.toString((int) cell.getNumericCellValue());
                  break;
                case Cell.CELL_TYPE_BOOLEAN:
                  validated = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 25){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  currentState = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  currentState = Integer.toString((int) cell.getNumericCellValue());
                  break;
                case Cell.CELL_TYPE_BOOLEAN:
                  currentState = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 26){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  applicationDate = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  applicationDate = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  applicationDate = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 27){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  catalogIdentifier = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  catalogIdentifier = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  catalogIdentifier = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 28){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  fieldIdentifier = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  fieldIdentifier = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  fieldIdentifier = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 29){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  value2 = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  value2 = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  value2 = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 30){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  createdBy = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  createdBy = Integer.toString((int) cell.getNumericCellValue());
                  break;
                case Cell.CELL_TYPE_BOOLEAN:
                  createdBy = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 30){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  createdOn = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  createdOn = Integer.toString((int) cell.getNumericCellValue());
                  break;
                case Cell.CELL_TYPE_BOOLEAN:
                  createdOn = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 31){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  lastModifiedBy = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  lastModifiedBy = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  lastModifiedBy = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 32) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  lastModifiedOn = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  lastModifiedOn = Integer.toString((int) cell.getNumericCellValue());
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  lastModifiedOn = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }

            column++;
          }

          DateOfBirth dateOfBirth = new DateOfBirth();
          dateOfBirth.setYear(Integer.parseInt(year));
          dateOfBirth.setMonth(Integer.parseInt(month));
          dateOfBirth.setDay(Integer.parseInt(day));

          Address address = new Address();
          address.setStreet(street);
          address.setCity(city);
          address.setRegion(region);
          address.setPostalCode(postalCode);
          address.setCountryCode(countryCode);
          address.setCountry(country);

          ContactDetail contactDetail=new ContactDetail();
          contactDetail.setType(typecontactDetail);
          contactDetail.setGroup(group);
          contactDetail.setValue(value);
          contactDetail.setPreferenceLevel(Integer.parseInt(preferenceLevel));
          contactDetail.setValidated(Boolean.parseBoolean(validated));

          Value value1=new Value();
          value1.setCatalogIdentifier(catalogIdentifier);
          value1.setFieldIdentifier(fieldIdentifier);
          value1.setValue(value2);

          Customer customer = new Customer();
          customer.setIdentifier(identifier);
          customer.setType(type);
          customer.setGivenName(givenName);
          customer.setMiddleName(middleName);
          customer.setSurname(surname);
          customer.setDateOfBirth(dateOfBirth);
          customer.setMember(Boolean.parseBoolean(member));
          customer.setAccountBeneficiary(accountBeneficiary);
          customer.setReferenceCustomer(referenceCustomer);
          customer.setAssignedOffice(assignedOffice);
          customer.setAssignedEmployee(assignedEmployee);
          customer.setAddress(address);
          customer.setContactDetails(Collections.singletonList(contactDetail));
          customer.setCurrentState(currentState);
          customer.setApplicationDate(applicationDate);
         // customer.setCustomValues(Collections.singletonList(value1));
          customer.setCreatedBy(createdBy);
          customer.setCreatedOn(createdOn);
          customer.setLastModifiedBy(lastModifiedBy);
          customer.setLastModifiedOn(lastModifiedOn);

          this.userManagement.authenticate();
          this.customerManager.createCustomer(customer);

        }

      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}

