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
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.slf4j.Logger;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartException;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.IntStream;


@Service
public class CustomerDatamigrationService {
  private final Logger logger;
  private final CustomerManager customerManager;
  private final UserManagement userManagement;


  @Autowired
  public CustomerDatamigrationService(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                                      final CustomerManager customerManager,
                                      final UserManagement userManagement) {
     super();
     this.logger = logger;
     this.customerManager = customerManager;
     this.userManagement = userManagement;
  }

  public static void customersSheetDownload(HttpServletResponse response){
     XSSFWorkbook workbook = new XSSFWorkbook();
     XSSFSheet worksheet = workbook.createSheet("Customers");

    Datavalidator.validator(worksheet,"PERSON","BUSINESS",1);
    Datavalidator.validatorState(worksheet,"PENDING","ACTIVE","LOCKED","CLOSED",24);

    Datavalidator.validator(worksheet,"BUSINESS","PRIVATE",20);
    Datavalidator.validatorType(worksheet,"EMAIL","PHONE","MOBILE",19);

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
     cell32.setCellValue("Last Modified By");
     cell32.setCellStyle(headerCellStyle);

     XSSFCell cell33= rowHeader.createCell(startColIndex+32);
     cell33.setCellValue("Last Modified On");
     cell33.setCellStyle(headerCellStyle);

     IntStream.range(0, 33).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

    response.setHeader("Content-Disposition", "inline; filename=Customer.xlsx");
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

  public void customersSheetUpload(MultipartFile file){
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

          Boolean member=false;
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
          Boolean validated = false;

          String currentState = null;
          String applicationDate = null;

          String catalogIdentifier = null;
          String fieldIdentifier = null;
          String value2 = null;

          String createdBy = null;
          String createdOn = null;
          String lastModifiedBy = null;
          String lastModifiedOn = null;
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
                  identifier = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    identifier =date.format(cell.getDateCellValue());
                  } else {
                    identifier=Integer.toString((int) cell.getNumericCellValue());
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
                  type = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    type =date.format(cell.getDateCellValue());
                  } else {
                    type=Integer.toString((int) cell.getNumericCellValue());
                  }
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
                  if (DateUtil.isCellDateFormatted(cell)) {

                    givenName =date.format(cell.getDateCellValue());
                  } else {
                    givenName =Integer.toString((int) cell.getNumericCellValue());
                  }
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
                  if (DateUtil.isCellDateFormatted(cell)) {

                    middleName =date.format(cell.getDateCellValue());
                  } else {
                    middleName=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  middleName = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 4) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  surname = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {
                    surname =date.format(cell.getDateCellValue());
                  } else {
                    surname=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  surname = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 5) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  year = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    year =date.format(cell.getDateCellValue());
                  } else {
                    year=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  year = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 6) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  month = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    month =date.format(cell.getDateCellValue());
                  } else {
                    month=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  month = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 7) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  day = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    day =date.format(cell.getDateCellValue());
                  } else {
                    day=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  day = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 8) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  member = Boolean.valueOf(cell.getStringCellValue());
                  break;
              }
            }
            if (column == 9){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  accountBeneficiary = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    accountBeneficiary =date.format(cell.getDateCellValue());
                  } else {
                    accountBeneficiary=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  accountBeneficiary = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 10){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  referenceCustomer = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    referenceCustomer =date.format(cell.getDateCellValue());
                  } else {
                    referenceCustomer=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  referenceCustomer = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 11){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  assignedOffice = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    assignedOffice =date.format(cell.getDateCellValue());
                  } else {
                    assignedOffice=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  assignedOffice = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 12){
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
            if (column == 13){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  street = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    street =date.format(cell.getDateCellValue());
                  } else {
                    street=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  street = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 14){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  city = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    city =date.format(cell.getDateCellValue());
                  } else {
                    city=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  city = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 15){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  region = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    region =date.format(cell.getDateCellValue());
                  } else {
                    region=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  region = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 16){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  postalCode = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    postalCode =date.format(cell.getDateCellValue());
                  } else {
                    postalCode=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  postalCode = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 17){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  countryCode = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    countryCode =date.format(cell.getDateCellValue());
                  } else {
                    countryCode=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  countryCode = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 18){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  country = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    country =date.format(cell.getDateCellValue());
                  } else {
                    country=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  country = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 19){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  typecontactDetail = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    typecontactDetail =date.format(cell.getDateCellValue());
                  } else {
                    typecontactDetail=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  typecontactDetail = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 20){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  group = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    group =date.format(cell.getDateCellValue());
                  } else {
                    group=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  group = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 21){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  value = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    value =date.format(cell.getDateCellValue());
                  } else {
                    value=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  value = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 22){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  preferenceLevel = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    preferenceLevel =date.format(cell.getDateCellValue());
                  } else {
                    preferenceLevel=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  preferenceLevel = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 23){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  validated = Boolean.valueOf(cell.getStringCellValue());
                  break;
              }
            }
            if (column == 24){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  currentState = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    currentState =date.format(cell.getDateCellValue());
                  } else {
                    currentState=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  currentState = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 25){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  applicationDate = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    applicationDate =date.format(cell.getDateCellValue());
                  } else {
                    applicationDate=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  applicationDate = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 26){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  catalogIdentifier = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    catalogIdentifier =date.format(cell.getDateCellValue());
                  } else {
                    catalogIdentifier=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  catalogIdentifier = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 27){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  fieldIdentifier = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    fieldIdentifier =date.format(cell.getDateCellValue());
                  } else {
                    fieldIdentifier=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  fieldIdentifier = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 28){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  value2 = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    value2 =date.format(cell.getDateCellValue());
                  } else {
                    value2=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  value2 = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 29){
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
            if (column == 30){
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
            if (column == 31){
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
            if (column == 32) {

              switch (cell.getCellType()) {

                case Cell.CELL_TYPE_STRING:
                  lastModifiedOn = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    lastModifiedOn = date.format(cell.getDateCellValue());
                  } else {
                    lastModifiedOn = Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  lastModifiedOn = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }

            column++;
          }
//deburging purpose
          logger.info(" "+identifier+" "+type+" "+givenName+ " "+middleName +" " +surname+" "+year+" "+ month+" "+day+" "+ member+" "+accountBeneficiary
                              +" "+referenceCustomer+" "+assignedOffice+" "+assignedEmployee+" "+street+" "+city+" "+region+" " +
                              " "+postalCode+" "+countryCode+" "+country+" "+typecontactDetail+" "+group+" " +
                              ""+value+" "+preferenceLevel+" "+validated+" "+currentState+" " +
                              ""+applicationDate+" "+catalogIdentifier+" "+fieldIdentifier+" "+value2+" " +
                              ""+createdBy+" "+createdOn+" "+lastModifiedBy+" "+lastModifiedOn

          );

          DateOfBirth dateOfBirth = new DateOfBirth();
          dateOfBirth.setYear(Integer.valueOf(year));
          dateOfBirth.setMonth(Integer.valueOf(month));
          dateOfBirth.setDay(Integer.valueOf(day));

          Address address = new Address();
          address.setStreet(String.valueOf(street));
          address.setCity(String.valueOf(city));
          address.setRegion(String.valueOf(region));
          address.setPostalCode(String.valueOf(postalCode));
          address.setCountryCode(String.valueOf(countryCode));
          address.setCountry(String.valueOf(country));

          ContactDetail contactDetail = new ContactDetail();
          contactDetail.setType(String.valueOf(typecontactDetail));
          contactDetail.setGroup(String.valueOf(group));
          contactDetail.setValue(String.valueOf(value));
          contactDetail.setPreferenceLevel(Integer.valueOf(preferenceLevel));
          contactDetail.setValidated(validated);

          List<ContactDetail> contactDetails = new ArrayList<>();
          contactDetails.add(contactDetail);

          Value value1=new Value();
          value1.setCatalogIdentifier(catalogIdentifier);
          value1.setFieldIdentifier(fieldIdentifier);
          value1.setValue(value2);

          List<Value> values = new ArrayList<>();
          values.add(value1);

          Customer customer= new Customer();
          customer.setIdentifier(String.valueOf(identifier));
          customer.setType(String.valueOf(type));
          customer.setGivenName(String.valueOf(givenName));
          customer.setMiddleName(String.valueOf(middleName));
          customer.setSurname(String.valueOf(surname));
          customer.setDateOfBirth(dateOfBirth);
          customer.setMember(member);
          customer.setAccountBeneficiary(accountBeneficiary);
          customer.setReferenceCustomer(referenceCustomer);
          customer.setAssignedOffice(assignedOffice);
          customer.setAssignedEmployee(assignedEmployee);
          customer.setAddress(address);
          customer.setContactDetails(contactDetails);
          customer.setCurrentState(currentState);
          customer.setApplicationDate(String.valueOf(applicationDate));

          //customer.setCustomValues(values);
          customer.setCreatedBy(String.valueOf(createdBy));
          customer.setCreatedOn(String.valueOf(createdOn));
          customer.setLastModifiedBy(String.valueOf(lastModifiedBy));
          customer.setLastModifiedOn(String.valueOf(lastModifiedOn));

          this.userManagement.authenticate();
          this.customerManager.createCustomer(customer);
        }

      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}

