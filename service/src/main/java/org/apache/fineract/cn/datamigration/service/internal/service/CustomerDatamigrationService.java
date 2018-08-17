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
import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.lang.DateOfBirth;
import org.apache.fineract.cn.office.api.v1.client.OrganizationManager;
import org.apache.fineract.cn.office.api.v1.domain.OfficePage;
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
import java.util.List;
import java.util.stream.IntStream;


@Service
public class CustomerDatamigrationService {
  private final Logger logger;
  private final CustomerManager customerManager;
  private final OrganizationManager organizationManager;


  @Autowired
  public CustomerDatamigrationService(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                                      final CustomerManager customerManager,
                                      final OrganizationManager organizationManager) {
     super();
     this.logger = logger;
     this.customerManager = customerManager;
     this.organizationManager = organizationManager;
  }

  public  void customersSheetDownload(HttpServletResponse response){
     XSSFWorkbook workbook = new XSSFWorkbook();
     XSSFSheet worksheet = workbook.createSheet("Customers");

    Datavalidator.validator(worksheet,"PERSON","BUSINESS",1);
    Datavalidator.validator(worksheet,"TRUE","FALSE",8);
    Datavalidator.validator(worksheet,"TRUE","FALSE",23);
    Datavalidator.validatorState(worksheet,"PENDING","ACTIVE","LOCKED","CLOSED",24);

    Datavalidator.validator(worksheet,"BUSINESS","PRIVATE",20);
    Datavalidator.validatorType(worksheet,"EMAIL","PHONE","MOBILE",19);
   // Datavalidator.validatorString(worksheet,officeIdentifier,11);

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

     IntStream.range(0, 26).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    response.setHeader("Content-Disposition", "inline; filename=Customer.xlsx");
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

  public void customersSheetUpload(MultipartFile file){
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

      SimpleDateFormat date=new SimpleDateFormat("yyyy-MM-dd");

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
          type = null;
        } else {
          switch (row.getCell(1) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              type = row.getCell(1).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                type =   String.valueOf(((Double)row.getCell(1).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(2) == null) {
          givenName = null;
        } else {
          switch (row.getCell(2) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              givenName = row.getCell(2).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                givenName =  String.valueOf(((Double)row.getCell(2).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(3) == null) {
          middleName = null;
        } else {
          switch (row.getCell(3) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              middleName = row.getCell(3).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                middleName =   String.valueOf(((Double)row.getCell(3).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(4) == null) {
          surname = null;
        } else {
          switch (row.getCell(4) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              surname = row.getCell(4).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                surname =   String.valueOf(((Double)row.getCell(4).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(5) == null) {
          year = null;
        } else {
          switch (row.getCell(5) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              year = row.getCell(5).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                year =  String.valueOf(((Double)row.getCell(5).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(6) == null) {
          month = null;
        } else {
          switch (row.getCell(6) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              month = row.getCell(6).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                month =   String.valueOf(((Double)row.getCell(6).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(7) == null) {
          day = null;
        } else {
          switch (row.getCell(7) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              day = String.valueOf(row.getCell(7).getStringCellValue());
              break;

            case Cell.CELL_TYPE_NUMERIC:
                day =  String.valueOf(((Double)row.getCell(7).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(8) == null) {
          member = false;
        } else {
          switch (row.getCell(8) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              member = Boolean.parseBoolean( String.valueOf(row.getCell(8).getStringCellValue()));
              break;

            case Cell.CELL_TYPE_NUMERIC:

              if(((Double)row.getCell(8).getNumericCellValue()).intValue()==0){
                member = Boolean.parseBoolean("false");
              }else{
                member = Boolean.parseBoolean("true");
              }
              break;
          }
        }

        if (row.getCell(9) == null) {
          accountBeneficiary = null;
        } else {
          switch (row.getCell(9) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              accountBeneficiary = row.getCell(9).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                accountBeneficiary =  String.valueOf(((Double)row.getCell(9).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(10) == null) {
          referenceCustomer = null;
        } else {
          switch (row.getCell(10) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              referenceCustomer = row.getCell(10).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                referenceCustomer =  String.valueOf(((Double)row.getCell(10).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(11) == null) {
          assignedOffice = null;
        } else {
          switch (row.getCell(11) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              assignedOffice = row.getCell(11).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                assignedOffice =  String.valueOf(((Double)row.getCell(11).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(12) == null) {
          assignedEmployee = null;
        } else {
          switch (row.getCell(12) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              assignedEmployee = row.getCell(12).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                assignedEmployee =  String.valueOf(((Double)row.getCell(12).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(13) == null) {
          street = null;
        } else {
          switch (row.getCell(13) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              street = row.getCell(13).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                street =  String.valueOf(((Double)row.getCell(13).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(14) == null) {
          city = null;
        } else {
          switch (row.getCell(14) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              city = row.getCell(14).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                city =  String.valueOf(((Double)row.getCell(14).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(15) == null) {
          region = null;
        } else {
          switch (row.getCell(15) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              region = row.getCell(15).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                region =  String.valueOf(((Double)row.getCell(15).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(16) == null) {
          postalCode = null;
        } else {
          switch (row.getCell(16) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              postalCode = row.getCell(16).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                postalCode =  String.valueOf(((Double)row.getCell(16).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(17) == null) {
          countryCode = null;
        } else {
          switch (row.getCell(17) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              countryCode = row.getCell(17).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                countryCode =  String.valueOf(((Double)row.getCell(17).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(18) == null) {
          country = null;
        } else {
          switch (row.getCell(18) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              country = row.getCell(18).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                country =  String.valueOf(((Double)row.getCell(18).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(19) == null) {
          typecontactDetail = null;
        } else {
          switch (row.getCell(19) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              typecontactDetail = row.getCell(19).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                typecontactDetail =  String.valueOf(((Double)row.getCell(19).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(20) == null) {
          group = null;
        } else {
          switch (row.getCell(20) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              group = row.getCell(20).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                group =  String.valueOf(((Double)row.getCell(20).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(21) == null) {
          value = null;
        } else {
          switch (row.getCell(21) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              value = row.getCell(21).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                value =  String.valueOf(((Double)row.getCell(21).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(22) == null) {
          preferenceLevel = null;
        } else {
          switch (row.getCell(22) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              preferenceLevel = row.getCell(22).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                preferenceLevel =  String.valueOf(((Double)row.getCell(22).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(23) == null) {
          validated = false;
        } else {
          switch (row.getCell(23) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              validated = Boolean.parseBoolean( String.valueOf(row.getCell(23).getStringCellValue()));
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if(((Double)row.getCell(23).getNumericCellValue()).intValue()==0){
                validated = Boolean.parseBoolean("false");
              }else{
                validated = Boolean.parseBoolean("true");
              }

              break;
          }
        }
        if (row.getCell(24) == null) {
          currentState = null;
        } else {
          switch (row.getCell(24) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              currentState = row.getCell(24).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
                currentState =  String.valueOf(((Double)row.getCell(24).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(25) == null) {
          applicationDate = null;
        } else {
          switch (row.getCell(25) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              applicationDate = row.getCell(25).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(25))) {

                applicationDate = date.format(row.getCell(25).getDateCellValue());
              }
              break;
          }
        }
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

          this.customerManager.createCustomer(customer);
        }

      } catch (IOException e) {
      e.printStackTrace();
    }
  }
}

