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
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.slf4j.Logger;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartException;
import org.springframework.web.multipart.MultipartFile;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.Collections;
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
     cell1.setCellValue("identifier");
     cell1.setCellStyle(headerCellStyle);

     XSSFCell cell2 = rowHeader.createCell(startColIndex+1);
     cell2.setCellValue("type");
     cell2.setCellStyle(headerCellStyle);


     XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
     cell3.setCellValue("givenName");
     cell3.setCellStyle(headerCellStyle);

     XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
     cell4.setCellValue("middleName");
     cell4.setCellStyle(headerCellStyle);

     XSSFCell cell5 = rowHeader.createCell(startColIndex+4);
     cell5.setCellValue("surname ");
     cell5.setCellStyle(headerCellStyle);

     //dateOfBirth
     XSSFCell cell6 = rowHeader.createCell(startColIndex+5);
     cell6.setCellValue("year ");
     cell6.setCellStyle(headerCellStyle);

     XSSFCell cell7 = rowHeader.createCell(startColIndex+6);
     cell7.setCellValue("month ");
     cell7.setCellStyle(headerCellStyle);

     XSSFCell cell8 = rowHeader.createCell(startColIndex+7);
     cell8.setCellValue("day ");
     cell8.setCellStyle(headerCellStyle);

     XSSFCell cell9= rowHeader.createCell(startColIndex+8);
     cell9.setCellValue("member");
     cell9.setCellStyle(headerCellStyle);

     XSSFCell cell10= rowHeader.createCell(startColIndex+9);
     cell10.setCellValue("accountBeneficiary");
     cell10.setCellStyle(headerCellStyle);

     XSSFCell cell11= rowHeader.createCell(startColIndex+10);
     cell11.setCellValue("referenceCustomer");
     cell11.setCellStyle(headerCellStyle);

     XSSFCell cell12= rowHeader.createCell(startColIndex+11);
     cell12.setCellValue("assignedOffice");
     cell12.setCellStyle(headerCellStyle);

     XSSFCell cell13= rowHeader.createCell(startColIndex+12);
     cell13.setCellValue("assignedEmployee");
     cell13.setCellStyle(headerCellStyle);

     //address
     XSSFCell cell14= rowHeader.createCell(startColIndex+13);
     cell14.setCellValue("street");
     cell14.setCellStyle(headerCellStyle);

     XSSFCell cell15= rowHeader.createCell(startColIndex+14);
     cell15.setCellValue("city");
     cell15.setCellStyle(headerCellStyle);

     XSSFCell cell16= rowHeader.createCell(startColIndex+15);
     cell16.setCellValue("region");
     cell16.setCellStyle(headerCellStyle);

     XSSFCell cell17= rowHeader.createCell(startColIndex+16);
     cell17.setCellValue("postalCode");
     cell17.setCellStyle(headerCellStyle);

     XSSFCell cell18= rowHeader.createCell(startColIndex+17);
     cell18.setCellValue("countryCode");
     cell18.setCellStyle(headerCellStyle);

     XSSFCell cell19= rowHeader.createCell(startColIndex+18);
     cell19.setCellValue("country");
     cell19.setCellStyle(headerCellStyle);

     //contactDetail
     XSSFCell cell20= rowHeader.createCell(startColIndex+19);
     cell20.setCellValue("type");
     cell20.setCellStyle(headerCellStyle);

     XSSFCell cell21= rowHeader.createCell(startColIndex+20);
     cell21.setCellValue("group");
     cell21.setCellStyle(headerCellStyle);

     XSSFCell cell22= rowHeader.createCell(startColIndex+21);
     cell22.setCellValue("value");
     cell22.setCellStyle(headerCellStyle);

     XSSFCell cell23= rowHeader.createCell(startColIndex+22);
     cell23.setCellValue("preferenceLevel");
     cell23.setCellStyle(headerCellStyle);

     XSSFCell cell24= rowHeader.createCell(startColIndex+23);
     cell24.setCellValue("validated");
     cell24.setCellStyle(headerCellStyle);


     XSSFCell cell25= rowHeader.createCell(startColIndex+24);
     cell25.setCellValue("currentState");
     cell25.setCellStyle(headerCellStyle);

     XSSFCell cell26= rowHeader.createCell(startColIndex+25);
     cell26.setCellValue("applicationDate");
     cell26.setCellStyle(headerCellStyle);

     //value
     XSSFCell cell27= rowHeader.createCell(startColIndex+26);
     cell27.setCellValue("catalogIdentifier");
     cell27.setCellStyle(headerCellStyle);

     XSSFCell cell28= rowHeader.createCell(startColIndex+27);
     cell28.setCellValue("fieldIdentifier");
     cell28.setCellStyle(headerCellStyle);

     XSSFCell cell29= rowHeader.createCell(startColIndex+28);
     cell29.setCellValue("value");
     cell29.setCellStyle(headerCellStyle);

     XSSFCell cell30= rowHeader.createCell(startColIndex+29);
     cell30.setCellValue("createdBy");
     cell30.setCellStyle(headerCellStyle);

     XSSFCell cell31= rowHeader.createCell(startColIndex+30);
     cell31.setCellValue("createdOn");
     cell31.setCellStyle(headerCellStyle);

     XSSFCell cell32= rowHeader.createCell(startColIndex+31);
     cell32.setCellValue("lastModifiedBy");
     cell32.setCellStyle(headerCellStyle);

     XSSFCell cell33= rowHeader.createCell(startColIndex+32);
     cell33.setCellValue("lastModifiedOn");
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
        } else {
        try{
           XSSFWorkbook customers = new XSSFWorkbook(file.getInputStream());
           XSSFSheet worksheet = customers.getSheetAt(0);
           XSSFRow entry;
           Integer noOfEntries=1;

           //getLastRowNum and getPhysicalNumberOfRows showing false values sometimes.
           while(worksheet.getRow(noOfEntries)!=null){
              noOfEntries++;
           }
           // logger.info(noOfEntries.toString());
          for(int rowIndex=1;rowIndex<noOfEntries;rowIndex++){
             entry=worksheet.getRow(rowIndex);

             String identifier =entry.getCell(1).getStringCellValue();
             String type =entry.getCell(2).getStringCellValue();
             String givenName =entry.getCell(3).getStringCellValue();
             String middleName =entry.getCell(4).getStringCellValue();
             String surname =entry.getCell(5).getStringCellValue();
             //Birtday
             Integer year =((Double)entry.getCell(6).getNumericCellValue()).intValue();
             Integer month = ((Double)entry.getCell(7).getNumericCellValue()).intValue();
             Integer day =((Double)entry.getCell(8).getNumericCellValue()).intValue();

             Boolean member =entry.getCell(9).getBooleanCellValue();
             String accountBeneficiary =entry.getCell(10).getStringCellValue();
             String referenceCustomer =entry.getCell(11).getStringCellValue();
             String assignedOffice =entry.getCell(12).getStringCellValue();
             String assignedEmployee =entry.getCell(13).getStringCellValue();
             //address
             String street =entry.getCell(14).getStringCellValue();
             String city =entry.getCell(15).getStringCellValue();
             String region =entry.getCell(16).getStringCellValue();
             String postalCode =entry.getCell(17).getStringCellValue();
             String countryCode =entry.getCell(18).getStringCellValue();
             String country =entry.getCell(19).getStringCellValue();
             //contactDetail
             String typecontactDetail =entry.getCell(20).getStringCellValue();
             String group =entry.getCell(21).getStringCellValue();
             String value =entry.getCell(22).getStringCellValue();
             Integer preferenceLevel =((Double)entry.getCell(23).getNumericCellValue()).intValue();
             Boolean validated =entry.getCell(24).getBooleanCellValue();

             String currentState =entry.getCell(25).getStringCellValue();
             String applicationDate =entry.getCell(26).getStringCellValue();
             //value
             String catalogIdentifier =entry.getCell(27).getStringCellValue();
             String fieldIdentifier =entry.getCell(28).getStringCellValue();
             String value2 =entry.getCell(29).getStringCellValue();

             String createdBy =entry.getCell(30).getStringCellValue();
             String createdOn =entry.getCell(31).getStringCellValue();
             String lastModifiedBy =entry.getCell(32).getStringCellValue();
             String lastModifiedOn =entry.getCell(33).getStringCellValue();

             DateOfBirth dateOfBirth = new DateOfBirth();
             dateOfBirth.setYear(year);
             dateOfBirth.setMonth(month);
             dateOfBirth.setDay(day);

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
             contactDetail.setPreferenceLevel(preferenceLevel);
             contactDetail.setValidated(validated);

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
             customer.setMember(member);
             customer.setAccountBeneficiary(accountBeneficiary);
             customer.setReferenceCustomer(referenceCustomer);
             customer.setAssignedOffice(assignedOffice);
             customer.setAssignedEmployee(assignedEmployee);
             customer.setAddress(address);
             customer.setContactDetails(Collections.singletonList(contactDetail));
             customer.setCurrentState(currentState);
             customer.setApplicationDate(applicationDate);
             customer.setCustomValues((List<Value>) value1);
             customer.setCreatedBy(createdBy);
             customer.setCreatedOn(createdOn);
             customer.setLastModifiedBy(lastModifiedBy);
             customer.setLastModifiedOn(lastModifiedOn);

             this.userManagement.authenticate();
             this.customerManager.createCustomer(customer);
             }
        }catch(Exception e){
           System.out.println(e.getMessage()+" "+e.getCause());
           throw new MultipartException("Constraints Violated");
        }
     }
  }
}
