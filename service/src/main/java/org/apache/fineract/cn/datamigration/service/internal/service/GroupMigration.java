package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.group.api.v1.client.GroupManager;
import org.apache.fineract.cn.group.api.v1.domain.Address;
import org.apache.fineract.cn.group.api.v1.domain.Group;
import org.apache.fineract.cn.group.api.v1.domain.GroupDefinition;
import org.apache.fineract.cn.office.api.v1.client.OrganizationManager;
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
import java.util.*;
import java.util.stream.IntStream;

@Service
public class GroupMigration {

  private final Logger logger;
  private final GroupManager groupManager;
  private final OrganizationManager organizationManager ;

  @Autowired
  public GroupMigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                        final GroupManager groupManager,
                        final OrganizationManager organizationManager) {
    super();
    this.logger = logger;
    this.groupManager = groupManager;
    this.organizationManager = organizationManager;
  }
  public void groupSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Group");


    List<GroupDefinition> groupDefinitions = this.groupManager.fetchGroupDefinitions();
    int sizeOfGroupDfinition= groupDefinitions.size();
    String [] groupDefinitionIdentifier=new String[sizeOfGroupDfinition];
    for (int i=0;i<sizeOfGroupDfinition;i++){
      groupDefinitionIdentifier[i]=groupDefinitions.get(i).getIdentifier();
    }



    Datavalidator.validatorWeekday(worksheet,"MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY",7);
    Datavalidator.validatorType(worksheet,"PENDING","ACTIVE","CLOSED",8);
   // Datavalidator.validatorString(worksheet,officeIdentifier,5);
    Datavalidator.validatorString(worksheet,groupDefinitionIdentifier,1);
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
    cell2.setCellValue("Group Definition Identifier");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Name");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Leaders");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5= rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Members");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6= rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Office");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7= rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Assigned Employee");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8= rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Weekday");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9= rowHeader.createCell(startColIndex+8);
    cell9.setCellValue("Status");
    cell9.setCellStyle(headerCellStyle);

    XSSFCell cell10= rowHeader.createCell(startColIndex+9);
    cell10.setCellValue("Street");
    cell10.setCellStyle(headerCellStyle);

    XSSFCell cell11= rowHeader.createCell(startColIndex+10);
    cell11.setCellValue("City");
    cell11.setCellStyle(headerCellStyle);

    XSSFCell cell12= rowHeader.createCell(startColIndex+11);
    cell12.setCellValue("Region");
    cell12.setCellStyle(headerCellStyle);

    XSSFCell cell13= rowHeader.createCell(startColIndex+12);
    cell13.setCellValue("Postal Code");
    cell13.setCellStyle(headerCellStyle);

    XSSFCell cell14= rowHeader.createCell(startColIndex+13);
    cell14.setCellValue("Country Code");
    cell14.setCellStyle(headerCellStyle);

    XSSFCell cell15= rowHeader.createCell(startColIndex+14);
    cell15.setCellValue("Country ");
    cell15.setCellStyle(headerCellStyle);


    IntStream.range(0, 17).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    response.setHeader("Content-Disposition", "inline; filename=Group.xlsx");
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

  public void groupSheetUpload(MultipartFile file){
    if (!file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
      throw new MultipartException("Only excel files accepted!");
    }
    try {
      XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
      Sheet firstSheet = workbook.getSheetAt(0);
      int rowCount = firstSheet.getLastRowNum() + 1;
      Row row;
      String identifier = null;
      String groupDefinitionIdentifier = null;
      String name = null;
      String leaders = null;
      String members = null;
      String office = null;
      String assignedEmployee = null;
      Integer weekdays=null;
      String status = null;
      String street = null;
      String city = null;
      String region = null;
      String postalCode = null;
      String countryCode = null;
      String country = null;
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
              identifier = String.valueOf(((Double)row.getCell(0).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(1) == null) {
          groupDefinitionIdentifier = null;
        } else {
          switch (row.getCell(1) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              groupDefinitionIdentifier = row.getCell(1).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              groupDefinitionIdentifier =   String.valueOf(((Double)row.getCell(1).getNumericCellValue()).intValue());
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
          leaders = null;
        } else {
          switch (row.getCell(3) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              leaders = row.getCell(3).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              leaders =  String.valueOf(((Double)row.getCell(3).getNumericCellValue()).intValue());
              break;

          }
        }

        if (row.getCell(4) == null) {
          members = null;
        } else {
          switch (row.getCell(4) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              members = row.getCell(4).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              members =  String.valueOf(((Double)row.getCell(4).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(5) == null) {
          office = null;
        } else {
          switch (row.getCell(5) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              office = row.getCell(5).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              office =  String.valueOf(((Double)row.getCell(5).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(6) == null) {
          assignedEmployee = null;
        } else {
          switch (row.getCell(6) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              assignedEmployee = row.getCell(6).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              assignedEmployee =  String.valueOf(((Double)row.getCell(6).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(7) == null) {
          weekdays = null;
        } else {
          switch (row.getCell(7).getStringCellValue()) {
            case "MONDAY":
              weekdays = 1;
              break;
            case "TUESDAY":
              weekdays = 2;
              break;
            case "WEDNESDAY":
              weekdays = 3;
              break;
            case "THURSDAY":
              weekdays = 4;
              break;
            case "FRIDAY":
              weekdays = 5;
              break;
            case "SATURDAY":
              weekdays = 6;
              break;
            case "SUNDAY":
              weekdays = 7;
              break;
          }
        }

        if (row.getCell(8) == null) {
          status = null;
        } else {
          switch (row.getCell(8) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              status = row.getCell(8).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              status =  String.valueOf(((Double)row.getCell(8).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(9) == null) {
          street = null;
        } else {
          switch (row.getCell(9) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              street = row.getCell(9).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              street =  String.valueOf(((Double)row.getCell(9).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(10) == null) {
          city = null;
        } else {
          switch (row.getCell(10) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              city = row.getCell(10).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              city =  String.valueOf(((Double)row.getCell(10).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(11) == null) {
          region = null;
        } else {
          switch (row.getCell(11) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              region = row.getCell(11).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              region =  String.valueOf(((Double)row.getCell(11).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(12) == null) {
          postalCode = null;
        } else {
          switch (row.getCell(12) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              postalCode = row.getCell(12).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              postalCode =  String.valueOf(((Double)row.getCell(12).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(13) == null) {
          countryCode = null;
        } else {
          switch (row.getCell(13) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              countryCode = row.getCell(13).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              countryCode =  String.valueOf(((Double)row.getCell(13).getNumericCellValue()).intValue());
              break;
          }
        }
        if (row.getCell(14) == null) {
          country = null;
        } else {
          switch (row.getCell(14) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              country = row.getCell(14).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              country =  String.valueOf(((Double)row.getCell(14).getNumericCellValue()).intValue());
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

        Set<String> leader = new HashSet<>();
        leader.add(leaders);

        Set<String> member = new HashSet<>();
        member.add(members);

        Group group= new Group();
        group.setIdentifier(String.valueOf(identifier));
        group.setGroupDefinitionIdentifier(String.valueOf(groupDefinitionIdentifier));
        group.setName(String.valueOf(name));
        group.setLeaders(leader);
        group.setMembers(member);
        group.setOffice(String.valueOf(office));
        group.setAssignedEmployee(String.valueOf(assignedEmployee));
        group.setWeekday(weekdays);
        group.setStatus(String.valueOf(status));
        group.setAddress(address);

        this.groupManager.createGroup(group);
      }

    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}

