package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.datamigration.service.connector.UserManagement;
import org.apache.fineract.cn.group.api.v1.client.GroupManager;
import org.apache.fineract.cn.group.api.v1.domain.Address;
import org.apache.fineract.cn.group.api.v1.domain.Group;
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
import java.util.*;
import java.util.stream.IntStream;

@Service
public class GroupDatamigration {

  private final Logger logger;
  private final GroupManager groupManager;
  private final UserManagement userManagement;

  @Autowired
  public GroupDatamigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                                    final GroupManager groupManager,
                                    final UserManagement userManagement) {
    super();
    this.logger = logger;
    this.groupManager = groupManager;
    this.userManagement = userManagement;
  }
  public void groupSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Group");

    Datavalidator.validatorType(worksheet,"PENDING","ACTIVE","CLOSED",8);
    Datavalidator.validatorWeekday(worksheet,"MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY",7);

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

    XSSFCell cell16= rowHeader.createCell(startColIndex+15);
    cell16.setCellValue("Created On");
    cell16.setCellStyle(headerCellStyle);

    XSSFCell cell17= rowHeader.createCell(startColIndex+16);
    cell17.setCellValue("Created By");
    cell17.setCellStyle(headerCellStyle);

    XSSFCell cell18= rowHeader.createCell(startColIndex+17);
    cell18.setCellValue("Last Modified On");
    cell18.setCellStyle(headerCellStyle);

    XSSFCell cell19= rowHeader.createCell(startColIndex+18);
    cell19.setCellValue("Last Modified By");
    cell19.setCellStyle(headerCellStyle);

    IntStream.range(0, 19).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
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
      int offset = 0;
      int currentPosition = 0;

      for (Row nextRow : firstSheet) {
        int column = 0;
        Iterator<Cell> cellIterator = nextRow.cellIterator();
        if ((currentPosition++ > offset)) {
          String identifier = null;
          String groupDefinitionIdentifier = null;
          String name = null;
          String leaders = null;
          String members = null;
          String office = null;
          String assignedEmployee = null;
          String weekday = null;
          Integer weekdays=null;
          String status = null;
          String street = null;
          String city = null;
          String region = null;
          String postalCode = null;
          String countryCode = null;
          String country = null;
          String createdOn = null;
          String createdBy = null;
          String lastModifiedOn = null;
          String lastModifiedBy = null;
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
                  groupDefinitionIdentifier = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    groupDefinitionIdentifier =date.format(cell.getDateCellValue());
                  } else {
                    groupDefinitionIdentifier=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  groupDefinitionIdentifier = String.valueOf(cell.getBooleanCellValue());
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

                    name =date.format(cell.getDateCellValue());
                  } else {
                    name=Integer.toString((int) cell.getNumericCellValue());
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
                  leaders = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    leaders =date.format(cell.getDateCellValue());
                  } else {
                    leaders =Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  leaders = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 4) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  members = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    members =date.format(cell.getDateCellValue());
                  } else {
                    members=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  members = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 5) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  office = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {
                    office =date.format(cell.getDateCellValue());
                  } else {
                    office=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  office = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 6) {
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
            if (column == 7) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  weekday = cell.getStringCellValue();
                  break;
              }
              switch (weekday) {
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
            if (column == 8) {
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  status = cell.getStringCellValue();
                  break;

                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    status =date.format(cell.getDateCellValue());
                  } else {
                    status=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  status = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }
            if (column == 9) {
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
            if (column == 10){
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
            if (column == 11){
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
            if (column == 12){
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
            if (column == 13){
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
            if (column == 14){
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
            if (column == 15){
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
            if (column == 16){
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
            if (column == 17){
              switch (cell.getCellType()) {
                case Cell.CELL_TYPE_STRING:
                  lastModifiedOn = cell.getStringCellValue();
                  break;
                case Cell.CELL_TYPE_NUMERIC:
                  if (DateUtil.isCellDateFormatted(cell)) {

                    lastModifiedOn =date.format(cell.getDateCellValue());
                  } else {
                    lastModifiedOn=Integer.toString((int) cell.getNumericCellValue());
                  }
                  break;

                case Cell.CELL_TYPE_BOOLEAN:
                  lastModifiedOn = String.valueOf(cell.getBooleanCellValue());
                  break;
              }
            }

            if (column == 18){
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
            column++;
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
          group.setStatus(status);
          group.setAddress(address);
          group.setCreatedOn(String.valueOf(createdOn));
          group.setCreatedBy(String.valueOf(createdBy));
          group.setLastModifiedOn(String.valueOf(lastModifiedOn));
          group.setLastModifiedBy(String.valueOf(lastModifiedBy));

          this.userManagement.authenticate();
          this.groupManager.createGroup(group);
        }

      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }




}
