package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.datamigration.service.connector.UserManagement;
import org.apache.fineract.cn.office.api.v1.client.OrganizationManager;
import org.apache.fineract.cn.office.api.v1.domain.ContactDetail;
import org.apache.fineract.cn.office.api.v1.domain.Employee;
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
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.stream.IntStream;

@Service
public class EmployeeDatamigration {
  private final Logger logger;
  private final OrganizationManager organizationManager;
  private final UserManagement userManagement;

  @Autowired
  public EmployeeDatamigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                                    final OrganizationManager organizationManager,
                                    final UserManagement userManagement) {
    super();
    this.logger = logger;
    this.organizationManager = organizationManager;
    this.userManagement = userManagement;
  }
  public void employeeSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Employees");

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
    cell2.setCellValue("Given Name");
    cell2.setCellStyle(headerCellStyle);


    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Middle Name");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Surname");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5 = rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Assigned Office ");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6 = rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Type ");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7 = rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Group ");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8 = rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Value ");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9= rowHeader.createCell(startColIndex+8);
    cell9.setCellValue("Preference Level");
    cell9.setCellStyle(headerCellStyle);

    IntStream.range(0, 9).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));
    response.setHeader("Content-Disposition", "inline; filename=Employees.xlsx");
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

  public void employeeSheetUpload(MultipartFile file){
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
          String givenName = null;
          String middleName = null;
          String surname = null;
          String assignedOffice = null;
          String type = null;
          String group = null;
          String value = null;
          String preferenceLevel = null;

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
            if (column == 2) {
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
            if (column == 3) {
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
            if (column == 4){
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
            if (column == 5) {
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
            if (column == 6){
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
            if (column == 7){
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
            if (column == 8){
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

            column++;
          }


          ContactDetail contactDetail = new ContactDetail();
          contactDetail.setType(String.valueOf(type));
          contactDetail.setGroup(String.valueOf(group));
          contactDetail.setValue(String.valueOf(value));
          contactDetail.setPreferenceLevel(Integer.valueOf(preferenceLevel));
          List<ContactDetail> contactDetails = new ArrayList<>();
          contactDetails.add(contactDetail);

          Employee employee= new Employee();
          employee.setIdentifier(String.valueOf(identifier));
          employee.setGivenName(String.valueOf(givenName));
          employee.setMiddleName(String.valueOf(middleName));
          employee.setSurname(String.valueOf(surname));
          employee.setAssignedOffice(assignedOffice);

          employee.setContactDetails(contactDetails);
          this.userManagement.authenticate();
          this.organizationManager.createEmployee(employee);
        }

      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }


}
