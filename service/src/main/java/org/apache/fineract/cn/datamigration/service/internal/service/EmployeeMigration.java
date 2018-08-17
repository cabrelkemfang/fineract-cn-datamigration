package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.office.api.v1.client.OrganizationManager;
import org.apache.fineract.cn.office.api.v1.domain.ContactDetail;
import org.apache.fineract.cn.office.api.v1.domain.Employee;
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
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

@Service
public class EmployeeMigration {
  private final Logger logger;
  private final OrganizationManager organizationManager;

  @Autowired
  public EmployeeMigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                           final OrganizationManager organizationManager) {
    super();
    this.logger = logger;
    this.organizationManager = organizationManager;
  }


  public void employeeSheetDownload(HttpServletResponse response){

    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Employees");

    Datavalidator.validator(worksheet,"BUSINESS","PRIVATE",6);
    Datavalidator.validatorType(worksheet,"EMAIL","PHONE","MOBILE",5);
    //Datavalidator.validatorString(worksheet,officeIdentifier,4);

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
      outputStream.close();
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
      int rowCount = firstSheet.getLastRowNum() + 1;
      Row row;
      String identifier =null ;
      String givenName  =null;
      String middleName  =null ;
      String surname   =null;
      String assignedOffice   =null;
      String type   =null;
      String group   =null;
      String value   =null;
      String preferenceLevel   =null;

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
          givenName = null;
        } else {
          switch (row.getCell(1) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              givenName = row.getCell(1).getStringCellValue();
              break;
            case Cell.CELL_TYPE_NUMERIC:
                givenName =  String.valueOf(((Double)row.getCell(1).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(2) == null) {
          middleName = null;
        } else {
          switch (row.getCell(2) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              middleName = row.getCell(2).getStringCellValue();
              break;
            case Cell.CELL_TYPE_NUMERIC:
                middleName =   String.valueOf(((Double)row.getCell(2).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(3) == null) {
          surname = null;
        } else {
          switch (row.getCell(3) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              surname = row.getCell(3).getStringCellValue();
              break;
            case Cell.CELL_TYPE_NUMERIC:
                surname =   String.valueOf(((Double)row.getCell(3).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(4) == null) {
          assignedOffice = null;
        } else {
          switch (row.getCell(4) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              assignedOffice = row.getCell(4).getStringCellValue();
              break;
            case Cell.CELL_TYPE_NUMERIC:
                assignedOffice =  String.valueOf(((Double)row.getCell(4).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(5) == null) {
          type = null;
        } else {
          switch (row.getCell(5) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              type = row.getCell(5).getStringCellValue();
              break;
            case Cell.CELL_TYPE_NUMERIC:
                type =   String.valueOf(((Double)row.getCell(5).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(6) == null) {
          group = null;
        } else {
          switch (row.getCell(6) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              group = String.valueOf(row.getCell(6).getStringCellValue());
              break;
            case Cell.CELL_TYPE_NUMERIC:
              if (DateUtil.isCellDateFormatted(row.getCell(6))) {
                group =  String.valueOf(row.getCell(6).getStringCellValue());
              } else {
                group =  String.valueOf(((Double)row.getCell(6).getNumericCellValue()).intValue());
              }
              break;
          }
        }

        if (row.getCell(7) == null) {
          value = null;
        } else {
          switch (row.getCell(7) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              value = String.valueOf(row.getCell(7).getStringCellValue());
              break;
            case Cell.CELL_TYPE_NUMERIC:
                value =  String.valueOf(((Double)row.getCell(7).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(8) == null) {
          preferenceLevel = null;
        } else {
          switch (row.getCell(8) .getCellType()) {
            case Cell.CELL_TYPE_STRING:
              preferenceLevel = row.getCell(8).getStringCellValue();
              break;
            case Cell.CELL_TYPE_NUMERIC:
                preferenceLevel =  String.valueOf(((Double)row.getCell(8).getNumericCellValue()).intValue());
              break;
          }
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

          this.organizationManager.createEmployee(employee);
        }

    } catch (IOException e) {
      e.printStackTrace();
    }
  }


}
