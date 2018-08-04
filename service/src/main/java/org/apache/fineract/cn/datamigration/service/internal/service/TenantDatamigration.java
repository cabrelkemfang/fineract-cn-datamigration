package org.apache.fineract.cn.datamigration.service.internal.service;

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.provisioner.api.v1.client.Provisioner;
import org.apache.fineract.cn.provisioner.api.v1.domain.CassandraConnectionInfo;
import org.apache.fineract.cn.provisioner.api.v1.domain.DatabaseConnectionInfo;
import org.apache.fineract.cn.provisioner.api.v1.domain.Tenant;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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
import java.math.BigDecimal;
import java.util.stream.IntStream;

@Service
public class TenantDatamigration {

  private final Logger logger;
  private final Provisioner provisioner;


  @Autowired
  public TenantDatamigration(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                             final Provisioner provisioner) {
    super();
    this.logger = logger;
    this.provisioner = provisioner;
  }

  public static void tenantSheetDownload(HttpServletResponse response){
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet worksheet = workbook.createSheet("Tenants");


    int startRowIndex = 0;
    int startColIndex = 0;

    Font font = worksheet.getWorkbook().createFont();
    XSSFCellStyle headerCellStyle = worksheet.getWorkbook().createCellStyle();

    headerCellStyle.setWrapText(true);
    headerCellStyle.setFont(font);
    XSSFRow rowHeader = worksheet.createRow((short) startRowIndex);
    rowHeader.setHeight((short) 500);


    XSSFCell cell1 = rowHeader.createCell(startColIndex+0);
    cell1.setCellValue("Tenant Identifier");
    cell1.setCellStyle(headerCellStyle);

    XSSFCell cell2 = rowHeader.createCell(startColIndex+1);
    cell2.setCellValue("Name");
    cell2.setCellStyle(headerCellStyle);

    XSSFCell cell3 = rowHeader.createCell(startColIndex+2);
    cell3.setCellValue("Description");
    cell3.setCellStyle(headerCellStyle);

    XSSFCell cell4 = rowHeader.createCell(startColIndex+3);
    cell4.setCellValue("Cluster Name ");
    cell4.setCellStyle(headerCellStyle);

    XSSFCell cell5 = rowHeader.createCell(startColIndex+4);
    cell5.setCellValue("Contact Points");
    cell5.setCellStyle(headerCellStyle);

    XSSFCell cell6 = rowHeader.createCell(startColIndex+5);
    cell6.setCellValue("Keyspace ");
    cell6.setCellStyle(headerCellStyle);

    XSSFCell cell7 = rowHeader.createCell(startColIndex+6);
    cell7.setCellValue("Replication Type ");
    cell7.setCellStyle(headerCellStyle);

    XSSFCell cell8 = rowHeader.createCell(startColIndex+7);
    cell8.setCellValue("Replicas");
    cell8.setCellStyle(headerCellStyle);

    XSSFCell cell9 = rowHeader.createCell(startColIndex+8);
    cell9.setCellValue(" Driver Class");
    cell9.setCellStyle(headerCellStyle);

    XSSFCell cell10= rowHeader.createCell(startColIndex+9);
    cell10.setCellValue("Database Name");
    cell10.setCellStyle(headerCellStyle);

    XSSFCell cell11= rowHeader.createCell(startColIndex+10);
    cell11.setCellValue("Host");
    cell11.setCellStyle(headerCellStyle);

    XSSFCell cell12= rowHeader.createCell(startColIndex+11);
    cell12.setCellValue("Port");
    cell12.setCellStyle(headerCellStyle);

    XSSFCell cell13= rowHeader.createCell(startColIndex+12);
    cell13.setCellValue("User");
    cell13.setCellStyle(headerCellStyle);

    XSSFCell cell14= rowHeader.createCell(startColIndex+13);
    cell14.setCellValue("Password");
    cell14.setCellStyle(headerCellStyle);


    IntStream.range(0, 14).forEach((columnIndex) -> worksheet.autoSizeColumn(columnIndex));

    response.setHeader("Content-Disposition", "inline; filename=Tenants.xlsx");
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

  public void tenantSheetUpload(MultipartFile file){
    if (!file.getContentType().equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
      throw new MultipartException("Only excel files accepted!");
    }
    try {
      XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream());
      Sheet firstSheet = workbook.getSheetAt(0);
      int rowCount = firstSheet.getLastRowNum() + 1;
      Row row;
      String identifier = null;
      String name = null;
      String description = null;
      String clusterName = null;
      String contactPoints = null;
      String keyspace = null;
      String replicationType = null;
      String replicas = null;
      String driverClass = null;
      String databaseName = null;
      String host = null;
      String port = null;
      String user = null;
      String password = null;

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
          name = null;
        } else {
          switch (row.getCell(1) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              name = row.getCell(1).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              name =   String.valueOf(((Double)row.getCell(1).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(2) == null) {
          description = null;
        } else {
          switch (row.getCell(2) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              description = row.getCell(2).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              description =  String.valueOf(((Double)row.getCell(2).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(3) == null) {
          clusterName = null;
        } else {
          switch (row.getCell(3).getCellType()){

            case Cell.CELL_TYPE_STRING:
              clusterName = row.getCell(3).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              clusterName = String.valueOf(((Double) row.getCell(3).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(4) == null) {
          contactPoints = null;
        } else {
          switch (row.getCell(4) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              contactPoints = row.getCell(4).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              contactPoints =   String.valueOf(((Double)row.getCell(4).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(5) == null) {
          keyspace = null;
        } else {
          switch (row.getCell(5) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              keyspace = row.getCell(5).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              keyspace =  String.valueOf(((Double)row.getCell(5).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(6) == null) {
          replicationType = null;
        } else {
          switch (row.getCell(6) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              replicationType = row.getCell(6).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              replicationType =   String.valueOf(((Double)row.getCell(6).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(7) == null) {
          replicas = null;
        } else {
          switch (row.getCell(7) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              replicas = String.valueOf(row.getCell(7).getStringCellValue());
              break;

            case Cell.CELL_TYPE_NUMERIC:
              replicas =  String.valueOf(((Double)row.getCell(7).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(8) == null) {
          driverClass = null;
        } else {
          switch (row.getCell(8) .getCellType()){

            case Cell.CELL_TYPE_STRING:
              driverClass = String.valueOf(row.getCell(8).getStringCellValue());
              break;

            case Cell.CELL_TYPE_NUMERIC:
              driverClass = String.valueOf(((Double) row.getCell(8).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(9) == null) {
          databaseName = null;
        } else {
          switch (row.getCell(9) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              databaseName = row.getCell(9).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              databaseName =  String.valueOf(((Double)row.getCell(9).getNumericCellValue()).intValue());
              break;
          }
        }

        if (row.getCell(10) == null) {
          host = null;
        } else {
          switch (row.getCell(10) .getCellType()) {

            case Cell.CELL_TYPE_STRING:
              host = row.getCell(10).getStringCellValue();
              break;

            case Cell.CELL_TYPE_NUMERIC:
              host =  String.valueOf(((Double)row.getCell(10).getNumericCellValue()).intValue());
              break;
          }
        }
          if (row.getCell(11) == null) {
            port = null;
          } else {
            switch (row.getCell(11) .getCellType()) {

              case Cell.CELL_TYPE_STRING:
                port = row.getCell(11).getStringCellValue();
                break;

              case Cell.CELL_TYPE_NUMERIC:
                port =  String.valueOf(((Double)row.getCell(11).getNumericCellValue()).intValue());
                break;
            }
          }
          if (row.getCell(12) == null) {
            user = null;
          } else {
            switch (row.getCell(12) .getCellType()) {

              case Cell.CELL_TYPE_STRING:
                user = row.getCell(12).getStringCellValue();
                break;

              case Cell.CELL_TYPE_NUMERIC:
                user =  String.valueOf(((Double)row.getCell(12).getNumericCellValue()).intValue());
                break;
            }
          }
          if (row.getCell(13) == null) {
            password = null;
          } else {
            switch (row.getCell(13) .getCellType()) {

              case Cell.CELL_TYPE_STRING:
                password = row.getCell(13).getStringCellValue();
                break;

              case Cell.CELL_TYPE_NUMERIC:
                password =  String.valueOf(((Double)row.getCell(13).getNumericCellValue()).intValue());
                break;
            }
          }

        CassandraConnectionInfo cassandraConnectionInfo =new CassandraConnectionInfo();
        cassandraConnectionInfo.setClusterName(String.valueOf(clusterName));
        cassandraConnectionInfo.setContactPoints(String.valueOf(contactPoints));
        cassandraConnectionInfo.setKeyspace(String.valueOf(keyspace));
        cassandraConnectionInfo.setReplicationType(String.valueOf(replicationType));
        cassandraConnectionInfo.setReplicas(String.valueOf(replicas));

        DatabaseConnectionInfo databaseConnectionInfo =new DatabaseConnectionInfo();
        databaseConnectionInfo.setDriverClass(String.valueOf(driverClass));
        databaseConnectionInfo.setDatabaseName(String.valueOf(databaseName));
        databaseConnectionInfo.setHost(String.valueOf(host));
        databaseConnectionInfo.setPort(String.valueOf(port));
        databaseConnectionInfo.setUser(String.valueOf(user));
        databaseConnectionInfo.setPassword(String.valueOf(password));

        Tenant tenant = new Tenant();
        tenant.setIdentifier(String.valueOf(identifier));
        tenant.setName(String.valueOf(name));
        tenant.setDescription(String.valueOf(description));
        tenant.setCassandraConnectionInfo(cassandraConnectionInfo);
        tenant.setDatabaseConnectionInfo(databaseConnectionInfo);

        this.provisioner.createTenant(tenant);

      }
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}
