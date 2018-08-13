package org.apache.fineract.cn.datamigration;


import org.apache.commons.io.IOUtils;
import org.apache.fineract.cn.datamigration.util.OfficeSheetGenerator;
import org.apache.fineract.cn.office.api.v1.EventConstants;
import org.apache.fineract.cn.office.api.v1.domain.Office;
import org.apache.tomcat.util.http.fileupload.disk.DiskFileItem;
import org.junit.Assert;
import org.junit.Test;
import org.springframework.http.MediaType;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

import java.io.File;
import java.io.FileInputStream;


public class TestOfficeMigration extends AbstractDataMigrationTest {

  @Test
  public void shouldMigrateOfficce() throws Exception {

    logger.info("genarating office excel sheet test");
    OfficeSheetGenerator.CreateofficeSheet();

    //final MockMultipartFile multipartFile = new MockMultipartFile("office.xlsx", new FileInputStream(new File("/home/le/Desktop/office.xlsx")));
    final MockMultipartFile file = new MockMultipartFile("portrait", "test.png", MediaType.IMAGE_PNG_VALUE, ("idon'care").getBytes());
    //final MockMultipartFile multipartFile = new MockMultipartFile("office", "office.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml", "/home/le/Desktop/office.xlsx".getBytes());
   // this.datamigrationManager.officeSheetUpload((MultipartFile) new FileInputStream(new File("/home/le/Desktop/office.xlsx")));

/*
    File file = new File("/home/le/Desktop/office.xlsx");
    FileInputStream input = new FileInputStream(file);
    MultipartFile multipartFile = new MockMultipartFile("office", file.getName(), "application/vnd.openxmlformats-officedocument.spreadsheetml", IOUtils.toByteArray(input));
*/

   /* File file = new File("/home/le/Desktop/office.xlsx");
    DiskFileItem fileItem = new DiskFileItem("file", "application/vnd.openxmlformats-officedocument.spreadsheetml", false, file.getName(), (int) file.length() , file.getParentFile());
    fileItem.getOutputStream();
    MultipartFile multipartFile = new CommonsMultipartFile(fileItem);*/

    this.datamigrationManager.officeSheetUpload(file);

    this.eventRecorder.wait(EventConstants.OPERATION_POST_OFFICE, OfficeSheetGenerator.identifier);
    final Office savedOffice = this.organizationManager.findOfficeByIdentifier(OfficeSheetGenerator.identifier);
    Assert.assertNotNull(savedOffice);

    Assert.assertEquals(OfficeSheetGenerator.identifier,savedOffice.getIdentifier());
    Assert.assertEquals(OfficeSheetGenerator.name,savedOffice.getName());
    Assert.assertEquals(OfficeSheetGenerator.description,savedOffice.getDescription());
    Assert.assertEquals(OfficeSheetGenerator.street,savedOffice.getAddress().getStreet());
    Assert.assertEquals(OfficeSheetGenerator.city,savedOffice.getAddress().getCity());
    Assert.assertEquals(OfficeSheetGenerator.region,savedOffice.getAddress().getRegion());
    Assert.assertEquals(OfficeSheetGenerator.postalCode,savedOffice.getAddress().getPostalCode());
    Assert.assertEquals(OfficeSheetGenerator.countryCode,savedOffice.getAddress().getCountryCode());
    Assert.assertEquals(OfficeSheetGenerator.country,savedOffice.getAddress().getCountry());
  }
}
