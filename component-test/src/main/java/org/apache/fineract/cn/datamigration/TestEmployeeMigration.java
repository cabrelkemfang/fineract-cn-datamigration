package org.apache.fineract.cn.datamigration;

import org.apache.fineract.cn.datamigration.util.EmployeeSheetGenerator;
import org.apache.fineract.cn.office.api.v1.EventConstants;
import org.apache.fineract.cn.office.api.v1.domain.ContactDetail;
import org.apache.fineract.cn.office.api.v1.domain.Employee;
import org.junit.Assert;
import org.junit.Test;

public class TestEmployeeMigration extends AbstractDataMigrationTest {

  @Test
  public void shouldMigrateEmployee() throws Exception {

    EmployeeSheetGenerator.employeeSheetDownload();

    //

    final boolean found = this.eventRecorder.wait(EventConstants.OPERATION_POST_EMPLOYEE,
            EmployeeSheetGenerator.identifier);
    Assert.assertTrue(found);

    final Employee savedEmployee = this.organizationManager.findEmployee(EmployeeSheetGenerator.identifier);

    Assert.assertNotNull(savedEmployee);
    Assert.assertEquals(EmployeeSheetGenerator.identifier, savedEmployee.getIdentifier());
    Assert.assertEquals(EmployeeSheetGenerator.givenName, savedEmployee.getGivenName());
    Assert.assertEquals(EmployeeSheetGenerator.middleName, savedEmployee.getMiddleName());
    Assert.assertEquals(EmployeeSheetGenerator.surname, savedEmployee.getSurname());
    Assert.assertEquals(EmployeeSheetGenerator.assignedOffice, savedEmployee.getAssignedOffice());
    ContactDetail contactDetail= new ContactDetail();
    contactDetail=savedEmployee.getContactDetails().get(0);
    Assert.assertEquals(EmployeeSheetGenerator.type, contactDetail.getType());
    Assert.assertEquals(EmployeeSheetGenerator.group, contactDetail.getGroup());
    Assert.assertEquals(EmployeeSheetGenerator.value, contactDetail.getValue());
    Assert.assertEquals(EmployeeSheetGenerator.preferenceLevel, contactDetail.getPreferenceLevel());

  }
}
