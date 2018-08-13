package org.apache.fineract.cn.datamigration;

import org.apache.fineract.cn.datamigration.util.BranchOfficeSheetGenerator;
import org.apache.fineract.cn.office.api.v1.EventConstants;
import org.apache.fineract.cn.office.api.v1.domain.Office;
import org.apache.fineract.cn.office.api.v1.domain.OfficePage;
import org.junit.Assert;
import org.junit.Test;

public class TestBranchOfficeMigration extends AbstractDataMigrationTest {

  @Test
  public void shouldMigrateBranch() throws Exception {

    BranchOfficeSheetGenerator.branchSheetDownload();

    //

    this.eventRecorder.wait(EventConstants.OPERATION_POST_OFFICE, BranchOfficeSheetGenerator.parentIdentifier);

    final OfficePage officePage = this.organizationManager.getBranches(BranchOfficeSheetGenerator.parentIdentifier, 0, 10, null, null);
    Assert.assertEquals(Long.valueOf(1L), officePage.getTotalElements());

    final Office savedBranch = officePage.getOffices().get(0);
    Assert.assertEquals(BranchOfficeSheetGenerator.identifier, savedBranch.getIdentifier());
    Assert.assertEquals(BranchOfficeSheetGenerator.parentIdentifier, savedBranch.getParentIdentifier());
    Assert.assertEquals(BranchOfficeSheetGenerator.name, savedBranch.getName());
    Assert.assertEquals(BranchOfficeSheetGenerator.description, savedBranch.getDescription());
    Assert.assertEquals(BranchOfficeSheetGenerator.street, savedBranch.getAddress().getStreet());
    Assert.assertEquals(BranchOfficeSheetGenerator.city, savedBranch.getAddress().getCity());
    Assert.assertEquals(BranchOfficeSheetGenerator.region, savedBranch.getAddress().getRegion());
    Assert.assertEquals(BranchOfficeSheetGenerator.postalCode, savedBranch.getAddress().getPostalCode());
    Assert.assertEquals(BranchOfficeSheetGenerator.countryCode, savedBranch.getAddress().getCountryCode());
    Assert.assertEquals(BranchOfficeSheetGenerator.country, savedBranch.getAddress().getCountry());
  }
}
