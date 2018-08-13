package org.apache.fineract.cn.datamigration;

import org.apache.fineract.cn.datamigration.util.GroupSheetGenerator;
import org.apache.fineract.cn.group.api.v1.EventConstants;
import org.apache.fineract.cn.group.api.v1.domain.Group;
import org.junit.Assert;
import org.junit.Test;

public class TestGroupMigration extends AbstractDataMigrationTest {

  @Test
  public void shouldMigrateGroup() throws Exception {

    GroupSheetGenerator.groupSheetDownload();

    //

    final boolean found =eventRecorder.wait(EventConstants.POST_GROUP, GroupSheetGenerator.identifier);
    Assert.assertTrue(found);

    final Group fetchedGroup = groupManager.findGroup(GroupSheetGenerator.identifier);
    Assert.assertEquals(GroupSheetGenerator.identifier,fetchedGroup.getIdentifier());
    Assert.assertEquals(GroupSheetGenerator.groupDefinitionIdentifier,fetchedGroup.getGroupDefinitionIdentifier());
    Assert.assertEquals(GroupSheetGenerator.name,fetchedGroup.getName());
    Assert.assertEquals(GroupSheetGenerator.leaders,fetchedGroup.getLeaders());
    Assert.assertEquals(GroupSheetGenerator.members,fetchedGroup.getMembers());
    Assert.assertEquals(GroupSheetGenerator.office,fetchedGroup.getOffice());
    Assert.assertEquals(GroupSheetGenerator.assignedEmployee,fetchedGroup.getAssignedEmployee());
    Assert.assertEquals(GroupSheetGenerator.weekdays,fetchedGroup.getWeekday());
    Assert.assertEquals(GroupSheetGenerator.status,fetchedGroup.getStatus());
    Assert.assertEquals(GroupSheetGenerator.street,fetchedGroup.getAddress().getStreet());
    Assert.assertEquals(GroupSheetGenerator.city,fetchedGroup.getAddress().getCity());
    Assert.assertEquals(GroupSheetGenerator.region,fetchedGroup.getAddress().getRegion());
    Assert.assertEquals(GroupSheetGenerator.postalCode,fetchedGroup.getAddress().getPostalCode());
    Assert.assertEquals(GroupSheetGenerator.countryCode,fetchedGroup.getAddress().getCountryCode());
    Assert.assertEquals(GroupSheetGenerator.country,fetchedGroup.getAddress().getCountry());
  }
}
