package org.apache.fineract.cn.datamigration;

import org.apache.fineract.cn.datamigration.util.GroupDefinitionGenerator;
import org.apache.fineract.cn.group.api.v1.EventConstants;
import org.apache.fineract.cn.group.api.v1.domain.GroupDefinition;
import org.junit.Assert;
import org.junit.Test;

public class TestGroupDefintionMigration extends AbstractDataMigrationTest{
  @Test
  public void shouldMigrateGroupDefinition() throws Exception {

    GroupDefinitionGenerator.groupDefinitionSheetDownload();

    //
    final boolean found =eventRecorder.wait(EventConstants.POST_GROUP_DEFINITION, GroupDefinitionGenerator.identifier);
    Assert.assertTrue(found);

    final GroupDefinition fetchedGroupDefinition = groupManager.findGroupDefinition(GroupDefinitionGenerator.identifier);
    Assert.assertEquals(GroupDefinitionGenerator.identifier,fetchedGroupDefinition.getIdentifier());
    Assert.assertEquals(GroupDefinitionGenerator.description,fetchedGroupDefinition.getDescription());
    Assert.assertEquals(GroupDefinitionGenerator.minimalSize,fetchedGroupDefinition.getMinimalSize());
    Assert.assertEquals(GroupDefinitionGenerator.maximalSize,fetchedGroupDefinition.getMaximalSize());
    Assert.assertEquals(GroupDefinitionGenerator.numberOfMeetings,fetchedGroupDefinition.getCycle().getNumberOfMeetings());
    Assert.assertEquals(GroupDefinitionGenerator.adjustment,fetchedGroupDefinition.getCycle().getAdjustment());
    Assert.assertEquals(GroupDefinitionGenerator.frequency,fetchedGroupDefinition.getCycle().getFrequency());
  }

}
