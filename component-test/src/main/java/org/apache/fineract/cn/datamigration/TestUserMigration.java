package org.apache.fineract.cn.datamigration;

import org.apache.fineract.cn.datamigration.util.UserSheetSheetGenerator;
import org.apache.fineract.cn.identity.api.v1.domain.User;
import org.apache.fineract.cn.identity.api.v1.events.EventConstants;
import org.junit.Assert;
import org.junit.Test;

public class TestUserMigration extends AbstractDataMigrationTest {
  @Test
  public void shouldMigrateUser() throws Exception {

    UserSheetSheetGenerator.userSheetDownload();

    //
    final boolean found = eventRecorder.wait(EventConstants.OPERATION_POST_USER, UserSheetSheetGenerator.identifier);
    Assert.assertTrue(found);

    final User user = identityManager.getUser(UserSheetSheetGenerator.identifier);
    Assert.assertNotNull(user);
    Assert.assertEquals(UserSheetSheetGenerator.identifier,user.getIdentifier());
    Assert.assertEquals(UserSheetSheetGenerator.role,user.getRole());

  }
}
