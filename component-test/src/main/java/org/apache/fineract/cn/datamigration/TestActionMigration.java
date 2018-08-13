package org.apache.fineract.cn.datamigration;

import org.apache.fineract.cn.datamigration.util.ActionSheetGenerator;
import org.apache.fineract.cn.deposit.api.v1.EventConstants;
import org.junit.Assert;
import org.junit.Test;

public class TestActionMigration extends AbstractDataMigrationTest {
  @Test
  public void shouldMigrateLedger() throws Exception {
    ActionSheetGenerator.actionSheetDownload();

    //
    final boolean found = eventRecorder.wait(EventConstants.POST_PRODUCT_ACTION, ActionSheetGenerator.identifier);
    Assert.assertTrue(found);

  }
}
