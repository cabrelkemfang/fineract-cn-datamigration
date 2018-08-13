package org.apache.fineract.cn.datamigration;

import org.apache.fineract.cn.accounting.api.v1.EventConstants;
import org.apache.fineract.cn.accounting.api.v1.domain.Ledger;
import org.apache.fineract.cn.datamigration.util.LedgerSheetGenerator;
import org.junit.Assert;
import org.junit.Test;

public class TestLedgerMigration extends AbstractDataMigrationTest{

  @Test
  public void shouldMigrateLedger() throws Exception {

    LedgerSheetGenerator.ledgerSheetDownload();

    //

    Assert.assertTrue(this.eventRecorder.wait(EventConstants.POST_LEDGER, LedgerSheetGenerator.identifier));

    final Ledger ledger = this.ledgerManager.findLedger(LedgerSheetGenerator.identifier);
    Assert.assertNotNull(ledger);

    Assert.assertEquals(LedgerSheetGenerator.type,ledger.getType());
    Assert.assertEquals(LedgerSheetGenerator.identifier,ledger.getIdentifier());
    Assert.assertEquals(LedgerSheetGenerator.name,ledger.getName());
    Assert.assertEquals(LedgerSheetGenerator.description,ledger.getDescription());
    Assert.assertEquals(LedgerSheetGenerator.showAccountsInChart,ledger.getShowAccountsInChart());
  }
}
