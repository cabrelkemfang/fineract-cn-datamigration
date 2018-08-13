package org.apache.fineract.cn.datamigration;

import org.apache.fineract.cn.accounting.api.v1.domain.Ledger;
import org.apache.fineract.cn.datamigration.util.LedgerSheetGenerator;
import org.apache.fineract.cn.datamigration.util.SubLedgerSheetGenerator;
import org.junit.Assert;
import org.junit.Test;

public class TestSubLedgerMigration extends AbstractDataMigrationTest {
  @Test
  public void shouldMigrateSubLedger() throws Exception {

    SubLedgerSheetGenerator.subLedgerSheetDownload();

    //

    final Ledger foundParentLedger = this.ledgerManager.findLedger(LedgerSheetGenerator.identifier);
    Assert.assertTrue(foundParentLedger.getSubLedgers().size() == 1);

    final Ledger foundSubLedger = foundParentLedger.getSubLedgers().get(0);

    Assert.assertEquals(SubLedgerSheetGenerator.type,foundSubLedger.getType());
    Assert.assertEquals(SubLedgerSheetGenerator.identifier,foundSubLedger.getIdentifier());
    Assert.assertEquals(SubLedgerSheetGenerator.name,foundSubLedger.getName());
    Assert.assertEquals(SubLedgerSheetGenerator.description,foundSubLedger.getDescription());
    Assert.assertEquals(SubLedgerSheetGenerator.showAccountsInChart,foundSubLedger.getShowAccountsInChart());
  }
}
