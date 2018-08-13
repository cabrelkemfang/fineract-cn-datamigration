package org.apache.fineract.cn.datamigration;

import org.apache.fineract.cn.accounting.api.v1.EventConstants;
import org.apache.fineract.cn.accounting.api.v1.domain.Account;
import org.apache.fineract.cn.datamigration.util.LedgerAccountSheetGenerator;
import org.junit.Assert;
import org.junit.Test;

public class TestLedgerAccountSheetGenerator extends AbstractDataMigrationTest {

  @Test
  public void shouldMigrateLedgerAccount() throws Exception {

    LedgerAccountSheetGenerator.accountSheetDownload();

    //
    this.eventRecorder.wait(EventConstants.POST_ACCOUNT, LedgerAccountSheetGenerator.identifier);

    final Account savedAccount = this.ledgerManager.findAccount(LedgerAccountSheetGenerator.identifier);
    Assert.assertNotNull(savedAccount);

    Assert.assertEquals(LedgerAccountSheetGenerator.type,savedAccount.getType());
    Assert.assertEquals(LedgerAccountSheetGenerator.identifier,savedAccount.getIdentifier());
    Assert.assertEquals(LedgerAccountSheetGenerator.name,savedAccount.getName());
    Assert.assertEquals(LedgerAccountSheetGenerator.holders,savedAccount.getHolders());
    Assert.assertEquals(LedgerAccountSheetGenerator.signatureAuthorities,savedAccount.getSignatureAuthorities());
    Assert.assertEquals(LedgerAccountSheetGenerator.balance,savedAccount.getBalance());
    Assert.assertEquals(LedgerAccountSheetGenerator.ledger,savedAccount.getLedger());
  }
}
