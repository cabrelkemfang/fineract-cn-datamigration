package org.apache.fineract.cn.datamigration;

import org.apache.fineract.cn.datamigration.util.OfficeSheetGenerator;
import org.apache.fineract.cn.datamigration.util.TellersSheetGenerator;
import org.apache.fineract.cn.teller.api.v1.EventConstants;
import org.apache.fineract.cn.teller.api.v1.domain.Teller;
import org.junit.Assert;
import org.junit.Test;

public class TestTellerMigration extends AbstractDataMigrationTest {
  @Test
  public void shouldMigrateUser() throws Exception {

    TellersSheetGenerator.tellerSheetDownload();

    //

    final boolean found = eventRecorder.wait(EventConstants.POST_TELLER, TellersSheetGenerator.code);
    Assert.assertTrue(found);

    final Teller teller  = tellerManager.find(OfficeSheetGenerator.identifier,TellersSheetGenerator.code);
    Assert.assertNotNull(teller);

    Assert.assertEquals(TellersSheetGenerator.code,teller.getCode());
    Assert.assertEquals(TellersSheetGenerator.cashdrawLimit,teller.getCashdrawLimit());
    Assert.assertEquals(TellersSheetGenerator.tellerAccountIdentifier,teller.getTellerAccountIdentifier());
    Assert.assertEquals(TellersSheetGenerator.vaultAccountIdentifier,teller.getVaultAccountIdentifier());
    Assert.assertEquals(TellersSheetGenerator.chequesReceivableAccount,teller.getChequesReceivableAccount());
    Assert.assertEquals(TellersSheetGenerator.cashOverShortAccount,teller.getCashOverShortAccount());
    Assert.assertEquals(TellersSheetGenerator.denominationRequired,teller.getDenominationRequired());
    Assert.assertEquals(TellersSheetGenerator.assignedEmployee,teller.getAssignedEmployee());
    Assert.assertEquals(TellersSheetGenerator.state,teller.getState());
  }

  }
