package org.apache.fineract.cn.datamigration;

import org.apache.fineract.cn.datamigration.util.ProductDefinitionSheetGenerator;
import org.apache.fineract.cn.deposit.api.v1.EventConstants;
import org.apache.fineract.cn.deposit.api.v1.definition.domain.ProductDefinition;
import org.junit.Assert;
import org.junit.Test;

public class TestProductDefinitionMigration extends AbstractDataMigrationTest {

  @Test
  public void shouldMigrateProductDefinition() throws Exception {

    ProductDefinitionSheetGenerator.productDefinitionSheetDownload();

    //
    final boolean found = eventRecorder.wait(EventConstants.POST_PRODUCT_DEFINITION, ProductDefinitionSheetGenerator.identifier);
    Assert.assertTrue(found);

    final ProductDefinition productDefinition = depositAccountManager.findProductDefinition(ProductDefinitionSheetGenerator.identifier);
    Assert.assertNotNull(productDefinition);

    Assert.assertEquals(ProductDefinitionSheetGenerator.identifier,productDefinition.getIdentifier());
    Assert.assertEquals(ProductDefinitionSheetGenerator.type,productDefinition.getType());
    Assert.assertEquals(ProductDefinitionSheetGenerator.name,productDefinition.getName());
    Assert.assertEquals(ProductDefinitionSheetGenerator.description,productDefinition.getDescription());
    Assert.assertEquals(ProductDefinitionSheetGenerator.currencyCode,productDefinition.getCurrency().getCode());
    Assert.assertEquals(ProductDefinitionSheetGenerator.currencyName,productDefinition.getCurrency().getName());
    Assert.assertEquals(ProductDefinitionSheetGenerator.currencySign,productDefinition.getCurrency().getSign());
    Assert.assertEquals(ProductDefinitionSheetGenerator.currencyScale,productDefinition.getCurrency().getScale());
    Assert.assertEquals(ProductDefinitionSheetGenerator.minimumBalance,productDefinition.getMinimumBalance());
    Assert.assertEquals(ProductDefinitionSheetGenerator.equityLedgerIdentifier,productDefinition.getEquityLedgerIdentifier());
    Assert.assertEquals(ProductDefinitionSheetGenerator.cashAccountIdentifier,productDefinition.getCashAccountIdentifier());
    Assert.assertEquals(ProductDefinitionSheetGenerator.expenseAccountIdentifier,productDefinition.getExpenseAccountIdentifier());
    Assert.assertEquals(ProductDefinitionSheetGenerator.accrueAccountIdentifier,productDefinition.getAccrueAccountIdentifier());
    Assert.assertEquals(ProductDefinitionSheetGenerator.interest,productDefinition.getInterest());
    Assert.assertEquals(ProductDefinitionSheetGenerator.termPeriod,productDefinition.getTerm().getPeriod());
    Assert.assertEquals(ProductDefinitionSheetGenerator.termTimeUnit,productDefinition.getTerm().getTimeUnit());
    Assert.assertEquals(ProductDefinitionSheetGenerator.termInterestPayable,productDefinition.getTerm().getInterestPayable());
    Assert.assertEquals(ProductDefinitionSheetGenerator.flexible,productDefinition.getFlexible());
    Assert.assertEquals(ProductDefinitionSheetGenerator.active,productDefinition.getActive());
  }
}
