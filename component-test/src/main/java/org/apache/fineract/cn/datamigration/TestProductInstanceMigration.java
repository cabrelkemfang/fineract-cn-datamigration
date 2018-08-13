package org.apache.fineract.cn.datamigration;

import org.apache.fineract.cn.datamigration.util.ProductInstanceSheetGenerator;
import org.apache.fineract.cn.deposit.api.v1.EventConstants;
import org.apache.fineract.cn.deposit.api.v1.instance.domain.ProductInstance;
import org.junit.Assert;
import org.junit.Test;

import java.util.List;

public class TestProductInstanceMigration extends AbstractDataMigrationTest {

  @Test
  public void shouldMigrateProductInstance() throws Exception {
    ProductInstanceSheetGenerator.productInstanceSheetDownload();

    //
    final boolean found = eventRecorder.wait(EventConstants.POST_PRODUCT_INSTANCE, ProductInstanceSheetGenerator.customerIdentifier);
    Assert.assertTrue(found);

    final List<ProductInstance> productInstances = depositAccountManager.findProductInstances(ProductInstanceSheetGenerator.accountIdentifier);
    Assert.assertNotNull(productInstances);

    Assert.assertEquals(1, productInstances.size());
    final ProductInstance foundProductInstance = productInstances.get(0);
    Assert.assertEquals(ProductInstanceSheetGenerator.customerIdentifier,foundProductInstance.getAccountIdentifier());
    Assert.assertEquals(ProductInstanceSheetGenerator.productIdentifier,foundProductInstance.getProductIdentifier());
    Assert.assertEquals(ProductInstanceSheetGenerator.accountIdentifier,foundProductInstance.getAccountIdentifier());
    Assert.assertEquals(ProductInstanceSheetGenerator.alternativeAccountNumber,foundProductInstance.getAlternativeAccountNumber());
    Assert.assertEquals(ProductInstanceSheetGenerator.state,foundProductInstance.getState());
    Assert.assertEquals(ProductInstanceSheetGenerator.balance,foundProductInstance.getBalance());

  }

  }
