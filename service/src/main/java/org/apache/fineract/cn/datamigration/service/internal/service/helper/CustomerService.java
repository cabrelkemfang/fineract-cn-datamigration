package org.apache.fineract.cn.datamigration.service.internal.service.helper;


import org.apache.fineract.cn.customer.api.v1.client.CustomerManager;
import org.apache.fineract.cn.customer.api.v1.domain.Customer;
import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.slf4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Service;

import java.util.Optional;

@Service
public class CustomerService {

  private final Logger logger;
  private final CustomerManager customerManager;

  @Autowired
  public CustomerService(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                           final CustomerManager customerManager) {
    super();
    this.logger = logger;
    this.customerManager = customerManager;
  }

  public Optional<String> createCustomer(final Customer customer) {
    try {
      this.customerManager.createCustomer(customer);
      return Optional.empty();
    } catch (Exception e) {
      return Optional.of("Error while processing the customer.");
    }
  }

}
