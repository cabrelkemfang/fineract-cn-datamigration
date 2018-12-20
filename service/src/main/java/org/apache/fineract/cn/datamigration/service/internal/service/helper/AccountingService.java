/*
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 *   http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
package org.apache.fineract.cn.datamigration.service.internal.service.helper;


import org.apache.fineract.cn.accounting.api.v1.client.LedgerManager;
import org.apache.fineract.cn.accounting.api.v1.domain.*;
import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.slf4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Service;

import java.util.Optional;

@Service
public class AccountingService {

  private final Logger logger;
  private final LedgerManager ledgerManager;

  @Autowired
  public AccountingService(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                           final LedgerManager ledgerManager) {
    super();
    this.logger = logger;
    this.ledgerManager = ledgerManager;
  }

  public Optional<String> createLeger(final Ledger ledger) {
    try {
      this.ledgerManager.createLedger(ledger);
      return Optional.empty();
    } catch (Exception e) {
      return Optional.of("Error while processing the ledger.");
    }
  }

  public LedgerPage fetchLeger() {
    return this.ledgerManager.fetchLedgers(false, null, null, null, null, null, null);
  }

  public Optional<String> createSubLeger(String ledgerIdentifier, final Ledger ledger) {
    try {
      this.ledgerManager.addSubLedger(ledgerIdentifier, ledger);
      return Optional.empty();
    } catch (Exception e) {
      return Optional.of("Error while processing the subLedger.");
    }
  }

  public Optional<String> createAccount(final Account account) {
    try {
      this.ledgerManager.createAccount(account);
      return Optional.empty();
    } catch (Exception e) {
      return Optional.of("Error while processing the account.");
    }
  }

  public Optional<String> createTransactionType(final TransactionType transactionType) {
    try {
      this.ledgerManager.createTransactionType(transactionType);
      return Optional.empty();
    } catch (Exception e) {
      return Optional.of("Error while processing the transaction type.");
    }
  }

  public Optional<String> createJournalEntry(final JournalEntry journalEntry) {
    try {
      this.ledgerManager.createJournalEntry(journalEntry);
      return Optional.empty();
    } catch (Exception e) {
      return Optional.of("Error while processing the journal entry.");
    }
  }

  public TransactionTypePage fetchTransactionTypes() {
    return this.ledgerManager.fetchTransactionTypes(null, 0, 100, "identifier", Sort.Direction.DESC.name());
  }
}
