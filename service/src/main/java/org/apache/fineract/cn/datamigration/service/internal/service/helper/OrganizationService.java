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

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.office.api.v1.client.OrganizationManager;
import org.apache.fineract.cn.office.api.v1.domain.Employee;
import org.apache.fineract.cn.office.api.v1.domain.Office;
import org.slf4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Service;

import java.util.Optional;

@Service
public class OrganizationService {

  private final Logger logger;
  private final OrganizationManager organizationManager;

  @Autowired
  public OrganizationService(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                             final OrganizationManager organizationManager) {
    super();
    this.logger = logger;
    this.organizationManager = organizationManager;
  }

  public Optional<String> createOffice(final Office office) {
    try {
      this.organizationManager.createOffice(office);
      return Optional.empty();
    } catch (Exception e) {
      return Optional.of("Error while processing the office.");
    }
  }

  public Optional<String> createEmployee(final Employee employee) {
    try {
      this.organizationManager.createEmployee(employee);
      return Optional.empty();
    } catch (Exception e) {
      return Optional.of("Error while processing the employee.");
    }
  }

  public Optional<String> createBranchOffice(String officeIdentifier, final Office office) {
    try {
      this.organizationManager.addBranch(officeIdentifier, office);
      return Optional.empty();
    } catch (Exception e) {
      return Optional.of("Error while processing the  office branch.");
    }
  }

}

