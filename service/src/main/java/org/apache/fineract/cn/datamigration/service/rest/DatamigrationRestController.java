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
package org.apache.fineract.cn.datamigration.service.rest;

import org.apache.fineract.cn.anubis.annotation.AcceptedTokenType;
import org.apache.fineract.cn.anubis.annotation.Permittable;
import org.apache.fineract.cn.command.gateway.CommandGateway;
import org.apache.fineract.cn.datamigration.api.v1.PermittableGroupIds;
import org.apache.fineract.cn.datamigration.service.internal.command.InitializeServiceCommand;
import org.apache.fineract.cn.datamigration.service.internal.service.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@RestController
@RequestMapping("/")
public class DatamigrationRestController {

  private final CommandGateway commandGateway;
  private final CustomerDatamigrationService customerDatamigrationService;
  private final OfficeMigration officeMigration;
  private final OfficeBranchMigration officeBranchMigration;
  private final EmployeeMigration employeeMigration;
  private final TellerDatamigration tellerDatamigration;
  private final GroupMigration groupMigration;
  private final LedgerMigration ledgerMigration;
  private final SubLedgerMigration subLedgerMigration;
  private final LedgersAccountsMigration ledgersAccountsMigration;
  private final UserMigration userMigration;
  private final JournalEntryMigration journalEntryMigration;
  private final LoanProductMigration loanProductMigration;
  private final TenantDatamigration tenantDatamigration  ;
  private final DepositActionMigration depositActionMigration ;
  private final ProductDefinitionMigration productDefinitionMigration ;
  private final ProductInstanceMigration productInstanceMigration  ;
  private final  GroupDefinitionMigration groupDefinitionMigration  ;
  private final  ChargeDefinitionMigration chargeDefinitionMigration  ;
  private final  TransactionTypeMigration transactionTypeMigration  ;




  @Autowired
  public DatamigrationRestController( final CommandGateway commandGateway,
                                      final CustomerDatamigrationService customerDatamigrationService,
                                      final OfficeMigration officeMigration,
                                      final OfficeBranchMigration officeBranchMigration,
                                      final EmployeeMigration employeeMigration,
                                      final TellerDatamigration  tellerDatamigration,
                                      final GroupMigration groupMigration,
                                      final LedgerMigration ledgerMigration,
                                      final SubLedgerMigration subLedgerMigration,
                                      final LedgersAccountsMigration ledgersAccountsMigration,
                                      final UserMigration userMigration,
                                      final JournalEntryMigration journalEntryMigration,
                                      final LoanProductMigration loanProductMigration,
                                      final TenantDatamigration  tenantDatamigration,
                                      final DepositActionMigration depositActionMigration,
                                      final ProductDefinitionMigration productDefinitionMigration,
                                      final ProductInstanceMigration productInstanceMigration,
                                      final  GroupDefinitionMigration groupDefinitionMigration,
                                      final  ChargeDefinitionMigration chargeDefinitionMigration,
                                      final  TransactionTypeMigration transactionTypeMigration
                                      ) {
    super();
    this.commandGateway = commandGateway;
    this.customerDatamigrationService = customerDatamigrationService;
    this.officeMigration = officeMigration;
    this.officeBranchMigration = officeBranchMigration;
    this.employeeMigration = employeeMigration;
    this.tellerDatamigration = tellerDatamigration;
    this.groupMigration = groupMigration;
    this.ledgerMigration = ledgerMigration;
    this.subLedgerMigration = subLedgerMigration;
    this.ledgersAccountsMigration = ledgersAccountsMigration;
    this.userMigration = userMigration;
    this.journalEntryMigration = journalEntryMigration;
    this.loanProductMigration = loanProductMigration;
    this.tenantDatamigration = tenantDatamigration;
    this.depositActionMigration=depositActionMigration;
    this.productDefinitionMigration=productDefinitionMigration;
    this.productInstanceMigration=productInstanceMigration;
    this.groupDefinitionMigration=groupDefinitionMigration;
    this.chargeDefinitionMigration=chargeDefinitionMigration;
    this.transactionTypeMigration=transactionTypeMigration;
  }

  @Permittable(value = AcceptedTokenType.SYSTEM)
  @RequestMapping(
      value = "/initialize",
      method = RequestMethod.POST,
      consumes = MediaType.ALL_VALUE,
      produces = MediaType.APPLICATION_JSON_VALUE
  )
  public ResponseEntity<Void> initialize()  {
      this.commandGateway.process(new InitializeServiceCommand());
      return ResponseEntity.accepted().build();
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/customers/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public  void  download(HttpServletResponse response) throws ClassNotFoundException {

    customerDatamigrationService.customersSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
            value = "/customers",
            method = RequestMethod.POST,
            consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> customersFormUpload(@RequestParam("file") MultipartFile file) throws IOException {
    customerDatamigrationService.customersSheetUpload(file);
        return new ResponseEntity<>("Upload successuly", HttpStatus.OK);

  }

  //Office Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/offices/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void officeSheetdownload(HttpServletResponse response) throws ClassNotFoundException {
    officeMigration.officeSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/offices",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> officeSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    officeMigration.officeSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //Branch Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/offices/branch/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void branchSheetDownload(HttpServletResponse response) throws ClassNotFoundException {
    officeBranchMigration.branchSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/offices/branch",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> branchSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    officeBranchMigration.branchSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //Employee Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/employees/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void employeeSheetdownload(HttpServletResponse response) throws ClassNotFoundException {
    employeeMigration.employeeSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/employees",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> employeeSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    employeeMigration.employeeSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //Teller Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/tellers/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void tellerSheetDownload(HttpServletResponse response) throws ClassNotFoundException {
    tellerDatamigration.tellerSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/tellers",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> tellerSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    tellerDatamigration.tellerSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //Group Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/group/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void groupSheetDownload(HttpServletResponse response) throws ClassNotFoundException {
    groupMigration.groupSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/group",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> groupSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    groupMigration.groupSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //Ledgers Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/ledgers/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void ledgerSheetdownload(HttpServletResponse response) throws ClassNotFoundException {
    ledgerMigration.ledgerSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/ledgers",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> ledgerSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    ledgerMigration.ledgerSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //SubLedgers Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/subLedgers/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void SubLedgerSheetdownload(HttpServletResponse response) throws ClassNotFoundException {
    subLedgerMigration.subLedgerSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/subLedgers",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> subLledgerSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    subLedgerMigration.subLedgerSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //Account Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/accounts/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void accountSheetdownload(HttpServletResponse response) throws ClassNotFoundException {
    ledgersAccountsMigration.accountSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/accounts",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> accountSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    ledgersAccountsMigration.accountSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //User Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/users/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void userSheetdownload(HttpServletResponse response) throws ClassNotFoundException {
    userMigration.userSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/users",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> userSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    userMigration.userSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }


  //JornalEnrty Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/journal/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void journalEntryDownload(HttpServletResponse response) throws ClassNotFoundException {
    journalEntryMigration.journalEntryDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/journal",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> journalEntrySheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    journalEntryMigration.journalEntrySheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //product Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/products/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void productSheetDownload(HttpServletResponse response) throws ClassNotFoundException {
    loanProductMigration.productSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/products",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> productSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    loanProductMigration.productSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //Tenant Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/tenants/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void tenantSheetDownload(HttpServletResponse response) throws ClassNotFoundException {
    tenantDatamigration.tenantSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/tenant",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> tenantSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    tenantDatamigration.tenantSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }


  //Charge definition Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/charge/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void chargeDefinitionSheetDownload(HttpServletResponse response) throws ClassNotFoundException {
    chargeDefinitionMigration.chargedefinitionSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/charge",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> chargeDefintionSheetUpload(@RequestParam("file") MultipartFile file) throws
          IOException {
    chargeDefinitionMigration.chargedefinitionSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }


  //Deposit Action Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/action/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void depositActionSheetDownload(HttpServletResponse response) throws ClassNotFoundException {
    depositActionMigration.actionSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/action",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> depositActionSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    depositActionMigration.actionSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }


  //Group Definition Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/group/definition/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void groupDefinitionSheetDownload(HttpServletResponse response) throws ClassNotFoundException {
    groupDefinitionMigration.groupDefinitionSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/group/definition",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> groupDefifnitionSheetUpload(@RequestParam("file") MultipartFile file) throws
          IOException {
    groupDefinitionMigration.groupDefinitionSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }


  //Product Definition Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/product/definition/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void productDefintionSheetDownload(HttpServletResponse response) throws ClassNotFoundException {
    productDefinitionMigration.productDefinitionSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/product/definition",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> productDefinitionSheetUpload(@RequestParam("file") MultipartFile file) throws
          IOException {
    productDefinitionMigration.productDefinitionSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }


  //Product Instance Migration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/product/instance/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void productInstanseSheetDownload(HttpServletResponse response) throws ClassNotFoundException {
    productInstanceMigration.productInstanceSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/product/instance",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> productInstanseSheetUpload(@RequestParam("file") MultipartFile file) throws
          IOException {
    productInstanceMigration.productInstanseSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

//Transaction Type
@Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
@RequestMapping(
        value = "/transation_type/download",
        method = RequestMethod.GET,
        consumes = MediaType.ALL_VALUE
)
public void transactionTypeSheetdownload(HttpServletResponse response) throws ClassNotFoundException {
  transactionTypeMigration.transactionTypeSheetDownload(response);
}

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/transation_type",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> transactionTypeSheetUpload(@RequestParam("file") MultipartFile file) throws
          IOException {
    transactionTypeMigration.transactionTypeSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }


}
