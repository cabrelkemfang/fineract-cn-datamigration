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
  private final OfficeDatamigrationService officeDatamigrationService;
  private final OfficeBranchDatamigration officeBranchDatamigration;
  private final EmployeeDatamigration employeeDatamigration;
  private final TellerDatamigration tellerDatamigration;
  private final GroupDatamigration groupDatamigration;

  @Autowired
  public DatamigrationRestController( final CommandGateway commandGateway,
                                      final CustomerDatamigrationService customerDatamigrationService,
                                      final OfficeDatamigrationService officeDatamigrationService,
                                      final OfficeBranchDatamigration officeBranchDatamigration,
                                      final EmployeeDatamigration employeeDatamigration,
                                      final TellerDatamigration  tellerDatamigration,
                                      final GroupDatamigration  groupDatamigration) {
    super();
    this.commandGateway = commandGateway;
    this.customerDatamigrationService = customerDatamigrationService;
    this.officeDatamigrationService = officeDatamigrationService;
    this.officeBranchDatamigration = officeBranchDatamigration;
    this.employeeDatamigration = employeeDatamigration;
    this.tellerDatamigration = tellerDatamigration;
    this.groupDatamigration = groupDatamigration;
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

  //Office Datamigration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/offices/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void officeSheetdownload(HttpServletResponse response) throws ClassNotFoundException {
    officeDatamigrationService.officeSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/offices",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> officeSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    officeDatamigrationService.officeSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //Branch Datamigration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/offices/branch/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void branchSheetDownload(HttpServletResponse response) throws ClassNotFoundException {
    officeBranchDatamigration.branchSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/offices/branch",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> branchSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    officeBranchDatamigration.branchSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //Employee Datamigration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/employees/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void employeeSheetdownload(HttpServletResponse response) throws ClassNotFoundException {
    employeeDatamigration.employeeSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/employees",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> employeeSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    employeeDatamigration.employeeSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

  //Teller Datamigration
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

  //Group Datamigration
  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/group/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  public void groupSheetDownload(HttpServletResponse response) throws ClassNotFoundException {
    groupDatamigration.groupSheetDownload(response);
  }

  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
          value = "/group",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> groupSheetUpload(@RequestParam("file") MultipartFile file) throws IOException {
    groupDatamigration.groupSheetUpload(file);
    return new ResponseEntity<>("Upload successuly", HttpStatus.OK);
  }

}
