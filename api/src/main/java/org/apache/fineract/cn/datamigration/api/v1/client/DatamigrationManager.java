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
package org.apache.fineract.cn.datamigration.api.v1.client;
import org.apache.fineract.cn.api.util.CustomFeignClientsConfiguration;
import org.apache.fineract.cn.datamigration.api.v1.PermittableGroupIds;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;

@FeignClient(value="datamigration-v1", path="/datamigration/v1", configuration = CustomFeignClientsConfiguration.class)
public interface DatamigrationManager {

  //customer Migration
  @RequestMapping(
          value = "/customers/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  void download(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/customers",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  ResponseEntity<String> customersFormUpload(@RequestParam("file") MultipartFile file) ;

  //Office Migration
  @RequestMapping(
          value = "/offices/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void officeSheetdownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/offices",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> officeSheetUpload(@RequestParam("file") MultipartFile file) ;

//Branch Migration
  @RequestMapping(
          value = "/offices/branch/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void branchSheetDownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/offices/branch",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  ResponseEntity<String> branchSheetUpload(@RequestParam("file") MultipartFile file) ;

  //employee Migration
  @RequestMapping(
          value = "/employees/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void employeeSheetdownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/employees",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> employeeSheetUpload(@RequestParam("file") MultipartFile file);

//tellers Migration
  @RequestMapping(
          value = "/tellers/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void tellerSheetDownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/tellers",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  ResponseEntity<String> tellerSheetUpload(@RequestParam("file") MultipartFile file) ;

  //group Migration
  @RequestMapping(
          value = "/group/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  void groupSheetDownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/group",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  ResponseEntity<String> groupSheetUpload(@RequestParam("file") MultipartFile file) ;


  //Ledgers Migration
  @RequestMapping(
          value = "/ledgers/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void ledgerSheetdownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/ledgers",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> ledgerSheetUpload(@RequestParam("file") MultipartFile file) ;


  //SubLedgers Migration
  @RequestMapping(
          value = "/subLedgers/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  void subLedgerSheetdownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/subLedgers",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  ResponseEntity<String> subLedgerSheetUpload(@RequestParam("file") MultipartFile file) ;


  //Account Migration
  @RequestMapping(
          value = "/accounts/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
 void accountSheetdownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/accounts",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> accountSheetUpload(@RequestParam("file") MultipartFile file) ;


  //User Migration
  @RequestMapping(
          value = "/users/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void userSheetdownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/users",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> userSheetUpload(@RequestParam("file") MultipartFile file) ;


  //JornalEnrty Migration
  @RequestMapping(
          value = "/journal/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void journalEntryDownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/journal",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> journalEntrySheetUpload(@RequestParam("file") MultipartFile file) ;


  //product Migration
  @RequestMapping(
          value = "/products/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void productSheetDownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/products",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> productSheetUpload(@RequestParam("file") MultipartFile file) ;

  //Tenant Migration
  @RequestMapping(
          value = "/tenants/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void tenantSheetDownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/tenant",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> tenantSheetUpload(@RequestParam("file") MultipartFile file);


  //Charge definition Migration
  @RequestMapping(
          value = "/charge/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void chargeDefinitionSheetDownload(HttpServletResponse response) throws ClassNotFoundException;


  @RequestMapping(
          value = "/charge",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> chargeDefintionSheetUpload(@RequestParam("file") MultipartFile file) throws
          IOException ;


  //Deposit Action Migration
  @RequestMapping(
          value = "/action/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void depositActionSheetDownload(HttpServletResponse response) throws ClassNotFoundException;

  @RequestMapping(
          value = "/action",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> depositActionSheetUpload(@RequestParam("file") MultipartFile file) throws IOException ;



  //Group Definition Migration
  @RequestMapping(
          value = "/group/definition/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void groupDefinitionSheetDownload(HttpServletResponse response) throws ClassNotFoundException ;

  @RequestMapping(
          value = "/group/definition",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> groupDefifnitionSheetUpload(@RequestParam("file") MultipartFile file) throws
          IOException ;



  //Product Definition Migration
  @RequestMapping(
          value = "/product/definition/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void productDefintionSheetDownload(HttpServletResponse response) throws ClassNotFoundException;


  @RequestMapping(
          value = "/product/definition",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> productDefinitionSheetUpload(@RequestParam("file") MultipartFile file) throws
          IOException ;



  //Product Instance Migration
  @RequestMapping(
          value = "/product/instance/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void productInstanseSheetDownload(HttpServletResponse response) throws ClassNotFoundException ;

  @RequestMapping(
          value = "/product/instance",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> productInstanseSheetUpload(@RequestParam("file") MultipartFile file) throws
          IOException ;


  //Transaction Type
  @RequestMapping(
          value = "/transation_type/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void transactionTypeSheetdownload(HttpServletResponse response) throws ClassNotFoundException ;


  @RequestMapping(
          value = "/transation_type",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> transactionTypeSheetUpload(@RequestParam("file") MultipartFile file) throws
          IOException ;



}
