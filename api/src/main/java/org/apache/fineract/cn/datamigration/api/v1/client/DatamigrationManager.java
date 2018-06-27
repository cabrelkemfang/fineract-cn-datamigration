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

  //customer datamigration
  @RequestMapping(
          value = "/customers/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  void download(HttpServletResponse response) ;

  @RequestMapping(
          value = "/customers",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  ResponseEntity<String> customersFormUpload(@RequestParam("file") MultipartFile file) ;

  //Office Datamigration
  @RequestMapping(
          value = "/offices/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void officeSheetdownload(HttpServletResponse response) ;


  @RequestMapping(
          value = "/offices",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> officeSheetUpload(@RequestParam("file") MultipartFile file) ;

  //employee datamigration
  @RequestMapping(
          value = "/employees/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void employeeSheetdownload(HttpServletResponse response) ;

  @RequestMapping(
          value = "/employees",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
   ResponseEntity<String> employeeSheetUpload(@RequestParam("file") MultipartFile file);

//tellers datmigration
  @RequestMapping(
          value = "/tellers/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
   void tellerSheetDownload(HttpServletResponse response) ;

  @RequestMapping(
          value = "/tellers",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  ResponseEntity<String> tellerSheetUpload(@RequestParam("file") MultipartFile file) ;

  //group datmigration
  @RequestMapping(
          value = "/group/download",
          method = RequestMethod.GET,
          consumes = MediaType.ALL_VALUE
  )
  void groupSheetDownload(HttpServletResponse response) ;

  @RequestMapping(
          value = "/group",
          method = RequestMethod.POST,
          consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  ResponseEntity<String> groupSheetUpload(@RequestParam("file") MultipartFile file) ;

}
