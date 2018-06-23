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
import org.apache.fineract.cn.datamigration.service.internal.service.DatamigrationService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.IOException;

@RestController
@RequestMapping("/")
public class DatamigrationRestController {


  private final CommandGateway commandGateway;
  private final DatamigrationService datamigrationService;

  @Autowired
  public DatamigrationRestController( final CommandGateway commandGateway,
                                      final DatamigrationService datamigrationService) {
    super();
    this.commandGateway = commandGateway;
    this.datamigrationService = datamigrationService;
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

      /*ByteArrayInputStream bis = datamigrationService.customersFormDownload();
      HttpHeaders headers = new HttpHeaders();
      headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
      headers.add("Content-Disposition", "attachment; filename=customers.xlsx");
      return ResponseEntity
              .ok()
              .headers(headers)
              .contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
              .body(new InputStreamResource(bis));
              */
   datamigrationService.customersFormDownload(response);

  }



  @Permittable(value = AcceptedTokenType.TENANT, groupId = PermittableGroupIds.DATAMIGRATION_MANAGEMENT)
  @RequestMapping(
            value = "/customers",
            method = RequestMethod.POST,
            consumes = MediaType.MULTIPART_FORM_DATA_VALUE
  )
  public ResponseEntity<String> customersFormUpload(@RequestParam("file") MultipartFile file) throws IOException {
        datamigrationService.customersFormUpload(file);
        return new ResponseEntity<>("Upload successuly", HttpStatus.OK);

  }


}
