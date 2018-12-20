package org.apache.fineract.cn.datamigration.service.internal.service.helper;

import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.group.api.v1.client.GroupManager;
import org.apache.fineract.cn.group.api.v1.domain.Group;
import org.apache.fineract.cn.group.api.v1.domain.GroupDefinition;
import org.slf4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Optional;

@Service
public class GroupService {
  private final Logger logger;
  private final GroupManager groupManager;

  @Autowired
  public GroupService(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                             final GroupManager groupManager) {
    super();
    this.logger = logger;
    this.groupManager = groupManager;
  }

  public Optional<String> createGroup(final Group group) {
    try {
      this.groupManager.createGroup(group);
      return Optional.empty();
    } catch (Exception e) {
      return Optional.of("Error while processing the group.");
    }
  }

  public Optional<String> createGroupDefinition(final GroupDefinition groupDefinition) {
    try {
      this.groupManager.createGroupDefinition(groupDefinition);
      return Optional.empty();
    } catch (Exception e) {
      return Optional.of("Error while processing the group definition.");
    }
  }

  public List<GroupDefinition> fetchGroupDefinitions() {
      return this.groupManager.fetchGroupDefinitions();
  }
}
