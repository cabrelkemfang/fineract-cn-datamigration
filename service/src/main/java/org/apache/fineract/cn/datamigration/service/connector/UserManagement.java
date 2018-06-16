package org.apache.fineract.cn.datamigration.service.connector;

import org.apache.fineract.cn.api.util.UserContextHolder;
import org.apache.fineract.cn.datamigration.service.ServiceConstants;
import org.apache.fineract.cn.identity.api.v1.client.IdentityManager;
import org.apache.fineract.cn.identity.api.v1.domain.Authentication;
import org.slf4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;
import org.springframework.util.Base64Utils;

@Service
public class UserManagement {

    private final Logger logger;
    private final IdentityManager identityManager;
    @Value("${fineract-cn.system.user.name}")
    private String systemUserName;
    @Value("${fineract-cn.system.user.password}")
    private String systemUserPassword;

    @Autowired
    public UserManagement(@Qualifier(ServiceConstants.LOGGER_NAME) final Logger logger,
                          final IdentityManager identityManager) {
        super();
        this.logger = logger;
        this.identityManager = identityManager;
    }

    public String getRoleByUser(final String userName) {
        return this.identityManager.getUser(userName).getRole();
    }

    public void authenticate() {
        UserContextHolder.clear();
        final Authentication authentication =
                this.identityManager.login(this.systemUserName, Base64Utils.encodeToString(this.systemUserPassword.getBytes()));
        UserContextHolder.setAccessToken(this.systemUserName, authentication.getAccessToken());
    }
}

