import { IConfigurationService } from './ConfigurationService';
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';

import pnp, { PermissionKind, Item } from 'sp-pnp-js';

export interface IUserContextService {
	currentUserIsSiteCollectionAdministrator() : Promise<boolean>;
}

export class UserContextService implements IUserContextService {
	constructor(serviceScope: ServiceScope) {}

	public currentUserIsSiteCollectionAdministrator() : Promise<boolean> {
    return pnp.sp.web.currentUser.select('IsSiteAdmin').get()
    .then(user => user.IsSiteAdmin as boolean);
  }
}

export const UserContextServiceKey = ServiceKey.create<IUserContextService>(
	'ypcode:usercontext-service',
	UserContextService
);
