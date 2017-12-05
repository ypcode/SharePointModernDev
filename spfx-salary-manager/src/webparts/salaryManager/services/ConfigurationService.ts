import { IConfigurationService } from './ConfigurationService';
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { EFolderViewMode } from '../common/EFolderViewMode';

export interface IConfigurationService {
	salariesLibraryName: string;
	addSalaryFolderApiUrl: string;
	payrollOfficersGroupName: string;
	folderViewMode: EFolderViewMode;
}

export class ConfigurationService implements IConfigurationService {
	constructor(serviceScope: ServiceScope) {}

	public salariesLibraryName: string;
	public addSalaryFolderApiUrl: string;
	public payrollOfficersGroupName: string;
	public folderViewMode: EFolderViewMode;
}

export const ConfigurationServiceKey = ServiceKey.create<IConfigurationService>(
	'ypcode:config-service',
	ConfigurationService
);
