import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Environment, EnvironmentType, ServiceScope } from '@microsoft/sp-core-library';
import pnp from 'sp-pnp-js';

import { IConfigurationService, ConfigurationServiceKey } from './services/ConfigurationService';
import { SalaryDocumentsServiceKey, MockSalaryDocumentsService } from './services/SalaryDocumentsService';
import {ISalaryManagerWebPartProps} from './SalaryManagerWebPart';
export default class AppStartup {
	public static configure(webPartContext: WebPartContext, webPartProperties: ISalaryManagerWebPartProps) : Promise<ServiceScope> {

    console.log("Environment= #############", Environment);
    console.log("Environment Type= #############", Environment.type);
		switch (Environment.type) {
			case EnvironmentType.SharePoint:
			case EnvironmentType.ClassicSharePoint:
				return AppStartup.configureForSharePoint(webPartContext.serviceScope, webPartContext, webPartProperties);
			// case EnvironmentType.Local:
			// case EnvironmentType.Test:
			default:
				return AppStartup.configureForLocalTesting(webPartContext.serviceScope, webPartContext, webPartProperties);
		}
	}

	private static configureForSharePoint(
		serviceScope: ServiceScope,
		webPartContext: WebPartContext,
		webPartProperties: ISalaryManagerWebPartProps
	) : Promise<ServiceScope> {
		return new Promise<any>((resolve, reject) => {
			serviceScope.whenFinished(() => {
				// Configure PnP Js for working seamlessly with SPFx
				pnp.setup({
					spfxContext: webPartContext
				});

				let config: IConfigurationService = serviceScope.consume(ConfigurationServiceKey);
        config.salariesLibraryName = webPartProperties.salariesLibraryName;
        config.addSalaryFolderApiUrl = webPartProperties.addSalaryFolderApiUrl;
        config.payrollOfficersGroupName = webPartProperties.payrollOfficersGroupName;
        config.folderViewMode = webPartProperties.folderViewMode;

				resolve(serviceScope);
			});
		});
	}

	private static configureForLocalTesting(
		serviceScope: ServiceScope,
		webPartContext: WebPartContext,
		webPartProperties: ISalaryManagerWebPartProps
	) : Promise<ServiceScope> {
		return new Promise<any>((resolve, reject) => {
			// Here create a dedicated service scope for test or local context
			const childScope: ServiceScope = serviceScope.startNewChild();
			// Register the services that will override default implementation
			childScope.createAndProvide(SalaryDocumentsServiceKey, MockSalaryDocumentsService);
			// Must call the finish() method to make sure the child scope is ready to be used
			childScope.finish();

			childScope.whenFinished(() => {
				// If other services must be used, it must done HERE

				resolve(childScope);
			});
		});
	}
}
