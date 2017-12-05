import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, ServiceScope } from '@microsoft/sp-core-library';

import {
	BaseClientSideWebPart,
	IPropertyPaneConfiguration,
	PropertyPaneTextField,
	PropertyPaneDropdown,
	IPropertyPaneGroup,
	WebPartContext
} from '@microsoft/sp-webpart-base';

import * as strings from 'SalaryManagerWebPartStrings';
import SalaryManager from './components/SalaryManager';
import { ISalaryManagerProps } from './components/ISalaryManagerProps';

import AppStartup from './AppStartup';
import { IUserContextService, UserContextServiceKey } from './services/UserContextService';
import { IConfigurationService, ConfigurationServiceKey } from './services/ConfigurationService';
import { EFolderViewMode } from './common/EFolderViewMode';

export interface ISalaryManagerWebPartProps {
	salariesLibraryName: string;
	payrollOfficersGroupName: string;
	addSalaryFolderApiUrl: string;
	folderViewMode: EFolderViewMode;
}

export default class SalaryManagerWebPart extends BaseClientSideWebPart<ISalaryManagerWebPartProps> {
	private serviceScope: ServiceScope;
	private configService: IConfigurationService;
	private isSiteAdmin: boolean = false;

	public onInit(): Promise<any> {
		return (
			super
				.onInit()
				// Set the global configuration of the application
				// This is where we will define the proper services according to the context (Local, Test, Prod,...)
				// or according to specific settings
				.then((_) => AppStartup.configure(this.context, this.properties))
				// When configuration is done, we get the instances of the services we want to use
				.then((serviceScope) => {
					this.serviceScope = serviceScope;
					// Keep a reference to the config service
					this.configService = this.serviceScope.consume(ConfigurationServiceKey);
					// Get the UserContext service to check if current user is site coll admin
					let userContextService: IUserContextService = this.serviceScope.consume(UserContextServiceKey);
					return userContextService.currentUserIsSiteCollectionAdministrator();
				})
				.then((isSiteAdmin) => (this.isSiteAdmin = isSiteAdmin))
		);
	}

  private _configLastChange: number = Date.now();
	public render(): void {
		const element: React.ReactElement<ISalaryManagerProps> = React.createElement(SalaryManager, {
			serviceScope: this.serviceScope,
      context: this.context,
      configLastModified: this._configLastChange
		});

		ReactDom.render(element, this.domElement);
	}

	protected onPropertyPaneFieldChanged(property: string, oldValue: any, newValue: any) {
			this.configService[property] = newValue;
      console.log(`Property ${property} = ${newValue}`);
      this._configLastChange =Date.now();
	}

	protected get dataVersion(): Version {
		return Version.parse('1.0');
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		let displayGroup: IPropertyPaneGroup = {
			groupName: strings.DisplayConfigGroup,
			groupFields: [
				PropertyPaneDropdown('folderViewMode', {
					label: strings.FoldersViewModeLabel,
					selectedKey: this.properties.folderViewMode,
					options: [
						{ key: EFolderViewMode.Tiles, text: 'Tiles' },
						{ key: EFolderViewMode.Details, text: 'Detailed view' },
						{ key: EFolderViewMode.BigIcons, text: 'Big icons' }
					]
				})
			]
		};

		let adminGroup: IPropertyPaneGroup = {
			groupName: strings.AdminConfigGroup,
			groupFields: [
        PropertyPaneTextField('salariesLibraryName', {
					label: strings.SalariesLibraryNameLabel
        }),
        PropertyPaneTextField('payrollOfficersGroupName', {
					label: strings.PayrollOfficersGroupNameLabel
				}),
				PropertyPaneTextField('addSalaryFolderApiUrl', {
					label: strings.FolderApiUrlLabel
				})
			]
		};

		let groups: IPropertyPaneGroup[] = [ displayGroup ];

		if (this.isSiteAdmin) {
			groups.push(adminGroup);
		}

		if (this.context)
			return {
				pages: [
					{
						header: {
							description: strings.PropertyPaneDescription
						},
						groups: groups
					}
				]
			};
	}
}
