import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';
import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import pnp, { PermissionKind, Item } from 'sp-pnp-js';

import { IConfigurationService, ConfigurationServiceKey } from './ConfigurationService';

export interface ISalaryFolder {
	id: number;
	name: string;
	url: string;
	employeeId: number;
	employeeName: string;
	employeeAccountName: string;
	updated: Date | string;
	documentsCount: number;
}

export interface ISalaryFolderCreationInfo {
	employeeLoginName: string;
}

export interface ISalaryFolderPermissions {
	canReadFiles: boolean;
	canAddFiles: boolean;
}

export interface IDocument {
	id: number;
	name: string;
	url: string;
	previewUrl: string;
}

export interface ISalaryDocumentsService {
	getAllSalaryFolders(): Promise<ISalaryFolder[]>;
	addSalaryFolder(folder: ISalaryFolderCreationInfo): Promise<any>;
	uploadDocuments(folder: ISalaryFolder, files: Blob[]): Promise<any>;
	getCurrentUserPermissionsOnFolder(folder: ISalaryFolder): Promise<ISalaryFolderPermissions>;
	getSalaryFolderDocuments(folder: ISalaryFolder): Promise<IDocument[]>;
	isCurrentUserPayrollOfficer(): Promise<boolean>;
}

export default class SalaryDocumentsService implements ISalaryDocumentsService {
	private config: IConfigurationService;
	private httpClient: HttpClient;

	constructor(private serviceScope: ServiceScope) {
		serviceScope.whenFinished(() => {
			this.config = serviceScope.consume(ConfigurationServiceKey);
			this.httpClient = serviceScope.consume(HttpClient.serviceKey);
		});
	}

	public getAllSalaryFolders(): Promise<ISalaryFolder[]> {
		let resultFolders: ISalaryFolder[] = null;
		return pnp.sp.web.lists
			.getByTitle(this.config.salariesLibraryName)
			.rootFolder.folders.filter("Name ne 'Forms'")
			.expand('ListItemAllFields')
			.select(
				'ListItemAllFields/Id',
				'Name',
				'TimeLastModified',
				'ListItemAllFields/EmployeeId',
				'ServerRelativeUrl',
				'ItemCount'
      )
      .orderBy('Name')
			.get()
			.then(
				(result: any[]) =>
					(resultFolders = result.map((f) => ({
						id: f.ListItemAllFields.Id,
						name: f.Name,
						url: f.ServerRelativeUrl,
						employeeId: f.ListItemAllFields.EmployeeId,
						employeeName: '',
						updated: f.TimeLastModified,
						documentsCount: f.ItemCount
					})) as ISalaryFolder[])
			)
			.then(() =>
				pnp.sp.web.siteUsers.select('Id', 'LoginName').get().then((allUsers: any[]) => {
					resultFolders.forEach((f) => {
						// Get user from allUsers
						let foundUsers = allUsers.filter((u) => u.Id == f.employeeId);
						if (foundUsers.length == 1) {
							let u = foundUsers[0];
							if (u.LoginName) {
								let accountName = u.LoginName.substr(u.LoginName.lastIndexOf('|') + 1);
								f.employeeAccountName = accountName;
							}
						}
					});
					return resultFolders;
				})
			)
			.then(() => resultFolders);
	}

	public addSalaryFolder(folder: ISalaryFolderCreationInfo): Promise<any> {
		const payload = { folder: folder };

		return this.httpClient
			.post(this.config.addSalaryFolderApiUrl, HttpClient.configurations.v1, {
				body: JSON.stringify(payload),
				mode: 'cors',
				credentials: 'include'
			})
			.then((response: HttpClientResponse) => {
				console.log('API response received.');
				return response.json();
			});
	}

	public uploadDocuments(folder: ISalaryFolder, files: File[]): Promise<any> {
		let count = 0;
		let promises = [];

		files.forEach((file) => {
			promises.push(
				pnp.sp.web.lists
					.getByTitle(this.config.salariesLibraryName)
					.items.getById(folder.id)
					.folder.files.add(file.name, file, true)
					.then((fileCreated) => fileCreated.file.getItem())
					.then((item: Item) =>
						item.update({
							EmployeeId: folder.employeeId
						})
					)
					.then(() => {
						count++;
					})
			);
		});

		return Promise.all(promises).then(() => files.length);
	}

	public isCurrentUserPayrollOfficer(): Promise<boolean> {
		return pnp.sp.web.currentUser.expand('Groups').select('Id', 'Groups').get().then((result) => {
			return (
				result.Groups &&
				result.Groups.filter((g) => g.Title == this.config.payrollOfficersGroupName).length == 1
			);
		});
	}

	public getCurrentUserPermissionsOnFolder(folder: ISalaryFolder): Promise<ISalaryFolderPermissions> {
		let spPermissions = null;
		return pnp.sp.web.lists
			.getByTitle(this.config.salariesLibraryName)
			.items.getById(folder.id)
			.getCurrentUserEffectivePermissions()
			.then((permissions) => {
				console.log(permissions);
				spPermissions = permissions;
			})
			.then(() => this.isCurrentUserPayrollOfficer())
			.then((isPayrollOfficer) => {
				return {
					canReadFiles: true,
					canAddFiles: true && isPayrollOfficer
				};
			});
	}

	public getSalaryFolderDocuments(folder: ISalaryFolder): Promise<IDocument[]> {
		return pnp.sp.web.lists
			.getByTitle(this.config.salariesLibraryName)
			.items.getById(folder.id)
			.folder.files.expand('ListItemAllFields')
			.select('Name', 'ServerRelativeUrl', 'ListItemAllFields/Id')
			.get()
			.then((files) => {
				console.log(files);
				return files;
			})
			.then((files) =>
				files.map((f) => ({
					id: f.ListItemAllFields.Id,
					name: f.Name,
					url: f.ServerRelativeUrl,
					previewUrl: ''
				}))
			);
	}
}

interface IMockFolder extends ISalaryFolder {
	documents: IDocument[];
	permissions: ISalaryFolderPermissions;
}

const MockSalaryFolders: IMockFolder[] = [
	{
		id: 1,
		name: 'Yannick Plenevaux',
		url: '/',
		employeeId: 8,
		employeeName: 'Yannick Plenevaux',
		employeeAccountName: '',
		updated: '2017-11-17',
		documentsCount: 5,
		permissions: {
			canAddFiles: true,
			canReadFiles: true
		},
		documents: [
			{
				id: 1,
				name: 'Document 1.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 2,
				name: 'Document 2.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 3,
				name: 'Document BLALBLBAL.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 4,
				name: 'Document QWERTZ.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 5,
				name: 'Document EWWEF.docx',
				url: '/',
				previewUrl: '/'
			}
		]
	},
	{
		id: 2,
		name: 'Christopher Clément',
		url: '/',
		employeeId: 9,
		employeeName: 'Christopher Clément',
		employeeAccountName: '',
		updated: '2017-11-17',
		documentsCount: 5,
		permissions: {
			canAddFiles: true,
			canReadFiles: true
		},
		documents: [
			{
				id: 1,
				name: 'Document 1.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 2,
				name: 'Document 2.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 3,
				name: 'Document BLALBLBAL.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 4,
				name: 'Document QWERTZ.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 5,
				name: 'Document EWWEF.docx',
				url: '/',
				previewUrl: '/'
			}
		]
	},
	{
		id: 3,
		name: 'Antoine Pichot',
		url: '/',
		employeeId: 10,
		employeeName: 'Antoine Pichot',
		employeeAccountName: '',
		updated: '2017-11-17',
		documentsCount: 5,
		permissions: {
			canAddFiles: true,
			canReadFiles: true
		},
		documents: [
			{
				id: 1,
				name: 'Document 1.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 2,
				name: 'Document 2.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 3,
				name: 'Document BLALBLBAL.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 4,
				name: 'Document QWERTZ.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 5,
				name: 'Document EWWEF.docx',
				url: '/',
				previewUrl: '/'
			}
		]
	},
	{
		id: 4,
		name: 'Stéphane Mertz',
		url: '/',
		employeeId: 12,
		employeeName: 'Stéphane Mertz',
		employeeAccountName: '',
		updated: '2017-11-17',
		documentsCount: 5,
		permissions: {
			canAddFiles: true,
			canReadFiles: true
		},
		documents: [
			{
				id: 1,
				name: 'Document 1.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 2,
				name: 'Document 2.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 3,
				name: 'Document BLALBLBAL.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 4,
				name: 'Document QWERTZ.docx',
				url: '/',
				previewUrl: '/'
			},
			{
				id: 5,
				name: 'Document EWWEF.docx',
				url: '/',
				previewUrl: '/'
			}
		]
	}
];

export class MockSalaryDocumentsService implements ISalaryDocumentsService {
	constructor(private serviceScope: ServiceScope) {}

	public getAllSalaryFolders(): Promise<ISalaryFolder[]> {
		return Promise.resolve(MockSalaryFolders);
	}

	public addSalaryFolder(folder: ISalaryFolderCreationInfo): Promise<any> {
		alert('MOCK ADD SALARY FOLDER');
		return Promise.resolve({});
	}

	public uploadDocuments(folder: ISalaryFolder, files: Blob[]): Promise<any> {
		alert('MOCK UPLOADED DOCUMENT');
		return Promise.resolve({});
	}

	public isCurrentUserPayrollOfficer(): Promise<boolean> {
		let currentUserId = null;
		return Promise.resolve(true);
	}

	public getCurrentUserPermissionsOnFolder(folder: ISalaryFolder): Promise<ISalaryFolderPermissions> {
		// TODO Implement this
		return Promise.resolve((folder as IMockFolder).permissions);
	}

	public getSalaryFolderDocuments(folder: ISalaryFolder): Promise<IDocument[]> {
		return Promise.resolve((folder as IMockFolder).documents);
	}
}

export const SalaryDocumentsServiceKey = ServiceKey.create<ISalaryDocumentsService>(
	'ypcode:salary-docs-service',
	SalaryDocumentsService
);
