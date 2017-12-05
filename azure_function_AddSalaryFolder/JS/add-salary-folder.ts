import pnp, { setup, Web, Item, NodeFetchClient } from 'sp-pnp-js';
declare var process;
declare var module;

var SITE_URL = process.env.SITE_URL;
var CLIENT_ID = process.env.CLIENT_ID;
var CLIENT_SECRET = process.env.CLIENT_SECRET;
var SALARIES_LIB_NAME = process.env.SALARIES_LIB_NAME;
var PAYROLL_OFFICERS_GROUP_NAME = process.env.PAYROLL_OFFICERS_GROUP_NAME;

const READ_ROLE_ID = 1073741826;
const CONTRIBUTE_ROLE_ID = 1073741827;

class SalaryFolderService {
	private _web: Web;

	constructor(
		private siteUrl: string,
		private clientId: string,
		private clientSecret: string,
		private currentUserLogin: string,
		private payrollOfficersGroupName: string,
		private salariesLibraryName: string,
		private log: (...s: string[]) => void
	) {
		pnp.setup({
			sp: {
				fetchClientFactory: () => new NodeFetchClient(siteUrl, clientId, clientSecret)
			}
		});

		this._web = new Web(siteUrl);
	}

	public isCurrentUserPayrollOfficer(): Promise<boolean> {
		return this._web.siteUsers
			.getByLoginName(`i:0#.f|membership|${this.currentUserLogin}`)
			.expand('Groups')
			.select('Id', 'Groups/Title')
			.get()
			.then((result) => {
				this.log('Current user groups are :', result.Groups);
				return (
					result.Groups &&
					result.Groups.filter((g) => g.Title == PAYROLL_OFFICERS_GROUP_NAME).length == 1
				);
			})
			.catch((error) => {
				this.log('Error while trying to check if current user is payroll officer', error);
				if (error && error.data && error.data.responseBody && error.data.responseBody['odata.error']) {
					this.log('ODATA Error', error.data.responseBody['odata.error']);
				}
				return false;
			});
	}

	public addSalaryFolder(employeeLoginName: string): Promise<any> {
		let employee: any = null;
		let newFolderItem: Item = null;
		let payrollOfficersGroup: any = null;

		// Get the payroll officers group
		return (
			this._web.siteGroups
				.getByName(this.payrollOfficersGroupName)
				.get()
				.then((pog) => (payrollOfficersGroup = pog))
				// Ensure the employee user on the current web
				.then(() => this._web.ensureUser(employeeLoginName))
				// Get the employee information
				.then((ensuredUser) => ensuredUser.user.get())
				.then((foundEmployee: any) => {
					employee = foundEmployee;
				})
				// Add a folder in the Salaries library
				.then(() => this._web.lists.getByTitle(this.salariesLibraryName).rootFolder.folders.add(employee.Title))
				.then((folderCreation) => folderCreation.folder.getItem())
				// Break permissions inherithance on the new folder
				.then((item) => {
					newFolderItem = item;
					return item.breakRoleInheritance(false);
				})
				// Assign unique permissions on the new folder
				.then(() => newFolderItem.roleAssignments.add(employee.Id, READ_ROLE_ID))
				.then(() => newFolderItem.roleAssignments.add(payrollOfficersGroup.Id, CONTRIBUTE_ROLE_ID))
				// Set metadata on the folder to bind it to the employee
				.then(() =>
					newFolderItem.update({
						EmployeeId: employee.Id
					})
				)
				// Return the folder item
				.then(() => newFolderItem)
		);
	}
}

function respond(context: any, data: any, status: number = 200) {
	context.res = {
		headers: {
			'Content-Type': 'application/json',
			'Access-Control-Allow-Credentials': 'true',
			'Access-Control-Allow-Origin': 'https://ike365.sharepoint.com',
			'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
			'Access-Control-Allow-Headers': 'Content-Type, Set-Cookie',
			'Access-Control-Max-Age': '86400'
		},
		status: status,
		body: data
	};
	context.log(context.res);
	context.done();
}

module.exports = function(context, req) {
	context.log('AZURE FUNCTION EXECUTION!!!');
	context.log(`Site Url  		: ${SITE_URL}`);
	context.log(`Client Id 		: ${CLIENT_ID}`);
	context.log(`Client Secret 	: ${CLIENT_SECRET}`);
	context.log(`Payroll Group	: ${PAYROLL_OFFICERS_GROUP_NAME}`);
	context.log(`Library Name	: ${SALARIES_LIB_NAME}`);

	var currentUserLoginName = req['headers']['x-ms-client-principal-name'];
	context.log('Current user is : ', currentUserLoginName);

	context.log('Request: ', req);

	if (!req.body) {
		context.log('No arguments to process. Terminates...');
		respond(context, { status: 'Properly authenticated' });
		return;
	}

	let payload = null;
	let response = { error: null, data: null };
	try {
		payload = JSON.parse(req.body);
	} catch (error) {
		response.error = error;
	}

	if (!payload.folder) {
		context.log('Folder argument is not specified or cannot be parsed');
		respond(context, response, response.error ? 400 : 200);
		return;
	}

	var employeeLoginName = payload.folder.employeeLoginName;
	context.log('ARGUMENT=', employeeLoginName);
	var service = new SalaryFolderService(
		SITE_URL,
		CLIENT_ID,
		CLIENT_SECRET,
		currentUserLoginName,
		PAYROLL_OFFICERS_GROUP_NAME,
		SALARIES_LIB_NAME,
		context.log
	);

	service
		.isCurrentUserPayrollOfficer()
		.then((isPayrollOfficer) => {
			if (!isPayrollOfficer) {
				response.data = 'You are not a Payroll Officer and are not allowed to add a folder';
				respond(context, response, 403);
				return;
			} else {
				service
					.addSalaryFolder(employeeLoginName)
					.then(function() {
						context.log('Folder created');
						response.error = null;
						response.data = 'Success';
						respond(context, response);
					})
					.catch(function(error) {
						context.log('Folder cannot be created');
						context.log(error);
						response.error = error;
						response.data = 'Failure';
						respond(context, response, 400);
					});
			}
		})
		.catch(function(error) {
			context.log(error);
			response.error = error;
			respond(context, response, 500);
		});
};
