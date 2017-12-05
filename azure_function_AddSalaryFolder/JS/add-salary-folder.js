"use strict";
exports.__esModule = true;
var sp_pnp_js_1 = require("sp-pnp-js");
var SITE_URL = process.env.SITE_URL;
var CLIENT_ID = process.env.CLIENT_ID;
var CLIENT_SECRET = process.env.CLIENT_SECRET;
var SALARIES_LIB_NAME = process.env.SALARIES_LIB_NAME;
var PAYROLL_OFFICERS_GROUP_NAME = process.env.PAYROLL_OFFICERS_GROUP_NAME;
var READ_ROLE_ID = 1073741826;
var CONTRIBUTE_ROLE_ID = 1073741827;
var SalaryFolderService = (function () {
    function SalaryFolderService(siteUrl, clientId, clientSecret, currentUserLogin, payrollOfficersGroupName, salariesLibraryName, log) {
        this.siteUrl = siteUrl;
        this.clientId = clientId;
        this.clientSecret = clientSecret;
        this.currentUserLogin = currentUserLogin;
        this.payrollOfficersGroupName = payrollOfficersGroupName;
        this.salariesLibraryName = salariesLibraryName;
        this.log = log;
        sp_pnp_js_1["default"].setup({
            sp: {
                fetchClientFactory: function () { return new sp_pnp_js_1.NodeFetchClient(siteUrl, clientId, clientSecret); }
            }
        });
        this._web = new sp_pnp_js_1.Web(siteUrl);
    }
    SalaryFolderService.prototype.isCurrentUserPayrollOfficer = function () {
        var _this = this;
        return this._web.siteUsers
            .getByLoginName("i:0#.f|membership|" + this.currentUserLogin)
            .expand('Groups')
            .select('Id', 'Groups/Title')
            .get()
            .then(function (result) {
            _this.log('Current user groups are :', result.Groups);
            return (result.Groups &&
                result.Groups.filter(function (g) { return g.Title == PAYROLL_OFFICERS_GROUP_NAME; }).length == 1);
        })["catch"](function (error) {
            _this.log('Error while trying to check if current user is payroll officer', error);
            if (error && error.data && error.data.responseBody && error.data.responseBody['odata.error']) {
                _this.log('ODATA Error', error.data.responseBody['odata.error']);
            }
            return false;
        });
    };
    SalaryFolderService.prototype.addSalaryFolder = function (employeeLoginName) {
        var _this = this;
        var employee = null;
        var newFolderItem = null;
        var payrollOfficersGroup = null;
        // Get the payroll officers group
        return (this._web.siteGroups
            .getByName(this.payrollOfficersGroupName)
            .get()
            .then(function (pog) { return (payrollOfficersGroup = pog); })
            .then(function () { return _this._web.ensureUser(employeeLoginName); })
            .then(function (ensuredUser) { return ensuredUser.user.get(); })
            .then(function (foundEmployee) {
            employee = foundEmployee;
        })
            .then(function () { return _this._web.lists.getByTitle(_this.salariesLibraryName).rootFolder.folders.add(employee.Title); })
            .then(function (folderCreation) { return folderCreation.folder.getItem(); })
            .then(function (item) {
            newFolderItem = item;
            return item.breakRoleInheritance(false);
        })
            .then(function () { return newFolderItem.roleAssignments.add(employee.Id, READ_ROLE_ID); })
            .then(function () { return newFolderItem.roleAssignments.add(payrollOfficersGroup.Id, CONTRIBUTE_ROLE_ID); })
            .then(function () {
            return newFolderItem.update({
                EmployeeId: employee.Id
            });
        })
            .then(function () { return newFolderItem; }));
    };
    return SalaryFolderService;
}());
function respond(context, data, status) {
    if (status === void 0) { status = 200; }
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
module.exports = function (context, req) {
    context.log('AZURE FUNCTION EXECUTION!!!');
    context.log("Site Url  \t\t: " + SITE_URL);
    context.log("Client Id \t\t: " + CLIENT_ID);
    context.log("Client Secret \t: " + CLIENT_SECRET);
    context.log("Payroll Group\t: " + PAYROLL_OFFICERS_GROUP_NAME);
    context.log("Library Name\t: " + SALARIES_LIB_NAME);
    var currentUserLoginName = req['headers']['x-ms-client-principal-name'];
    context.log('Current user is : ', currentUserLoginName);
    context.log('Request: ', req);
    if (!req.body) {
        context.log('No arguments to process. Terminates...');
        respond(context, { status: 'Properly authenticated' });
        return;
    }
    var payload = null;
    var response = { error: null, data: null };
    try {
        payload = JSON.parse(req.body);
    }
    catch (error) {
        response.error = error;
    }
    if (!payload.folder) {
        context.log('Folder argument is not specified or cannot be parsed');
        respond(context, response, response.error ? 400 : 200);
        return;
    }
    var employeeLoginName = payload.folder.employeeLoginName;
    context.log('ARGUMENT=', employeeLoginName);
    var service = new SalaryFolderService(SITE_URL, CLIENT_ID, CLIENT_SECRET, currentUserLoginName, PAYROLL_OFFICERS_GROUP_NAME, SALARIES_LIB_NAME, context.log);
    service
        .isCurrentUserPayrollOfficer()
        .then(function (isPayrollOfficer) {
        if (!isPayrollOfficer) {
            response.data = 'You are not a Payroll Officer and are not allowed to add a folder';
            respond(context, response, 403);
            return;
        }
        else {
            service
                .addSalaryFolder(employeeLoginName)
                .then(function () {
                context.log('Folder created');
                response.error = null;
                response.data = 'Success';
                respond(context, response);
            })["catch"](function (error) {
                context.log('Folder cannot be created');
                context.log(error);
                response.error = error;
                response.data = 'Failure';
                respond(context, response, 400);
            });
        }
    })["catch"](function (error) {
        context.log(error);
        response.error = error;
        respond(context, response, 500);
    });
};
