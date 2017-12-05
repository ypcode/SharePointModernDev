import * as React from 'react';
import {
	Icon,
	IconType,
	Link,
	ProgressIndicator,
	Panel,
	DialogFooter,
	Button,
	ButtonType,
	Label,
	IPersonaProps,
	Spinner,
  SpinnerType,
  Persona
} from 'office-ui-fabric-react';
import styles from './SalaryManager.module.scss';
import { ISalaryManagerProps } from './ISalaryManagerProps';
import { escape } from '@microsoft/sp-lodash-subset';

import * as moment from 'moment';
import Dropzone from 'react-dropzone';

import {
	ISalaryDocumentsService,
	SalaryDocumentsServiceKey,
	IDocument,
	ISalaryFolderPermissions
} from '../services/SalaryDocumentsService';
import { ISalaryFolder } from '../services/SalaryDocumentsService';

import SharePointPeoplePicker from './SharePointPeoplePicker';
import { IConfigurationService, ConfigurationServiceKey } from '../services/ConfigurationService';
import { EFolderViewMode } from '../common/EFolderViewMode';

export interface ISalaryManagerState {
	folders: ISalaryFolder[];
	openedFolder: ISalaryFolder;
	dropzoneActive: boolean;
	documents: IDocument[];
	uploading: boolean;
	currentUploadProgress: number;
	userPermissionsOnOpenedFolder: ISalaryFolderPermissions;
	canAddFolder: boolean;
	addingFolder: boolean;
	creatingFolder: boolean;
	folderCreationStatus: string;
	folderCreationSuccess: boolean;
}

const UPLOAD_INTERVAL = 100;
const UPLOAD_INCREMENT = 0.05;

export default class SalaryManager extends React.Component<ISalaryManagerProps, ISalaryManagerState> {
	private salaryDocumentsService: ISalaryDocumentsService;
	private configService: IConfigurationService;
	private _authenticated: boolean = false;

	constructor(props: ISalaryManagerProps) {
		super(props);

		this.state = {
			folders: [],
			openedFolder: null,
			dropzoneActive: false,
			documents: null,
			uploading: false,
			currentUploadProgress: 0,
			userPermissionsOnOpenedFolder: null,
			canAddFolder: false,
			addingFolder: false,
			creatingFolder: false,
			folderCreationSuccess: false,
			folderCreationStatus: null
		};
	}

	public componentWillMount() {
		this.refreshComponent();
	}

	public componentWillReceiveProps(nextProps) {
		this.refreshComponent();
	}

	private refreshComponent() {
		this.props.serviceScope.whenFinished(() => {
			this.configService = this.props.serviceScope.consume(ConfigurationServiceKey);
			this.salaryDocumentsService = this.props.serviceScope.consume(SalaryDocumentsServiceKey);

      this.loadFolders()
      .then((folders: ISalaryFolder[]) => {
				this.salaryDocumentsService.isCurrentUserPayrollOfficer().then((isPayrollOfficer) => {
					this.setState({
						canAddFolder: isPayrollOfficer,
						folders: folders
					});
				});
			}).catch(error => {
        console.log('An error occured while loading folders: ', error);
        this.setState({
          canAddFolder: false,
          folders: []
        });
      });
		});
	}

	private loadFolders(): Promise<ISalaryFolder[]> {
		return this.salaryDocumentsService.getAllSalaryFolders();
	}

	private loadFolderDocuments(folder: ISalaryFolder): Promise<IDocument[]> {
		return this.salaryDocumentsService.getSalaryFolderDocuments(folder);
	}

	private openFolder(folder: ISalaryFolder) {
		let permissions: ISalaryFolderPermissions = null;
		this.salaryDocumentsService
			.getCurrentUserPermissionsOnFolder(folder)
			.then((perms) => {
				permissions = perms;
			})
			.then(() => this.loadFolderDocuments(folder))
			.then((documents) => {
				console.log(documents);
				this.setState({
					userPermissionsOnOpenedFolder: permissions,
					openedFolder: folder,
					documents: documents
				});
			});
	}

	private closeCurrentFolder() {
		this.setState({
			openedFolder: null,
			documents: null,
			userPermissionsOnOpenedFolder: null
		});
	}

	private executeOrDelayUntilAuthenticated(action: Function): void {
		if (this._authenticated) {
			console.log(this.configService.addSalaryFolderApiUrl);
			console.log('Is authenticated');
			action();
		} else {
			console.log('Still not authenticated');
			setTimeout(() => {
				this.executeOrDelayUntilAuthenticated(action);
			}, 1000);
		}
	}

	private onDragEnter() {
		let { dropzoneActive } = this.state;
		if (!dropzoneActive) {
			this.setState({
				dropzoneActive: true
			});
		}
	}

	private onDragLeave() {
		let { dropzoneActive } = this.state;
		if (dropzoneActive) {
			this.setState({
				dropzoneActive: false
			});
		}
	}

	private onDrop(files: File[]) {
		console.log(files);
		let { dropzoneActive, openedFolder, uploading, currentUploadProgress } = this.state;
		if (!openedFolder || !files || files.length == 0 || uploading) {
			return;
		}

		this.setState({
			uploading: true,
			dropzoneActive: false
		});

		let interval = setInterval(() => {
			if (currentUploadProgress < 0.9) {
				this.setState({
					currentUploadProgress: currentUploadProgress + UPLOAD_INCREMENT
				});
			}
		}, UPLOAD_INTERVAL);

		let folders: ISalaryFolder[] = null;
		let refreshOpenedFolder: ISalaryFolder = null;
		this.salaryDocumentsService
			.uploadDocuments(openedFolder, files)
			.then(() => this.loadFolders())
			.then((f) => (folders = f))
			.then(() => {
				let found = folders.filter((f) => f.id == openedFolder.id);
				if (found.length == 1) {
					refreshOpenedFolder = found[0];
				}
			})
			.then(() => this.loadFolderDocuments(openedFolder))
			.then((documents) => {
				clearInterval(interval);
				this.setState({
					openedFolder: refreshOpenedFolder,
					documents: documents,
					dropzoneActive: false,
					currentUploadProgress: 1.0,
					folders: folders
				});
				setTimeout(() => {
					this.setState({
						currentUploadProgress: 0,
						uploading: false
					});
				}, 1600);
			});
	}

	private _currentSelectedEmployee: IPersonaProps;
	public onEmployeeSelected(items: IPersonaProps[]) {
		if (items && items.length == 1) {
			this._currentSelectedEmployee = items[0];
			console.log('CURRENTLY SELECTED EMPLOYEE: ', this._currentSelectedEmployee);
		}
	}

	public addNewFolder() {
		// this.salaryDocumentsService.addSalaryFolder({
		//   name
		// })
		this.setState({
			addingFolder: true,
			folderCreationSuccess: null,
			folderCreationStatus: null
		});
	}

	public confirmAddFolder() {
		let {} = this.state;
		this.setState({
			creatingFolder: true
		});
		this.executeOrDelayUntilAuthenticated(() => {
			this.salaryDocumentsService
				.addSalaryFolder({
					employeeLoginName: this._currentSelectedEmployee.itemID
				})
				.then(() => {
					console.log('FOLDER CREATED');

					this.loadFolders().then((folders: ISalaryFolder[]) => {
						this.setState({
							folders: folders,
							addingFolder: false,
							creatingFolder: false,
							folderCreationSuccess: true,
							folderCreationStatus: 'The folder has been created'
						});
					});
				})
				.catch((error) => {
					console.log(error);
					this.setState({
						creatingFolder: false,
						folderCreationSuccess: false,
						folderCreationStatus: 'The folder could not be created'
					});
				});
		});
	}

	public cancelAddFolder() {
		this.setState({
			addingFolder: false,
			folderCreationSuccess: null,
			folderCreationStatus: null
		});
	}

	private _renderFoldersView() {
		let { folders } = this.state;

		return this._renderViewContainer(
			folders.map((f) => {
				switch (this.configService.folderViewMode) {
					case EFolderViewMode.BigIcons:
						return this._renderFolderBigIcon(f);
					case EFolderViewMode.Details:
						return this._renderFolderDetailRow(f);
					case EFolderViewMode.Tiles:
					default:
						return this._renderFolderTile(f);
				}
			})
		);
	}

	private _renderViewContainer(content: any): JSX.Element {
		switch (this.configService.folderViewMode) {
			case EFolderViewMode.BigIcons:
				return <div>{content}</div>;
			case EFolderViewMode.Details:
				return (
					<div>
						<div className="ms-Grid-row  ">
							<div className="ms-Grid-col ms-sm1" />
							<div className="ms-Grid-col ms-sm5">
								<Label>Name</Label>
							</div>
							<div className="ms-Grid-col ms-sm2">
								<Label>Documents</Label>
							</div>
							<div className="ms-Grid-col ms-sm4">
								<Label>Updated</Label>
							</div>
						</div>
						{content}
					</div>
				);
			case EFolderViewMode.Tiles:
			default:
				return <div>{content}</div>;
		}
	}

	private _renderFolderTile(salaryFolder: ISalaryFolder): JSX.Element {
		return (
			<div className="ms-Grid-col ms-sm12 ms-md6 ms-fontColor-themePrimary--hover  ">
				<div className={styles.folderTile} onClick={() => this.openFolder(salaryFolder)}>
					<div className="ms-Grid-col ms-sm6 ms-md3">
						<div className={styles.folderIcon}>
							<Icon iconName="FolderHorizontal" />
						</div>
					</div>
					<div className="ms-Grid-col ms-sm6 ms-md9 ms-font-l">
						<div className={styles.folderInfo}>
							<div className="ms-font-l"> {salaryFolder.name}</div>
							<div className="ms-font-m">
								{salaryFolder.documentsCount} document{salaryFolder.documentsCount > 1 ? 's' : ''}
							</div>
							<div className="ms-font-s-plus">
								{salaryFolder.updated ? 'Last update ' + moment(salaryFolder.updated).fromNow() : ''}
							</div>
						</div>
					</div>
				</div>
			</div>
		);
	}

	private _renderFolderDetailRow(salaryFolder: ISalaryFolder) {
		return (
			<div className="ms-Grid-row ms-fontColor-themePrimary--hover  ">
				<div className="ms-Grid-col ms-sm1">
					<Icon iconName="FolderHorizontal" />
				</div>
				<div className="ms-Grid-col ms-sm5">
					<Link href="#" onClick={() => this.openFolder(salaryFolder)}>{salaryFolder.name}</Link>
				</div>
				<div className="ms-Grid-col ms-sm2">{salaryFolder.documentsCount}</div>
				<div className="ms-Grid-col ms-sm4">
					{salaryFolder.updated ? moment(salaryFolder.updated).format('DD/MM/YYYY - hh:mm:ss') : ''}
				</div>
			</div>
		);
	}

	private _renderFolderBigIcon(salaryFolder: ISalaryFolder) {
		const tooltip = () =>
			`${salaryFolder.documentsCount} document${salaryFolder.documentsCount > 1
				? 's'
				: ''}\n${salaryFolder.updated ? 'Last update ' + moment(salaryFolder.updated).fromNow() : ''}`;

		return (
			<div className="ms-Grid-col ms-sm12 ms-md6 ms-xl4 ms-xxl3 ms-fontColor-themePrimary--hover  ">
				<div className={styles.bigIcon} title={tooltip()} onClick={() => this.openFolder(salaryFolder)}>
					<div className={styles.folderIcon}>
            <Persona imageUrl={`/_layouts/15/userphoto.aspx?size=S&accountname=${salaryFolder.employeeAccountName}`} className={styles.persona} />
						<Icon iconName="FolderHorizontal" />
					</div>
					<div>{salaryFolder.name}</div>
				</div>
			</div>
		);
	}

	public render(): React.ReactElement<ISalaryManagerProps> {
		let {
			folders,
			openedFolder,
			dropzoneActive,
			documents,
			uploading,
			currentUploadProgress,
			userPermissionsOnOpenedFolder,
			canAddFolder,
			addingFolder,
			creatingFolder,
			folderCreationStatus,
			folderCreationSuccess
		} = this.state;

		const getDocumentIcon = (doc: IDocument) => {
			let extension = '';
			let parts = doc.name && doc.name.split('.');
			if (parts && parts.length > 1) {
				extension = parts[parts.length - 1];
			}
			switch (extension) {
				case 'doc':
				case 'docx':
					return 'WordDocument';
				case 'xls':
				case 'xlsx':
					return 'ExcelDocument';
				case 'ppt':
				case 'pptx':
					return 'PowerPointDocument';
				case 'pdf':
					return 'Pdf';
				default:
					return 'Document';
			}
		};

		let folderBrowse = openedFolder && (
			<div className={styles.folderContent}>
				<div className="ms-Grid-col ms-sm6 ms-md2">
					<span className={styles.backIcon}>
						<a onClick={() => this.closeCurrentFolder()}>
							<Icon iconName="ChromeBack" title="Go back" />
						</a>
					</span>
					<span className={styles.folderIcon}>
						<Icon iconName="OpenFolderHorizontal" />
					</span>
				</div>
				<div className="ms-Grid-col ms-sm6 ms-md10 ms-font-l">
					<div className={styles.folderInfo}>
						<div className="ms-Grid-col ms-sm12 ms-font-xxl"> {openedFolder.name}</div>
						<div className="ms-Grid-col ms-sm6 ms-font-l">
							{openedFolder.documentsCount} document{openedFolder.documentsCount > 1 ? 's' : ''}
						</div>
						<div className="ms-Grid-col ms-sm6 ms-font-l-plus">
							{openedFolder.updated ? 'Last update ' + moment(openedFolder.updated).fromNow() : ''}
						</div>
					</div>
				</div>
				<div className="ms-Grid-col ms-sm12">
					<hr />
				</div>
				<div className="ms-Grid-col ms-sm12">
					{uploading && (
						<ProgressIndicator percentComplete={currentUploadProgress} label="Uploading document..." />
					)}
					<Dropzone
						disabled={userPermissionsOnOpenedFolder && !userPermissionsOnOpenedFolder.canAddFiles}
						onDrop={this.onDrop.bind(this)}
						disableClick
						style={{ position: 'relative' }}
						onDragEnter={this.onDragEnter.bind(this)}
						onDragLeave={this.onDragLeave.bind(this)}
					>
						{dropzoneActive && <div className={styles.dragOverlay}>Drop files...</div>}
						<div className={styles.filesPlaceHolder}>
							{documents &&
								documents.map((doc) => (
									<div className="ms-Grid-col ms-sm6">
										<Link href={doc.url}>
											<Icon iconName={getDocumentIcon(doc)} />&nbsp;
											<span className="ms-font-l">{doc.name}</span>
										</Link>
									</div>
								))}
						</div>
					</Dropzone>
				</div>
			</div>
		);

		let foldersList = (
			<div>
				<iframe
					src={this.configService.addSalaryFolderApiUrl}
					style={{ display: 'none' }}
					onLoad={() => (this._authenticated = true)}
				/>
				{canAddFolder &&
				this.configService.addSalaryFolderApiUrl && (
					<div className="ms-Grid-col ms-sm12 ">
						<Link href="#" onClick={() => this.addNewFolder()}>
							<Icon iconName="FabricNewFolder" /> Add folder
						</Link>
					</div>
				)}
				{addingFolder && (
					<Panel isOpen={addingFolder}>
						<div className="ms-Grid-row">
							<div className="ms-Grid-col ms-sm12 ">
								<h2>Create a new Salary Folder</h2>
							</div>
						</div>
						<div className="ms-Grid-row">
							<div className="ms-Grid-col ms-sm12 ">
								<Label>Select the employee</Label>
								<SharePointPeoplePicker
									typePicker="Compact"
									description="Enter the employee's name"
									spHttpClient={this.props.context.spHttpClient}
									siteUrl={this.props.context.pageContext.web.absoluteUrl}
									principalTypeUser={true}
									principalTypeSharePointGroup={false}
									principalTypeSecurityGroup={false}
									principalTypeDistributionList={false}
									numberOfItems={10}
									maxSelectedItems={1}
									onItemsChange={(items) => this.onEmployeeSelected(items)}
								/>
							</div>
						</div>
						<div className="ms-Grid-row">
							<div className="ms-Grid-col ms-sm12 ">
								<br />
								<br />
								<br />
								<br />
								{creatingFolder && <Spinner type={SpinnerType.large} label="Creating folder..." />}
								{folderCreationStatus && (
									<span
										className={
											folderCreationSuccess == true ? (
												'ms-fontColor-green'
											) : folderCreationSuccess == false ? (
												'ms-fontColor-red'
											) : (
												''
											)
										}
									>
										{folderCreationStatus}
									</span>
								)}
							</div>
						</div>
						<DialogFooter>
							<Button
								buttonType={ButtonType.primary}
								disabled={creatingFolder}
								onClick={() => this.confirmAddFolder()}
							>
								Ok
							</Button>
							<Button
								buttonType={ButtonType.default}
								disabled={creatingFolder}
								onClick={() => this.cancelAddFolder()}
							>
								Cancel
							</Button>
						</DialogFooter>
					</Panel>
				)}
				{this._renderFoldersView()}
			</div>
		);

		return (
			<div className={styles.salaryManager}>
				<div className={styles.container}>
					<div className={`ms-Grid-row ms-fontColor-themeDarker  ${styles.row}`}>
						{openedFolder ? folderBrowse : foldersList}
					</div>
				</div>
			</div>
		);
	}
}
