import { ServiceScope } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { EFolderViewMode } from '../common/EFolderViewMode';

export interface ISalaryManagerProps {
  serviceScope: ServiceScope;
  context: WebPartContext;
  configLastModified: number;
}
