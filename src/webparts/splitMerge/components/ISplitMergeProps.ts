import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISplitMergeProps {
  title: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  sourceLibraryTitle: string;
  destinationLibraryTitle: string;
  destinationDocumentRepositoryTitle: string;
  documentTypeConfigListTitle: string;
  azureFunctionUrl: string;
  documentModelId: string;
  context: WebPartContext;
}
