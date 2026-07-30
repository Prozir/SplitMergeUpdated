import type { IDropdownOption } from '@fluentui/react';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

// Stores the data for one PDF page in the preview flow.
export interface IPageInfo {
  id: string;
  sourceFileRef: string;
  sourceFileName: string;
  sourcePageNumber: number;
  selected: boolean;
  thumbnail?: string;
}

// Holds the basic file details needed to open a PDF.
export interface IPdfSelection {
  fileRef: string;
  fileName: string;
}

// Stores one detected document type and its page numbers.
export interface IClassificationResult {
  key: string;
  text: string;
  confidence: number;
  pageNumbers: number[];
}

// Passes the modal settings and context needed for auto-classify.
export interface IAutoClassifyModalProps {
  isOpen: boolean;
  onDismiss: () => void;
  selectedPdfFile: IPdfSelection | null;
  entityOptions: IDropdownOption[];
  sourceLibraryTitle: string;
  destinationLibraryTitle: string;
  destinationDocumentRepositoryTitle: string;
  context: WebPartContext;
  isBusy?: boolean;
  onUploadSuccess: () => Promise<void>;
}
