import * as React from 'react';
import styles from './SplitMerge.module.scss';
import type { ISplitMergeProps } from './ISplitMergeProps';
import { SPHttpClient } from '@microsoft/sp-http';
import { PrimaryButton, TextField, Checkbox, Label, Spinner, SpinnerSize, DetailsList, IColumn, Selection, SelectionMode, Modal, IconButton, Link, Dropdown, IDropdownOption } from '@fluentui/react';
import * as pdfjsLib from 'pdfjs-dist';
import { PDFDocument } from 'pdf-lib';
import AutoClassifyModal from './AutoClassifyModal';
import { buildDistinctEntityOptions, checkInFile, getRepositoryFolderUrl, safeODataString } from './splitMergeHelpers';
import { IPageInfo, IPdfSelection } from './splitMergeTypes';

// Set PDF.js worker to use local worker
pdfjsLib.GlobalWorkerOptions.workerSrc = require('pdfjs-dist/build/pdf.worker.min.js');

export default class SplitMerge extends React.Component<ISplitMergeProps, {
  pdfFiles: any[];
  selectedPdfFiles: IPdfSelection[];
  selectedPdfName: string;
  pages: IPageInfo[];
  currentPageNumber: number;
  loading: boolean;
  newContractNumber: string;
  newDocumentType: string;
  uploading: boolean;
  uploadingSource: boolean;
  errorMessage: string;
  showModal: boolean;
  documentTypes: IDropdownOption[];
  loadingDocumentTypes: boolean;
  entityOptions: IDropdownOption[];
  loadingEntities: boolean;
  selectedEntityKey: string;
  selectedEntitySiteUrl: string;
  showAutoClassifyModal: boolean;
  selectedAutoClassifyFile: IPdfSelection | null;
  disableAutoClassify: boolean;
}> {
  private pdfDocuments: { [fileRef: string]: any } = {};
  private selection: Selection;
  private previewCanvasRef = React.createRef<HTMLCanvasElement>();
  private fileInputRef = React.createRef<HTMLInputElement>();

  // Create the component state and selection helper.
  constructor(props: ISplitMergeProps) {
    super(props);
    this.selection = new Selection({ onSelectionChanged: this.handleSelectionChanged });
    this.state = {
      pdfFiles: [],
      selectedPdfFiles: [],
      selectedPdfName: '',
      pages: [],
      currentPageNumber: 1,
      loading: false,
      newContractNumber: '',
      newDocumentType: '',
      uploading: false,
      uploadingSource: false,
      errorMessage: '',
      showModal: false,
      showAutoClassifyModal: false,
      selectedAutoClassifyFile: null,
      documentTypes: [],
      loadingDocumentTypes: false,
      entityOptions: [],
      loadingEntities: false,
      selectedEntityKey: '',
      selectedEntitySiteUrl: ''
      ,
      disableAutoClassify: false
    };
  }

  // Start loading the main content when the component opens.
  componentDidMount() {
    this.initializePdfJs();
    this.loadPdfFiles();
    this.loadDocumentTypes();
    this.loadEntityOptions();
  }

  // Prepare the PDF library for use.
  private async initializePdfJs() {
    try {
      // PDF.js worker is configured at module load; nothing to initialize here
    } catch (error) {
      this.setState({ errorMessage: 'PDF.js library failed to initialize. Please refresh the page.' });
    }
  }

  // React when props or modal state change.
  componentDidUpdate(prevProps: ISplitMergeProps, prevState: Readonly<any>) {
    if (prevProps.sourceLibraryTitle !== this.props.sourceLibraryTitle) {
      this.setState({ pdfFiles: [], selectedPdfFiles: [], selectedPdfName: '', pages: [], errorMessage: '', showModal: false });
      this.loadPdfFiles();
    }

    if (prevProps.documentTypeConfigListTitle !== this.props.documentTypeConfigListTitle) {
      this.loadDocumentTypes();
    }

    if (!prevState.showModal && this.state.showModal) {
      if (this.state.documentTypes.length === 0) {
        this.loadDocumentTypes();
      }
      if (this.state.pages.length > 0) {
        requestAnimationFrame(() => {
          this.renderPdfPage(this.state.currentPageNumber).catch(error => {
            console.error('Error rendering PDF page after modal open:', error);
          });
        });
      }
    }
  }

  private async loadPdfFiles() {
    const { sourceLibraryTitle, context } = this.props;
    if (!sourceLibraryTitle) return;

    try {
      // Load the PDF files from the source library.
      const files: any[] = [];
      let nextUrl: string | undefined = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${safeODataString(sourceLibraryTitle)}')/items?$filter=substringof('.pdf',FileLeafRef)&$select=FileLeafRef,FileRef,Created,Modified,Author/Title,AssignedTo/Title,AssignedTo/EMail,AzureResponse,AutoClassifyStatus&$expand=Author,AssignedTo&$orderby=Created desc&$top=5000`;

      while (nextUrl) {
        const response = await context.spHttpClient.get(nextUrl, SPHttpClient.configurations.v1);

        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const data = await response.json();
        const pageItems = data.value || data.d?.results || [];
        files.push(...pageItems);
        nextUrl = data['@odata.nextLink'] || data['odata.nextLink'] || data.d?.__next;
      }

      this.setState({ pdfFiles: files, errorMessage: '' });
    } catch (error) {
      console.error('Error loading PDF files:', error);
      this.setState({ errorMessage: 'Error loading PDF files. Please check the library title and permissions.' });
    }
  }

  // Track which PDF files are selected.
  private handleSelectionChanged = () => {
    const selectedItems = this.selection.getSelection() as any[];
    const selectedPdfFiles = selectedItems.map(item => ({ fileRef: item.fileRef, fileName: item.fileName }));
    const disable = !(selectedItems.length === 1 && selectedItems[0]?.azureResponse && String(selectedItems[0].azureResponse).trim() !== '');
    this.setState({ selectedPdfFiles, disableAutoClassify: disable });
  };

  // Open one PDF file for viewing.
  private async loadPdf(fileUrl: string, fileName: string) {
    return this.loadSelectedPdfs([{ fileRef: fileUrl, fileName }]);
  }

  // Open the selected PDFs and prepare their pages.
  private async loadSelectedPdfs(selectedFiles: IPdfSelection[]) {
    this.setState({ loading: true, pages: [], errorMessage: '' });
    this.pdfDocuments = {};

    try {
      // Build the page list for the selected PDFs.
      const pages: IPageInfo[] = [];

      for (const file of selectedFiles) {
        const response = await this.props.context.spHttpClient.get(file.fileRef, SPHttpClient.configurations.v1);

        if (!response.ok) {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const arrayBuffer = await response.arrayBuffer();
        if (arrayBuffer.byteLength === 0) {
          throw new Error('PDF file is empty');
        }

        const uint8Array = new Uint8Array(arrayBuffer);
        if (uint8Array.length < 4 || uint8Array[0] !== 37 || uint8Array[1] !== 80 || uint8Array[2] !== 68 || uint8Array[3] !== 70) {
          throw new Error('File is not a valid PDF');
        }

        // Open the PDF so its pages can be previewed.
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        this.pdfDocuments[file.fileRef] = pdf;

        for (let i = 1; i <= pdf.numPages; i++) {
          pages.push({
            id: `${file.fileRef}|${i}`,
            sourceFileRef: file.fileRef,
            sourceFileName: file.fileName,
            sourcePageNumber: i,
            selected: false
          });
        }
      }

      this.setState({
        pages,
        loading: false,
        selectedPdfFiles: selectedFiles,
        selectedPdfName: selectedFiles.map(f => f.fileName).join(', '),
        currentPageNumber: 1,
        showModal: true
      });
    } catch (error) {
      console.error('Error loading PDF(s):', error);
      let errorMessage = 'Error loading PDF(s). ';

      if (error instanceof Error) {
        if (error.message.indexOf('HTTP') !== -1) {
          errorMessage += 'File access denied or file not found. Please check permissions and file path.';
        } else if (error.message.indexOf('empty') !== -1) {
          errorMessage += 'One of the PDF files appears to be empty.';
        } else if (error.message.indexOf('valid PDF') !== -1) {
          errorMessage += 'One of the files is not a valid PDF document.';
        } else if (error.message.indexOf('InvalidPDFException') !== -1) {
          errorMessage += 'One of the PDF files is corrupted or invalid.';
        } else if (error.message.indexOf('MissingPDFException') !== -1) {
          errorMessage += 'One of the PDF files was not found or is inaccessible.';
        } else {
          errorMessage += error.message;
        }
      } else {
        errorMessage += 'Please check file permissions and try again.';
      }

      this.setState({ loading: false, errorMessage });
    }
  }

  // Load document type choices for the dropdown.
  private async loadDocumentTypes() {
    const { documentTypeConfigListTitle, context } = this.props;
    if (!documentTypeConfigListTitle) {
      this.setState({ documentTypes: [], loadingDocumentTypes: false });
      return;
    }

    this.setState({ loadingDocumentTypes: true });

    try {
      // Load the document type options from the config list.
      const response = await context.spHttpClient.get(
        `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${safeODataString(documentTypeConfigListTitle)}')/items?$top=1000`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const data = await response.json();

      // Filter for active items (handle various possible column names and value formats)
      const options: IDropdownOption[] = data.value
        .filter((item: any) => {
          // Check multiple possible column names for Yes/No field
          const isActive = item.IsActive || item.IsActive_x0020_Status || item.IsActiveStatus;
          // For Yes/No columns, true means yes, false means no
          return isActive === true || isActive === 1 || isActive === 'Yes' || isActive === 'true';
        })
        .map((item: any) => {
          // Try multiple possible column names for document type
          const docType = item.Title || item.DocumentType || item.Document_x0020_Type || item['Document Type'];
          return {
            key: docType,
            text: docType
          };
        })
        .sort((a: IDropdownOption, b: IDropdownOption) => 
          (a.text as string).localeCompare(b.text as string)
        );

      this.setState({ documentTypes: options, loadingDocumentTypes: false });
    } catch (error) {
      console.error('Error loading document types:', error);
      this.setState({ documentTypes: [], loadingDocumentTypes: false });
    }
  }

  // Load entity options for the repository dropdown.
  private async loadEntityOptions() {
    const { context } = this.props;
    const listTitle = 'EntityDocumentRepositoryConfig';
    this.setState({ loadingEntities: true });

    try {
      // Load entity choices that have a CMS site URL.
      const response = await context.spHttpClient.get(
        `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${safeODataString(listTitle)}')/items?$select=Entity,CMSSiteURL&$filter=CMSSiteURL ne null and CMSSiteURL ne ''&$top=500`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to load entity config list: ${response.status} ${response.statusText}`);
      }

      const data = await response.json();
      const options: IDropdownOption[] = buildDistinctEntityOptions(data.value as any[]);

      this.setState({ entityOptions: options, loadingEntities: false });
    } catch (error) {
      console.error('Error loading entity options:', error);
      this.setState({ entityOptions: [], loadingEntities: false });
    }
  }

  // Close the preview modal and clear the state.
  private handleModalClose = () => {
    this.pdfDocuments = {};
    this.selection.setAllSelected(false);
    this.setState({ showModal: false, pages: [], selectedPdfFiles: [], selectedPdfName: '', currentPageNumber: 1 });
  };

  // Open a PDF when the user clicks its name.
  private handlePdfSelect = (fileRef: string, fileName: string) => {
    this.loadPdf(fileRef, fileName);
  };

  // Open the selected PDFs for review.
  private handleOpenSelectedClick = () => {
    const { selectedPdfFiles } = this.state;
    if (selectedPdfFiles.length === 0) {
      return;
    }
    this.loadSelectedPdfs(selectedPdfFiles);
  };

  // Open the auto-classify workflow for one file.
  private handleOpenClassifyClick = () => {
    const { selectedPdfFiles } = this.state;
    if (selectedPdfFiles.length !== 1) {
      alert('Please select a single PDF file to auto classify.');
      return;
    }

    this.setState({
      selectedAutoClassifyFile: selectedPdfFiles[0],
      showAutoClassifyModal: true
    });
  };

  // Close the auto-classify modal and reset its state.
  private handleAutoClassifyDismiss = () => {
    this.pdfDocuments = {};
    this.selection.setAllSelected(false);
    this.setState({
      showAutoClassifyModal: false,
      selectedAutoClassifyFile: null,
      pages: [],
      selectedPdfFiles: [],
      selectedPdfName: '',
      currentPageNumber: 1
    });
  };

  // Open the file picker for uploading a PDF.
  private handleUploadButtonClick = () => {
    if (this.fileInputRef.current) {
      this.fileInputRef.current.value = '';
      this.fileInputRef.current.click();
    }
  };

  // Upload a file chosen by the user.
  private handleFileInputChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const { sourceLibraryTitle, context } = this.props;
    const file = event.target.files && event.target.files[0];

    if (!file || !sourceLibraryTitle) {
      return;
    }

    if (file.type !== 'application/pdf') {
      alert('Please select a PDF file.');
      return;
    }

    this.setState({ uploadingSource: true, errorMessage: '' });

    try {
      // Upload the new PDF into the source library.
      const fileName = file.name;
      const uploadUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${safeODataString(sourceLibraryTitle)}')/RootFolder/Files/add(url='${encodeURIComponent(fileName)}',overwrite=true)`;

      const uploadResponse = await context.spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
        body: file,
        headers: {
          'Content-Type': 'application/pdf',
          'Accept': 'application/json;odata=nometadata'
        }
      });

      if (!uploadResponse.ok) {
        throw new Error(`File upload failed: ${uploadResponse.status} ${uploadResponse.statusText}`);
      }

      alert(`Uploaded ${fileName} successfully to ${sourceLibraryTitle}.`);
      await this.loadPdfFiles();
    } catch (error) {
      console.error('Error uploading PDF file:', error);
      this.setState({ errorMessage: 'Error uploading PDF file. Please try again.' });
    } finally {
      this.setState({ uploadingSource: false });
    }
  };

  // Mark a page as selected or not.
  private handlePageSelect = (pageId: string, selected: boolean) => {
    this.setState(prevState => ({
      pages: prevState.pages.map(page =>
        page.id === pageId ? { ...page, selected } : page
      )
    }));
  };

  // Show the chosen page in the preview area.
  private async renderPdfPage(pageNumber: number) {
    if (!this.previewCanvasRef.current) {
      return;
    }

    const pageInfo = this.state.pages[pageNumber - 1];
    if (!pageInfo) {
      return;
    }

    const pdf = this.pdfDocuments[pageInfo.sourceFileRef];
    if (!pdf) {
      return;
    }

    // Render the selected page in the preview canvas.
    const page = await pdf.getPage(pageInfo.sourcePageNumber);
    const scale = 1.5;
    const viewport = page.getViewport({ scale });

    const canvas = this.previewCanvasRef.current;
    const context = canvas.getContext('2d');
    if (!context) {
      throw new Error('Failed to get canvas context');
    }

    canvas.width = viewport.width;
    canvas.height = viewport.height;

    await page.render({ canvasContext: context, viewport }).promise;
  }

  // Move to the next or previous page.
  private async handlePageNavigation(pageNumber: number) {
    if (pageNumber < 1 || pageNumber > this.state.pages.length) {
      return;
    }

    this.setState({ currentPageNumber: pageNumber }, () => {
      this.renderPdfPage(pageNumber);
    });
  }

  // Merge selected pages and upload the new document.
  private handleCollateAndUpload = async () => {
    const { pages, newContractNumber, newDocumentType, selectedEntityKey, selectedEntitySiteUrl } = this.state;
    const { destinationLibraryTitle, destinationDocumentRepositoryTitle, sourceLibraryTitle, context } = this.props;
    const selectedPages = pages.filter(p => p.selected);

    if (selectedPages.length === 0 || !newContractNumber || !newDocumentType || !selectedEntityKey) {
      alert('Please select pages and enter contract number, document type, and entity.');
      return;
    }

    if (!destinationLibraryTitle) {
      alert('Merged library title is not configured. Please set it in the web part properties.');
      return;
    }

    if (!destinationDocumentRepositoryTitle) {
      alert('Destination document repository title is not configured. Please set it in the web part properties.');
      return;
    }

    if (!sourceLibraryTitle) {
      alert('Source library title is not configured. Please set it in the web part properties.');
      return;
    }

    if (!selectedEntitySiteUrl) {
      alert('Selected entity does not have a CMS site URL configured.');
      return;
    }

    const repositoryInfo = await getRepositoryFolderUrl(this.props.context, selectedEntitySiteUrl, destinationDocumentRepositoryTitle, newContractNumber);
    if (!repositoryInfo) {
      alert(`Contract Number '${newContractNumber}' was not found in the selected entity repository '${destinationDocumentRepositoryTitle}'.`);
      return;
    }

    this.setState({ uploading: true });

    try {
      // Prepare the source PDFs and build the merged output.
      const sourcePdfMap: { [fileRef: string]: PDFDocument } = {};
      const uniqueSourceFiles = selectedPages.reduce<{ [fileRef: string]: string }>((acc, pageInfo) => {
        acc[pageInfo.sourceFileRef] = pageInfo.sourceFileName;
        return acc;
      }, {});

      const newPdf = await PDFDocument.create();

      for (const pageInfo of selectedPages) {
        if (!sourcePdfMap[pageInfo.sourceFileRef]) {
          const response = await context.spHttpClient.get(pageInfo.sourceFileRef, SPHttpClient.configurations.v1);
          if (!response.ok) {
            throw new Error(`Failed to load source PDF ${pageInfo.sourceFileName}: ${response.status} ${response.statusText}`);
          }
          const arrayBuffer = await response.arrayBuffer();
          sourcePdfMap[pageInfo.sourceFileRef] = await PDFDocument.load(arrayBuffer);
        }

        const originalPdf = sourcePdfMap[pageInfo.sourceFileRef];
        const [copiedPage] = await newPdf.copyPages(originalPdf, [pageInfo.sourcePageNumber - 1]);
        newPdf.addPage(copiedPage);
      }

      // Save the merged PDF before uploading it.
      const pdfBytes = await newPdf.save();
      const timestamp = new Date().toISOString().replace(/[T:.]/g, '-').substring(0, 19);
      const fileName = `${newContractNumber}_${newDocumentType}_${timestamp}.pdf`;
      // Upload the merged file to the destination library.
      const uploadUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${safeODataString(destinationLibraryTitle)}')/RootFolder/Files/add(url='${encodeURIComponent(fileName)}',overwrite=true)`;

      const uploadResponse = await context.spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
        body: pdfBytes,
        headers: {
          'Content-Type': 'application/pdf',
          'Accept': 'application/json;odata=nometadata'
        }
      });

      if (!uploadResponse.ok) {
        throw new Error(`File upload failed: ${uploadResponse.status} ${uploadResponse.statusText}`);
      }

      const uploadResult = await uploadResponse.json();
      const serverRelativeUrl = uploadResult?.ServerRelativeUrl || uploadResult?.ServerRelativeUrlRaw || '';
      if (!serverRelativeUrl) {
        throw new Error('Uploaded file response did not include ServerRelativeUrl. Cannot update metadata.');
      }

      const fileUrlEncoded = encodeURIComponent(serverRelativeUrl);
      const metadataUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${fileUrlEncoded}')/ListItemAllFields`;

      const metadataResponse = await context.spHttpClient.post(metadataUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Content-Type': 'application/json;odata=nometadata',
          'Accept': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify({
          ContractNo: newContractNumber,
          DocumentType: newDocumentType
        })
      });

      if (!metadataResponse.ok) {
        throw new Error(`Metadata update failed: ${metadataResponse.status} ${metadataResponse.statusText}`);
      }

      const documentRepositoryUploadUrl = `${repositoryInfo.siteApiBase}/_api/web/GetFolderByServerRelativeUrl('${repositoryInfo.folderUrl.replace(/'/g, "''")}')/Files/add(url='${safeODataString(fileName)}',overwrite=true)`;
      const documentRepositoryUploadResponse = await context.spHttpClient.post(documentRepositoryUploadUrl, SPHttpClient.configurations.v1, {
        body: pdfBytes,
        headers: {
          'Content-Type': 'application/pdf',
          'Accept': 'application/json;odata=nometadata'
        }
      });

      if (!documentRepositoryUploadResponse.ok) {
        throw new Error(`Document repository upload failed: ${documentRepositoryUploadResponse.status} ${documentRepositoryUploadResponse.statusText}`);
      }

      const documentRepositoryUploadResult = await documentRepositoryUploadResponse.json();
      const documentRepositoryServerRelativeUrl = documentRepositoryUploadResult?.ServerRelativeUrl || `${repositoryInfo.folderUrl}/${fileName}`;
      const repositoryFileUrlEncoded = encodeURIComponent(documentRepositoryServerRelativeUrl);
      const repositoryMetadataUrl = `${repositoryInfo.siteApiBase}/_api/web/GetFileByServerRelativeUrl('${repositoryFileUrlEncoded}')/ListItemAllFields`;

      const repositoryMetadataResponse = await context.spHttpClient.post(repositoryMetadataUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Content-Type': 'application/json;odata=nometadata',
          'Accept': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify({
          DocumentNumber: newContractNumber,
          DocumentType: newDocumentType
        })
      });

      if (!repositoryMetadataResponse.ok) {
        throw new Error(`Document repository metadata update failed: ${repositoryMetadataResponse.status} ${repositoryMetadataResponse.statusText}`);
      }

      await checkInFile(context, documentRepositoryServerRelativeUrl, repositoryInfo.siteApiBase, 'Uploaded by SplitMerge');

      const currentUserId = context.pageContext.legacyPageContext?.userId;
      if (!currentUserId) {
        throw new Error('Unable to determine the current user ID for AssignedTo update.');
      }

      const sourceFileRefs = Object.keys(uniqueSourceFiles);
      const remainingFileRefs = new Set<string>();

      for (const sourceFileRef of sourceFileRefs) {
        const filePages = pages.filter(p => p.sourceFileRef === sourceFileRef);
        const remainingPages = filePages.filter(p => !p.selected);

        if (remainingPages.length === 0) {
          // Remove the source file when all pages are uploaded.
          const deleteUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(sourceFileRef)}')`;
          const deleteResponse = await context.spHttpClient.post(deleteUrl, SPHttpClient.configurations.v1, {
            headers: {
              'IF-MATCH': '*',
              'X-HTTP-Method': 'DELETE'
            }
          });

          if (!deleteResponse.ok) {
            throw new Error(`Source file delete failed for ${sourceFileRef}: ${deleteResponse.status} ${deleteResponse.statusText}`);
          }
        } else {
          remainingFileRefs.add(sourceFileRef);
          const remainingPdf = await PDFDocument.create();
          const originalPdf = sourcePdfMap[sourceFileRef];

          for (const pageInfo of remainingPages) {
            const [copiedPage] = await remainingPdf.copyPages(originalPdf, [pageInfo.sourcePageNumber - 1]);
            remainingPdf.addPage(copiedPage);
          }

          const remainingBytes = await remainingPdf.save();
          const fileName = uniqueSourceFiles[sourceFileRef];
          const sourceUploadUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${safeODataString(sourceLibraryTitle)}')/RootFolder/Files/add(url='${encodeURIComponent(fileName)}',overwrite=true)`;

          const overwriteResponse = await context.spHttpClient.post(sourceUploadUrl, SPHttpClient.configurations.v1, {
            body: remainingBytes,
            headers: {
              'Content-Type': 'application/pdf',
              'Accept': 'application/json;odata=nometadata'
            }
          });

          if (!overwriteResponse.ok) {
            throw new Error(`Source file overwrite failed for ${sourceFileRef}: ${overwriteResponse.status} ${overwriteResponse.statusText}`);
          }
        }
      }

      const remainingFileRefsArray: string[] = [];
      remainingFileRefs.forEach(ref => remainingFileRefsArray.push(ref));
      for (const sourceFileRef of remainingFileRefsArray) {
        const sourceMetadataUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(sourceFileRef)}')/ListItemAllFields`;
        const sourceMetadataResponse = await context.spHttpClient.post(sourceMetadataUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Content-Type': 'application/json;odata=nometadata',
            'Accept': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: JSON.stringify({
            AssignedToId: currentUserId
          })
        });

        if (!sourceMetadataResponse.ok) {
          throw new Error(`Source file assignment update failed for ${sourceFileRef}: ${sourceMetadataResponse.status} ${sourceMetadataResponse.statusText}`);
        }
      }

      alert('PDF collated and uploaded successfully.');
      await this.loadPdfFiles();
      this.setState({ uploading: false, pages: [], selectedPdfFiles: [], selectedPdfName: '', showModal: false, newContractNumber: '', newDocumentType: '', errorMessage: '' });
      this.handleModalClose();
    } catch (error) {
      console.error('Error collating, uploading, or updating source PDF:', error);
      alert('Error occurred. Please try again.');
      this.setState({ uploading: false });
    }
  }

  // Draw the UI for the split and merge workflow.
  public render(): React.ReactElement<ISplitMergeProps> {
    const { pdfFiles, pages, loading, newContractNumber, newDocumentType, uploading, uploadingSource, errorMessage, entityOptions, loadingEntities, selectedEntityKey, selectedPdfFiles, showAutoClassifyModal, selectedAutoClassifyFile } = this.state;
    const isMergeUploading = uploading;

    const allPagesSelected = pages.length > 0 && pages.every(p => p.selected);
    const somePagesSelected = pages.some(p => p.selected) && !allPagesSelected;

    const renderCell = (item: any, field: string) => {
      const color = item.azureResponse && String(item.azureResponse).trim() !== '' ? undefined : 'grey';
      return <span style={{ color }}>{item[field]}</span>;
    };

    const columns: IColumn[] = [
      {
        key: 'fileName',
        name: 'File Name',
        fieldName: 'fileName',
        minWidth: 200,
        maxWidth: 300,
        onRender: (item) => (
          <Link
            onClick={(event) => {
              event.preventDefault();
              this.handlePdfSelect(item.fileRef, item.fileName);
            }}
            disabled={loading}
          >
            <span style={{ color: item.azureResponse && String(item.azureResponse).trim() !== '' ? undefined : 'grey' }}>{item.fileName}</span>
          </Link>
        )
      },
      {
        key: 'created',
        name: 'Created',
        fieldName: 'created',
        minWidth: 150,
        maxWidth: 200
      },
      {
        key: 'modified',
        name: 'Modified',
        fieldName: 'modified',
        minWidth: 150,
        maxWidth: 200
      },
      {
        key: 'author',
        name: 'Author',
        fieldName: 'author',
        minWidth: 100,
        maxWidth: 150
      },
      {
        key: 'assignedTo',
        name: 'Assigned To',
        fieldName: 'assignedTo',
        minWidth: 150,
        maxWidth: 200
      }
    ];

    // Add the AutoClassifyStatus column and ensure items include AzureResponse
    columns.push({
      key: 'autoClassifyStatus',
      name: 'AutoClassifyStatus',
      fieldName: 'autoClassifyStatus',
      minWidth: 150,
      maxWidth: 200,
      onRender: (item) => renderCell(item, 'autoClassifyStatus')
    });

    // Ensure other columns render with grey when AzureResponse is missing
    // Add onRender to created/modified/author/assignedTo columns
    columns.forEach(col => {
      if (!col.onRender) {
        col.onRender = (item: any) => renderCell(item, col.fieldName || '');
      }
    });

    const items = pdfFiles.map(file => ({
      fileName: file.FileLeafRef,
      fileRef: file.FileRef,
      created: new Date(file.Created).toLocaleString(),
      modified: new Date(file.Modified).toLocaleString(),
      author: file.Author?.Title || 'Unknown',
      assignedTo: file.AssignedTo?.Title || 'Unassigned',
      azureResponse: file.AzureResponse || '',
      autoClassifyStatus: file.AutoClassifyStatus || ''
    }));

    return (
      <section className={`${styles.splitMerge} ${this.props.hasTeamsContext ? styles.teams : ''}`}>
        <div>
          {this.props.title && <h3>{this.props.title}</h3>}
          <div style={{ marginBottom: 16, display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
            <PrimaryButton
              text="Upload New PDF"
              onClick={this.handleUploadButtonClick}
              disabled={!this.props.sourceLibraryTitle || uploadingSource || isMergeUploading}
            />
            <PrimaryButton
              text="Manual Split & Classify"
              onClick={this.handleOpenSelectedClick}
              disabled={selectedPdfFiles.length === 0 || loading || isMergeUploading}
            />
            <PrimaryButton
              text="Auto Split & Classify"
              onClick={this.handleOpenClassifyClick}
              disabled={selectedPdfFiles.length !== 1 || loading || this.state.disableAutoClassify || isMergeUploading}
            />
            <input
              ref={this.fileInputRef}
              type="file"
              accept="application/pdf"
              style={{ display: 'none' }}
              onChange={this.handleFileInputChange}
            />
            {uploadingSource && <Spinner size={SpinnerSize.small} label="Uploading PDF..." />}
          </div>
          <div>
            {errorMessage && <div style={{ color: 'red' }}>{errorMessage}</div>}
            {pdfFiles.length === 0 && !errorMessage ? (
              <p>No PDF files found or library not specified.</p>
            ) : (
              <DetailsList
                items={items}
                columns={columns}
                selection={this.selection}
                selectionMode={isMergeUploading ? SelectionMode.none : SelectionMode.multiple}
                setKey="pdfFiles"
              />
            )}
          </div>
          {loading && <Spinner size={SpinnerSize.medium} label="Loading PDF..." />}
        </div>

        <Modal
          isOpen={this.state.showModal}
          onDismiss={this.handleModalClose}
          isBlocking={isMergeUploading}
          containerClassName={styles.modalContainer}
        >
          <div className={styles.modalHeader}>
            <h3>Select Pages from: {this.state.selectedPdfName || 'selected PDF(s)'}</h3>
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              onClick={this.handleModalClose}
              title="Close"
              disabled={isMergeUploading}
            />
          </div>
          <div className={styles.modalBody}>
            {pages.length > 0 && (
              <div className={styles.modalContent}>
                <div className={styles.previewPanel}>
                  <div className={styles.previewControls}>
                    <PrimaryButton
                      text="Previous"
                      onClick={() => this.handlePageNavigation(this.state.currentPageNumber - 1)}
                      disabled={this.state.currentPageNumber <= 1 || isMergeUploading}
                    />
                    <Label className={styles.currentPageLabel}>
                        Page {this.state.currentPageNumber} of {pages.length}
                        {this.state.pages.length > 0 && (() => {
                          const currentPage = this.state.pages[this.state.currentPageNumber - 1];
                          return currentPage ? ` (${currentPage.sourceFileName} page ${currentPage.sourcePageNumber})` : '';
                        })()}
                      </Label>
                    <PrimaryButton
                      text="Next"
                      onClick={() => this.handlePageNavigation(this.state.currentPageNumber + 1)}
                      disabled={this.state.currentPageNumber >= pages.length || isMergeUploading}
                    />
                  </div>
                  <div className={styles.previewCanvasWrapper}>
                    <canvas ref={this.previewCanvasRef} className={styles.previewCanvas} />
                  </div>
                </div>
                <div className={styles.selectionPanel}>
                  <Label>Page Selection</Label>
                  <Checkbox
                    label="Select All"
                    checked={allPagesSelected}
                    indeterminate={somePagesSelected}
                    disabled={isMergeUploading}
                    onChange={(ev, checked) => this.setState(prevState => ({
                      pages: prevState.pages.map(page => ({ ...page, selected: checked || false }))
                    }))}
                  />
                  <div className={styles.pageSelectionList}>
                    {pages.map((page, index) => (
                      <div key={page.id} className={styles.pageSelectionItem}>
                        <Checkbox
                          label={`Page ${page.sourcePageNumber} from ${page.sourceFileName}`}
                          checked={page.selected}
                          disabled={isMergeUploading}
                          onChange={(ev, checked) => this.handlePageSelect(page.id, checked || false)}
                        />
                        <PrimaryButton
                          text="Preview"
                          onClick={() => this.handlePageNavigation(index + 1)}
                          disabled={this.state.currentPageNumber === index + 1 || isMergeUploading}
                        />
                      </div>
                    ))}
                  </div>
                  <div className={styles.formSection}>
                    <TextField
                      label="Contract Number"
                      value={newContractNumber}
                      onChange={(ev, value) => this.setState({ newContractNumber: value || '' })}
                      required
                      disabled={isMergeUploading}
                    />
                    <Dropdown
                      label="Document Type"
                      options={this.state.documentTypes}
                      selectedKey={newDocumentType || undefined}
                      onChange={(ev, option) => this.setState({ newDocumentType: option?.key as string || '' })}
                      required
                      disabled={this.state.loadingDocumentTypes || this.state.documentTypes.length === 0 || isMergeUploading}
                      placeholder={this.state.loadingDocumentTypes ? "Loading document types..." : "Select a document type"}
                    />
                    <Dropdown
                      label="Entity"
                      options={entityOptions}
                      selectedKey={selectedEntityKey || undefined}
                      onChange={(ev, option) => this.setState({ selectedEntityKey: option?.key as string || '', selectedEntitySiteUrl: option?.data as string || '' })}
                      required
                      disabled={loadingEntities || entityOptions.length === 0 || isMergeUploading}
                      placeholder={loadingEntities ? "Loading entities..." : "Select an entity"}
                    />
                    <PrimaryButton
                      text="Merge and Upload"
                      onClick={() => this.handleCollateAndUpload()}
                      disabled={uploading || !newContractNumber || !newDocumentType || !selectedEntityKey || pages.filter(p => p.selected).length === 0}
                    />
                    {uploading && <Spinner size={SpinnerSize.small} label="Uploading..." />}
                  </div>
                </div>
              </div>
            )}
          </div>
        </Modal>

        <AutoClassifyModal
          isOpen={showAutoClassifyModal}
          onDismiss={this.handleAutoClassifyDismiss}
          selectedPdfFile={selectedAutoClassifyFile}
          entityOptions={this.state.entityOptions}
          sourceLibraryTitle={this.props.sourceLibraryTitle}
          destinationLibraryTitle={this.props.destinationLibraryTitle}
          destinationDocumentRepositoryTitle={this.props.destinationDocumentRepositoryTitle}
          context={this.props.context}
          isBusy={isMergeUploading}
          onUploadSuccess={async () => {
            await this.loadPdfFiles();
          }}
        />
      </section>
    );
  }
}