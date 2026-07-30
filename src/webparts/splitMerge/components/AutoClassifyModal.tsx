import * as React from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PrimaryButton, TextField, Checkbox, Label, Spinner, SpinnerSize, Modal, IconButton, Dropdown } from '@fluentui/react';
import * as pdfjsLib from 'pdfjs-dist';
import { PDFDocument } from 'pdf-lib';
import styles from './SplitMerge.module.scss';
import { checkInFile, getRepositoryFolderUrl, getServerRelativeUrl, parseClassificationResults, sanitizeFileNamePart, safeODataString } from './splitMergeHelpers';
import { IAutoClassifyModalProps, IClassificationResult, IPdfSelection, IPageInfo } from './splitMergeTypes';

pdfjsLib.GlobalWorkerOptions.workerSrc = require('pdfjs-dist/build/pdf.worker.min.js');

interface IAutoClassifyModalState {
  pages: IPageInfo[];
  currentPageNumber: number;
  loading: boolean;
  classificationLoading: boolean;
  classificationError: string;
  classificationResults: IClassificationResult[];
  selectedDetectedDocumentType: string;
  newContractNumber: string;
  selectedEntityKey: string;
  selectedEntitySiteUrl: string;
  uploading: boolean;
  errorMessage: string;
}

export default class AutoClassifyModal extends React.Component<IAutoClassifyModalProps, IAutoClassifyModalState> {
  private previewCanvasRef = React.createRef<HTMLCanvasElement>();
  private pdfDocument: any = null;

  // Sets the initial modal state.
  constructor(props: IAutoClassifyModalProps) {
    super(props);
    this.state = {
      pages: [],
      currentPageNumber: 1,
      loading: false,
      classificationLoading: false,
      classificationError: '',
      classificationResults: [],
      selectedDetectedDocumentType: '',
      newContractNumber: '',
      selectedEntityKey: '',
      selectedEntitySiteUrl: '',
      uploading: false,
      errorMessage: ''
    };
  }

  // Starts the PDF flow when the modal opens.
  componentDidMount() {
    if (this.props.isOpen && this.props.selectedPdfFile) {
      void this.openSelectedPdfForClassification();
    }
  }

  // Refreshes the view when the modal state changes.
  componentDidUpdate(prevProps: Readonly<IAutoClassifyModalProps>, prevState: Readonly<IAutoClassifyModalState>) {
    const selectedFileChanged = prevProps.selectedPdfFile?.fileRef !== this.props.selectedPdfFile?.fileRef;

    if ((!prevProps.isOpen && this.props.isOpen) || (this.props.isOpen && selectedFileChanged)) {
      if (this.props.selectedPdfFile) {
        void this.openSelectedPdfForClassification();
      }
    }

    if (this.props.isOpen && prevState.currentPageNumber !== this.state.currentPageNumber) {
      void this.renderPdfPage(this.state.currentPageNumber).catch(error => {
        console.error('Error rendering PDF page in AutoClassifyModal:', error);
      });
    }
  }

  // Clears the current PDF and resets the form state.
  private resetState = () => {
    this.pdfDocument = null;
    this.setState({
      pages: [],
      currentPageNumber: 1,
      loading: false,
      classificationLoading: false,
      classificationError: '',
      classificationResults: [],
      selectedDetectedDocumentType: '',
      newContractNumber: '',
      selectedEntityKey: '',
      selectedEntitySiteUrl: '',
      uploading: false,
      errorMessage: ''
    });
  };

  // Loads and classifies the selected PDF file.
  private async openSelectedPdfForClassification() {
    const { selectedPdfFile } = this.props;
    if (!selectedPdfFile) {
      // Reset the modal when no file is available.
      this.resetState();
      return;
    }

    // Clear the previous state before loading a new PDF.
    this.resetState();
    const bytes = await this.loadPdf(selectedPdfFile);
    if (bytes) {
      await this.classifySelectedDocument();
    }
  }

  // Reads the PDF file and builds the page list.
  private async loadPdf(selectedFile: IPdfSelection): Promise<ArrayBuffer | null> {
    this.setState({ loading: true, errorMessage: '', pages: [] });
    this.pdfDocument = null;

    try {
      // Load the PDF bytes
      const response = await this.props.context.spHttpClient.get(selectedFile.fileRef, SPHttpClient.configurations.v1);
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      const arrayBuffer = await response.arrayBuffer();
      if (arrayBuffer.byteLength === 0) {
        throw new Error('PDF file is empty');
      }

      const pdfBytes = arrayBuffer.slice(0);
      const pdfBytesForPdfJs = pdfBytes.slice(0);
      const uint8Array = new Uint8Array(pdfBytes);
      if (uint8Array.length < 4 || uint8Array[0] !== 37 || uint8Array[1] !== 80 || uint8Array[2] !== 68 || uint8Array[3] !== 70) {
        throw new Error('File is not a valid PDF');
      }

      // Parse the file so the preview can render pages.
      const pdf = await pdfjsLib.getDocument({ data: pdfBytesForPdfJs }).promise;
      this.pdfDocument = pdf;

      const loadedPages: IPageInfo[] = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        loadedPages.push({
          id: `${selectedFile.fileRef}|${i}`,
          sourceFileRef: selectedFile.fileRef,
          sourceFileName: selectedFile.fileName,
          sourcePageNumber: i,
          selected: false
        });
      }

      this.setState(
        {
          pages: loadedPages,
          currentPageNumber: 1
        },
        () => {
          if (this.props.isOpen) {
            void this.renderPdfPage(1).catch(error => {
              console.error('Error rendering PDF page in AutoClassifyModal:', error);
            });
          }
        }
      );

      return pdfBytes;
    } catch (error) {
      console.error('Error loading PDF for classification:', error);
      this.setState({ errorMessage: error instanceof Error ? error.message : 'Error loading PDF.' });
      return null;
    } finally {
      this.setState({ loading: false });
    }
  }

  // Draws the selected PDF page in the preview canvas.
  private async renderPdfPage(pageNumber: number) {
    if (!this.previewCanvasRef.current || !this.pdfDocument) {
      return;
    }

    const pageInfo = this.state.pages[pageNumber - 1];
    if (!pageInfo) {
      return;
    }

    // Get the requested page from the PDF document.
    const page = await this.pdfDocument.getPage(pageInfo.sourcePageNumber);
    const scale = 1.5;
    const viewport = page.getViewport({ scale });

    const canvas = this.previewCanvasRef.current;
    const context2d = canvas.getContext('2d');
    if (!context2d) {
      throw new Error('Failed to get canvas context');
    }

    canvas.width = viewport.width;
    canvas.height = viewport.height;

    // Draw the page onto the preview canvas.
    await page.render({ canvasContext: context2d, viewport }).promise;
  }

  // Moves to the previous or next preview page.
  private handlePageNavigation = (pageNumber: number) => {
    if (pageNumber < 1 || pageNumber > this.state.pages.length) {
      return;
    }

    this.setState({ currentPageNumber: pageNumber }, () => {
      void this.renderPdfPage(pageNumber).catch(error => {
        console.error('Error rendering PDF page in AutoClassifyModal:', error);
      });
    });
  };

  // Marks a page as selected or unselected.
  private handlePageSelect = (pageId: string, selected: boolean) => {
    this.setState(prevState => ({
      pages: prevState.pages.map(page => (page.id === pageId ? { ...page, selected } : page))
    }));
  };

  // Switches the preview to the selected detected document type.
  private handleDetectedDocumentTypeChange = (selectedKey: string) => {
    let result: IClassificationResult | undefined;
    // Find the selected document type in the results.
    for (let i = 0; i < this.state.classificationResults.length; i++) {
      if (this.state.classificationResults[i].key === selectedKey) {
        result = this.state.classificationResults[i];
        break;
      }
    }

    if (!result) {
      this.setState({ selectedDetectedDocumentType: '', currentPageNumber: 1 }, () => {
        void this.renderPdfPage(1).catch(error => {
          console.error('Error rendering PDF page in AutoClassifyModal:', error);
        });
      });
      return;
    }

    if (result.pageNumbers.length > 0) {
      const firstPageNumber = result.pageNumbers[0];
      let pageIndex = -1;
      for (let i = 0; i < this.state.pages.length; i++) {
        if (this.state.pages[i].sourcePageNumber === firstPageNumber) {
          pageIndex = i;
          break;
        }
      }

      if (pageIndex !== -1) {
        this.setState(
          {
            selectedDetectedDocumentType: selectedKey,
            currentPageNumber: pageIndex + 1
          },
          () => {
            void this.renderPdfPage(pageIndex + 1).catch(error => {
              console.error('Error rendering PDF page in AutoClassifyModal:', error);
            });
          }
        );
        return;
      }
    }

    this.setState({ selectedDetectedDocumentType: selectedKey });
  };

  // Retrieves classification results for the selected PDF.
  private classifySelectedDocument = async () => {
    const { selectedPdfFile, context } = this.props;
    if (!selectedPdfFile) {
      this.setState({ classificationError: 'No file selected for classification.' });
      return;
    }

    this.setState({
      classificationLoading: true,
      classificationError: '',
      classificationResults: [],
      selectedDetectedDocumentType: ''
    });

    try {
      // Build the SharePoint metadata URL for the file.
      const serverRelative = getServerRelativeUrl(selectedPdfFile.fileRef);
      if (!serverRelative) {
        throw new Error('Unable to determine server-relative URL for the selected file.');
      }

      const fileUrlEncoded = encodeURIComponent(serverRelative);
      const listItemUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${fileUrlEncoded}')/ListItemAllFields?$select=AzureResponse`;

      const metaResponse = await context.spHttpClient.get(listItemUrl, SPHttpClient.configurations.v1);
      if (!metaResponse.ok) {
        const body = await metaResponse.text();
        throw new Error(`Failed to retrieve AzureResponse metadata: ${metaResponse.status} ${metaResponse.statusText} - ${body}`);
      }

      const metaJson = await metaResponse.json();
      const azureResponseText = metaJson?.AzureResponse || '';
      if (!azureResponseText || typeof azureResponseText !== 'string' || azureResponseText.trim() === '') {
        throw new Error('AzureResponse column is empty for this file.');
      }

      let resultJson: any = null;
      try {
        resultJson = JSON.parse(azureResponseText);
      } catch (e) {
        throw new Error('AzureResponse contains invalid JSON.');
      }

      const results = parseClassificationResults(resultJson);
      if (results.length === 0) {
        throw new Error('No document types were detected by the stored Document Intelligence response.');
      }

      const firstKey = results[0].key;
      const detectedPageNumbers = new Set<number>();
      for (let i = 0; i < results.length; i++) {
        for (let j = 0; j < results[i].pageNumbers.length; j++) {
          detectedPageNumbers.add(results[i].pageNumbers[j]);
        }
      }

      // Mark the detected pages as selected.
      this.setState({
        classificationResults: results,
        selectedDetectedDocumentType: firstKey,
        pages: this.state.pages.map(page => ({
          ...page,
          selected: detectedPageNumbers.has(page.sourcePageNumber)
        }))
      });
    } catch (error) {
      console.error('Classification error:', error);
      const rawMessage = error instanceof Error ? error.message : String(error);
      this.setState({ classificationError: rawMessage || 'Unknown classification error' });
    } finally {
      this.setState({ classificationLoading: false });
    }
  };

  // Merges selected pages and uploads them to the target locations.
  private handleCollateAndUpload = async () => {
    const { destinationLibraryTitle, destinationDocumentRepositoryTitle, sourceLibraryTitle, context, onUploadSuccess } = this.props;
    const { pages, newContractNumber, selectedEntityKey, selectedEntitySiteUrl, classificationResults } = this.state;
    // Only use the pages chosen by the user.
    const selectedPages = pages.filter(p => p.selected);
    const contractNumber = newContractNumber.trim();

    if (selectedPages.length === 0 || !contractNumber || !selectedEntityKey) {
      alert('Please select pages and enter contract number and entity.');
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

    // Resolve the destination folder for the contract number.
    const repositoryInfo = await getRepositoryFolderUrl(context, selectedEntitySiteUrl, destinationDocumentRepositoryTitle, contractNumber);
    if (!repositoryInfo) {
      alert(`Contract Number '${contractNumber}' was not found in the selected entity repository '${destinationDocumentRepositoryTitle}'.`);
      return;
    }

    this.setState({ uploading: true });

    try {
      const sourcePdfMap: { [fileRef: string]: PDFDocument } = {};
      const uniqueSourceFiles = pages.reduce<{ [fileRef: string]: string }>((acc, pageInfo) => {
        acc[pageInfo.sourceFileRef] = pageInfo.sourceFileName;
        return acc;
      }, {});

      for (let i = 0; i < pages.length; i++) {
        const pageInfo = pages[i];
        if (!sourcePdfMap[pageInfo.sourceFileRef]) {
          const response = await context.spHttpClient.get(pageInfo.sourceFileRef, SPHttpClient.configurations.v1);
          if (!response.ok) {
            throw new Error(`Failed to load source PDF ${pageInfo.sourceFileName}: ${response.status} ${response.statusText}`);
          }
          const arrayBuffer = await response.arrayBuffer();
          sourcePdfMap[pageInfo.sourceFileRef] = await PDFDocument.load(arrayBuffer);
        }
      }

      const detectedGroups = classificationResults
        .map(result => {
          const pageNumbers = new Set<number>(result.pageNumbers);
          const selectedGroupPages = selectedPages.filter(page => pageNumbers.has(page.sourcePageNumber));
          return {
            result,
            pages: selectedGroupPages
          };
        })
        .filter(group => group.pages.length > 0);

      if (detectedGroups.length === 0) {
        throw new Error('No selected pages match any detected document type.');
      }

      const timestamp = new Date().toISOString().replace(/[T:.]/g, '-').substring(0, 19);
      const successfulPageIds = new Set<string>();
      const successfulTypes: string[] = [];
      const failedTypes: string[] = [];

      for (let i = 0; i < detectedGroups.length; i++) {
        const group = detectedGroups[i];

        try {
          // Create a new PDF for each detected document type.
          const newPdf = await PDFDocument.create();
          for (let j = 0; j < group.pages.length; j++) {
            const pageInfo = group.pages[j];
            const originalPdf = sourcePdfMap[pageInfo.sourceFileRef];
            const [copiedPage] = await newPdf.copyPages(originalPdf, [pageInfo.sourcePageNumber - 1]);
            newPdf.addPage(copiedPage);
          }

          const pdfBytes = await newPdf.save();
          const cleanDocType = sanitizeFileNamePart(group.result.text);
          const cleanContractNumber = sanitizeFileNamePart(contractNumber);
          const fileName = `${cleanDocType}_${cleanContractNumber}_${timestamp}.pdf`;

          // Upload the merged PDF to the main library.
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
              ContractNo: contractNumber,
              DocumentType: group.result.text
            })
          });

          if (!metadataResponse.ok) {
            throw new Error(`Metadata update failed: ${metadataResponse.status} ${metadataResponse.statusText}`);
          }

          // Upload the same file to the document repository.
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
              DocumentNumber: contractNumber,
              DocumentType: group.result.text
            })
          });

          if (!repositoryMetadataResponse.ok) {
            throw new Error(`Document repository metadata update failed: ${repositoryMetadataResponse.status} ${repositoryMetadataResponse.statusText}`);
          }

          await checkInFile(context, documentRepositoryServerRelativeUrl, repositoryInfo.siteApiBase, 'Uploaded by AutoClassify');

          for (let j = 0; j < group.pages.length; j++) {
            successfulPageIds.add(group.pages[j].id);
          }
          successfulTypes.push(group.result.text);
        } catch (groupError) {
          console.error(`Upload failed for detected type ${group.result.text}:`, groupError);
          failedTypes.push(group.result.text);
        }
      }

      if (successfulPageIds.size === 0) {
        throw new Error('Upload failed for all detected document types.');
      }

      const currentUserId = context.pageContext.legacyPageContext?.userId;
      if (!currentUserId) {
        throw new Error('Unable to determine the current user ID for AssignedTo update.');
      }

      const sourceFileRefs = Object.keys(uniqueSourceFiles);
      const remainingFileRefs = new Set<string>();

      for (let i = 0; i < sourceFileRefs.length; i++) {
        const sourceFileRef = sourceFileRefs[i];
        const filePages = pages.filter(p => p.sourceFileRef === sourceFileRef);
        const remainingPages = filePages.filter(p => !successfulPageIds.has(p.id));

        if (remainingPages.length === 0) {
          // Delete the source file when no pages remain.
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
          // Keep the source file and upload the remaining pages.
          remainingFileRefs.add(sourceFileRef);
          const remainingPdf = await PDFDocument.create();
          const originalPdf = sourcePdfMap[sourceFileRef];

          for (let j = 0; j < remainingPages.length; j++) {
            const pageInfo = remainingPages[j];
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
      for (let i = 0; i < remainingFileRefsArray.length; i++) {
        const sourceFileRef = remainingFileRefsArray[i];
        const sourceMetadataUrl = `${context.pageContext.web.absoluteUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(sourceFileRef)}')/ListItemAllFields`;
        const sourceMetadataResponse = await context.spHttpClient.post(sourceMetadataUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Content-Type': 'application/json;odata=nometadata',
            'Accept': 'application/json;odata=nometadata',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: JSON.stringify({
            AssignedToId: currentUserId,
            AzureResponse: ''
          })
        });

        if (!sourceMetadataResponse.ok) {
          throw new Error(`Source file assignment update failed for ${sourceFileRef}: ${sourceMetadataResponse.status} ${sourceMetadataResponse.statusText}`);
        }
      }

      const summaryMessage = failedTypes.length === 0
        ? `Uploaded ${successfulTypes.length} detected document type file(s) successfully.`
        : `Uploaded ${successfulTypes.length} detected document type file(s). Failed types: ${failedTypes.join(', ')}.`;

      alert(summaryMessage);
      await onUploadSuccess();
      this.handleDismiss();
    } catch (error) {
      console.error('Error collating, uploading, or updating source PDF:', error);
      alert('Error occurred. Please try again.');
      this.setState({ uploading: false });
    }
  };

  // Closes the modal and clears the current work.
  private handleDismiss = () => {
    this.resetState();
    this.props.onDismiss();
  };

  // Finds the selected document type result from the classification list.
  private getSelectedDocumentTypeResult() {
    for (let i = 0; i < this.state.classificationResults.length; i++) {
      if (this.state.classificationResults[i].key === this.state.selectedDetectedDocumentType) {
        return this.state.classificationResults[i];
      }
    }

    return undefined;
  }

  // Renders the classification and upload UI.
  public render() {
    const { isOpen, entityOptions, selectedPdfFile, isBusy = false } = this.props;
    const {
      pages,
      currentPageNumber,
      loading,
      classificationLoading,
      classificationError,
      classificationResults,
      selectedDetectedDocumentType,
      newContractNumber,
      selectedEntityKey,
      uploading,
      errorMessage
    } = this.state;
    const isModalBusy = isBusy || loading || classificationLoading || uploading;
    const selectedDocumentTypeResult = this.getSelectedDocumentTypeResult();
    const detectedPageNumbers = selectedDocumentTypeResult ? selectedDocumentTypeResult.pageNumbers : [];
    const selectedDocPages = pages.filter(page => detectedPageNumbers.indexOf(page.sourcePageNumber) !== -1);

    return (
      <Modal
        isOpen={isOpen}
        onDismiss={isModalBusy ? undefined : this.handleDismiss}
        isBlocking={isModalBusy}
        containerClassName={styles.modalContainer}
      >
        <div className={styles.modalHeader}>
          <h3>Auto Classify Document: {selectedPdfFile?.fileName || 'selected PDF'}</h3>
          <IconButton iconProps={{ iconName: 'Cancel' }} onClick={this.handleDismiss} title="Close" disabled={isModalBusy} />
        </div>
        <div className={styles.modalBody}>
          <div className={styles.modalContent}>
            <div className={styles.previewPanel}>
              <div className={styles.previewControls}>
                <PrimaryButton text="Previous" onClick={() => this.handlePageNavigation(currentPageNumber - 1)} disabled={currentPageNumber <= 1 || isModalBusy} />
                <Label className={styles.currentPageLabel}>
                  Page {currentPageNumber} of {pages.length}
                  {pages.length > 0 && (() => {
                    const currentPage = pages[currentPageNumber - 1];
                    return currentPage ? ` (${currentPage.sourceFileName} page ${currentPage.sourcePageNumber})` : '';
                  })()}
                </Label>
                <PrimaryButton text="Next" onClick={() => this.handlePageNavigation(currentPageNumber + 1)} disabled={currentPageNumber >= pages.length || isModalBusy} />
              </div>
              <div className={styles.previewCanvasWrapper}>
                <canvas ref={this.previewCanvasRef} className={styles.previewCanvas} />
                {loading && <Spinner size={SpinnerSize.large} label="Loading PDF..." />}
              </div>
            </div>

            <div className={styles.selectionPanel}>
              <Label>Document Classification</Label>
              {classificationLoading && <Spinner size={SpinnerSize.small} label="Classifying document..." />}
              {classificationError && <div style={{ color: 'red' }}>{classificationError}</div>}
              {errorMessage && <div style={{ color: 'red' }}>{errorMessage}</div>}

              {classificationResults.length > 0 && (
                <>
                  <Dropdown
                    label="Detected Document Type"
                    options={classificationResults.map(result => ({ key: result.key, text: `${result.text} (${result.confidence.toFixed(2)})` }))}
                    selectedKey={selectedDetectedDocumentType || undefined}
                    onChange={(ev, option) => this.handleDetectedDocumentTypeChange(option?.key as string)}
                    placeholder="Select a detected document type"
                    disabled={isModalBusy}
                  />
                  {selectedDocumentTypeResult && (
                    <div style={{ padding: 12, border: '1px solid #e1e1e1', borderRadius: 4, background: '#fff' }}>
                      <p><strong>Type:</strong> {selectedDocumentTypeResult.text}</p>
                      <p><strong>Confidence:</strong> {selectedDocumentTypeResult.confidence.toFixed(2)}</p>
                      <p><strong>Pages:</strong> {selectedDocumentTypeResult.pageNumbers.join(', ')}</p>
                    </div>
                  )}

                  <Label>Pages for selected type</Label>
                  <div className={styles.pageSelectionList}>
                    {selectedDocPages.length === 0 ? (
                      <div>No pages found for the selected document type.</div>
                    ) : selectedDocPages.map(page => (
                      <div key={page.id} className={styles.pageSelectionItem}>
                        <Checkbox
                          label={`Page ${page.sourcePageNumber}`}
                          checked={page.selected}
                          disabled={isModalBusy}
                          onChange={(ev, checked) => this.handlePageSelect(page.id, checked || false)}
                        />
                        <PrimaryButton
                          text="Preview"
                          onClick={() => {
                            let pageIndex = -1;
                            for (let i = 0; i < pages.length; i++) {
                              if (pages[i].id === page.id) {
                                pageIndex = i;
                                break;
                              }
                            }

                            if (pageIndex !== -1) {
                              this.handlePageNavigation(pageIndex + 1);
                            }
                          }}
                          disabled={isModalBusy}
                        />
                      </div>
                    ))}
                  </div>

                  <div className={styles.formSection}>
                    <TextField label="Contract Number" value={newContractNumber} onChange={(ev, value) => this.setState({ newContractNumber: value || '' })} required disabled={isModalBusy} />
                    <Dropdown
                      label="Entity"
                      options={entityOptions}
                      selectedKey={selectedEntityKey || undefined}
                      onChange={(ev, option) => {
                        this.setState({
                          selectedEntityKey: option?.key as string || '',
                          selectedEntitySiteUrl: option?.data as string || ''
                        });
                      }}
                      required
                      disabled={entityOptions.length === 0 || isModalBusy}
                      placeholder={entityOptions.length === 0 ? 'No entities available' : 'Select an entity'}
                    />
                    <PrimaryButton
                      text="Merge and Upload"
                      onClick={this.handleCollateAndUpload}
                      disabled={uploading || isBusy || !newContractNumber.trim() || !selectedEntityKey || pages.filter(page => page.selected).length === 0}
                    />
                    {uploading && <Spinner size={SpinnerSize.small} label="Uploading..." />}
                  </div>
                </>
              )}
            </div>
          </div>
        </div>
      </Modal>
    );
  }
}
