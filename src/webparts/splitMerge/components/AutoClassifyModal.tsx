import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { PrimaryButton, TextField, Checkbox, Label, Spinner, SpinnerSize, Modal, IconButton, Dropdown } from '@fluentui/react';
import * as pdfjsLib from 'pdfjs-dist';
import { PDFDocument } from 'pdf-lib';
import styles from './SplitMerge.module.scss';
import { checkInFile, getRepositoryFolderUrl, getServerRelativeUrl, parseClassificationResults, sanitizeFileNamePart, safeODataString } from './splitMergeHelpers';
import { IAutoClassifyModalProps, IClassificationResult, IPdfSelection, IPageInfo } from './splitMergeTypes';

pdfjsLib.GlobalWorkerOptions.workerSrc = require('pdfjs-dist/build/pdf.worker.min.js');

const AutoClassifyModal: React.FC<IAutoClassifyModalProps> = ({
  isOpen,
  onDismiss,
  selectedPdfFile,
  entityOptions,
  sourceLibraryTitle,
  destinationLibraryTitle,
  destinationDocumentRepositoryTitle,
  context,
  isBusy = false,
  onUploadSuccess
}) => {
  const [pages, setPages] = useState<IPageInfo[]>([]);
  const [currentPageNumber, setCurrentPageNumber] = useState(1);
  const [loading, setLoading] = useState(false);
  
  const [classificationLoading, setClassificationLoading] = useState(false);
  const [classificationError, setClassificationError] = useState('');
  const [classificationResults, setClassificationResults] = useState<IClassificationResult[]>([]);
  const [selectedDetectedDocumentType, setSelectedDetectedDocumentType] = useState('');
  const [newContractNumber, setNewContractNumber] = useState('');
  const [selectedEntityKey, setSelectedEntityKey] = useState('');
  const [selectedEntitySiteUrl, setSelectedEntitySiteUrl] = useState('');
  const [uploading, setUploading] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');
  const [pdfDocument, setPdfDocument] = useState<any>(null);
  const isModalBusy = isBusy || loading || classificationLoading || uploading;
  
  const previewCanvasRef = useRef<HTMLCanvasElement>(null);

  useEffect(() => {
    if (isOpen && selectedPdfFile) {
      (async () => {
        const bytes = await loadPdf(selectedPdfFile);
        try {
          await classifySelectedDocument(bytes);
        } catch (e) {
          // error handled inside classifySelectedDocument
        }
      })();
    } else {
      resetState();
    }
  }, [isOpen, selectedPdfFile]);

  useEffect(() => {
    if (isOpen && pages.length > 0) {
      renderPdfPage(currentPageNumber).catch(error => {
        console.error('Error rendering PDF page in AutoClassifyModal:', error);
      });
    }
  }, [isOpen, pages, currentPageNumber]);

  // Reset the modal before a new run.
  const resetState = () => {
    setPages([]);
    setCurrentPageNumber(1);
    setLoading(false);
    setClassificationLoading(false);
    setClassificationError('');
    setClassificationResults([]);
    setSelectedDetectedDocumentType('');
    
    setNewContractNumber('');
    setSelectedEntityKey('');
    setSelectedEntitySiteUrl('');
    setUploading(false);
    setErrorMessage('');
    setPdfDocument(null);
  };

  // Load the PDF and build the page list.
  const loadPdf = async (selectedFile: IPdfSelection): Promise<ArrayBuffer | null> => {
    setLoading(true);
    setErrorMessage('');
    setPages([]);
    setPdfDocument(null);

    try {
      // Fetch the PDF file from SharePoint.
      const response = await context.spHttpClient.get(selectedFile.fileRef, SPHttpClient.configurations.v1);
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

      // Confirm the file is a real PDF before continuing.
      const pdf = await pdfjsLib.getDocument({ data: pdfBytesForPdfJs }).promise;
      setPdfDocument(pdf);

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

      setPages(loadedPages);
      setCurrentPageNumber(1);

      return pdfBytes;
    } catch (error) {
      console.error('Error loading PDF for classification:', error);
      setErrorMessage(error instanceof Error ? error.message : 'Error loading PDF.');
      return null;
    } finally {
      setLoading(false);
    }
  };

  // Show the selected page in the preview.
  const renderPdfPage = async (pageNumber: number) => {
    if (!previewCanvasRef.current || !pdfDocument) {
      return;
    }

    const pageInfo = pages[pageNumber - 1];
    if (!pageInfo) {
      return;
    }

    // Render the requested page on the preview canvas.
    const page = await pdfDocument.getPage(pageInfo.sourcePageNumber);
    const scale = 1.5;
    const viewport = page.getViewport({ scale });

    const canvas = previewCanvasRef.current;
    const context2d = canvas.getContext('2d');
    if (!context2d) {
      throw new Error('Failed to get canvas context');
    }

    canvas.width = viewport.width;
    canvas.height = viewport.height;

    await page.render({ canvasContext: context2d, viewport }).promise;
  };

  // Move to the next or previous page.
  const handlePageNavigation = (pageNumber: number) => {
    if (pageNumber < 1 || pageNumber > pages.length) {
      return;
    }
    setCurrentPageNumber(pageNumber);
  };

  // Mark a page as selected or not.
  const handlePageSelect = (pageId: string, selected: boolean) => {
    setPages(prev => prev.map(page => page.id === pageId ? { ...page, selected } : page));
  };

  // Switch the selected document type.
  const handleDetectedDocumentTypeChange = (selectedKey: string) => {
    let result: IClassificationResult | undefined;
    for (let i = 0; i < classificationResults.length; i++) {
      if (classificationResults[i].key === selectedKey) {
        result = classificationResults[i];
        break;
      }
    }

    if (!result) {
      setSelectedDetectedDocumentType('');
      setCurrentPageNumber(1);
      return;
    }

    setSelectedDetectedDocumentType(selectedKey);
    
    // Navigate to the first page of the selected document type
    if (result.pageNumbers.length > 0) {
      const firstPageNumber = result.pageNumbers[0];
      let pageIndex = -1;
      for (let i = 0; i < pages.length; i++) {
        if (pages[i].sourcePageNumber === firstPageNumber) {
          pageIndex = i;
          break;
        }
      }
      if (pageIndex !== -1) {
        setCurrentPageNumber(pageIndex + 1);
      }
    }
  };

  // Read the stored result and mark detected pages.
  const classifySelectedDocument = async (bytes?: ArrayBuffer | null) => {
    if (!selectedPdfFile) {
      setClassificationError('No file selected for classification.');
      return;
    }
    // When reading stored `AzureResponse` we do not need the PDF bytes.

    // New flow: read stored Azure Document Intelligence response from the file's
    // `AzureResponse` column in the same library instead of calling the Azure Function.
    setClassificationLoading(true);
    setClassificationError('');
    setClassificationResults([]);
    setSelectedDetectedDocumentType('');

    try {
      // Read the stored Azure response for this file.
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

      // Parse the classification result into document groups.
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

      setClassificationResults(results);
      setSelectedDetectedDocumentType(firstKey);
      setPages(prev => prev.map(page => ({
        ...page,
        selected: detectedPageNumbers.has(page.sourcePageNumber)
      })));
    } catch (error) {
      console.error('Classification error:', error);
      const rawMessage = error instanceof Error ? error.message : String(error);
      setClassificationError(rawMessage || 'Unknown classification error');
    } finally {
      setClassificationLoading(false);
    }
  };

  

  // Create merged files and upload them.
  const handleCollateAndUpload = async () => {
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

    const repositoryInfo = await getRepositoryFolderUrl(context, selectedEntitySiteUrl, destinationDocumentRepositoryTitle, contractNumber);
    if (!repositoryInfo) {
      alert(`Contract Number '${contractNumber}' was not found in the selected entity repository '${destinationDocumentRepositoryTitle}'.`);
      return;
    }

    setUploading(true);

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

      const detectedGroups = classificationResults.map(result => {
        const pageNumbers = new Set<number>(result.pageNumbers);
        const selectedGroupPages = selectedPages.filter(page => pageNumbers.has(page.sourcePageNumber));
        return {
          result,
          pages: selectedGroupPages
        };
      }).filter(group => group.pages.length > 0);

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
          // Build a new PDF from the selected pages.
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
              ContractNo: contractNumber,
              DocumentType: group.result.text
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
          // Remove the source file when nothing is left to keep.
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
      handleDismiss();
    } catch (error) {
      console.error('Error collating, uploading, or updating source PDF:', error);
      alert('Error occurred. Please try again.');
      setUploading(false);
    }
  };

  // Close the modal and reset the state.
  const handleDismiss = () => {
    resetState();
    onDismiss();
  };

  const selectedDocumentTypeResult = (() => {
    let result: IClassificationResult | undefined;
    for (let i = 0; i < classificationResults.length; i++) {
      if (classificationResults[i].key === selectedDetectedDocumentType) {
        result = classificationResults[i];
        break;
      }
    }
    return result;
  })();

  const detectedPageNumbers = selectedDocumentTypeResult ? selectedDocumentTypeResult.pageNumbers : [];
  const selectedDocPages = pages.filter(page => detectedPageNumbers.indexOf(page.sourcePageNumber) !== -1);

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={isModalBusy ? undefined : handleDismiss}
      isBlocking={isModalBusy}
      containerClassName={styles.modalContainer}
    >
      <div className={styles.modalHeader}>
        <h3>Auto Classify Document: {selectedPdfFile?.fileName || 'selected PDF'}</h3>
        <IconButton iconProps={{ iconName: 'Cancel' }} onClick={handleDismiss} title="Close" disabled={isModalBusy} />
      </div>
      <div className={styles.modalBody}>
        <div className={styles.modalContent}>
          <div className={styles.previewPanel}>
            <div className={styles.previewControls}>
              <PrimaryButton text="Previous" onClick={() => handlePageNavigation(currentPageNumber - 1)} disabled={currentPageNumber <= 1 || isModalBusy} />
              <Label className={styles.currentPageLabel}>
                Page {currentPageNumber} of {pages.length}
                {pages.length > 0 && (() => {
                  const currentPage = pages[currentPageNumber - 1];
                  return currentPage ? ` (${currentPage.sourceFileName} page ${currentPage.sourcePageNumber})` : '';
                })()}
              </Label>
              <PrimaryButton text="Next" onClick={() => handlePageNavigation(currentPageNumber + 1)} disabled={currentPageNumber >= pages.length || isModalBusy} />
            </div>
            <div className={styles.previewCanvasWrapper}>
              <canvas ref={previewCanvasRef} className={styles.previewCanvas} />
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
                  onChange={(ev, option) => handleDetectedDocumentTypeChange(option?.key as string)}
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
                        onChange={(ev, checked) => handlePageSelect(page.id, checked || false)}
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
                            handlePageNavigation(pageIndex + 1);
                          }
                        }}
                        disabled={isModalBusy}
                      />
                    </div>
                  ))}
                </div>

                <div className={styles.formSection}>
                  <TextField label="Contract Number" value={newContractNumber} onChange={(ev, value) => setNewContractNumber(value || '')} required disabled={isModalBusy} />
                  <Dropdown
                    label="Entity"
                    options={entityOptions}
                    selectedKey={selectedEntityKey || undefined}
                    onChange={(ev, option) => {
                      setSelectedEntityKey(option?.key as string || '');
                      setSelectedEntitySiteUrl(option?.data as string || '');
                    }}
                    required
                    disabled={entityOptions.length === 0 || isModalBusy}
                    placeholder={entityOptions.length === 0 ? 'No entities available' : 'Select an entity'}
                  />
                  <PrimaryButton
                    text="Merge and Upload"
                    onClick={handleCollateAndUpload}
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
};

export default AutoClassifyModal;
