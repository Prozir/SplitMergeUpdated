import * as React from 'react';
import { useEffect, useRef, useState } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { PrimaryButton, TextField, Checkbox, Label, Spinner, SpinnerSize, Modal, IconButton, Dropdown, IDropdownOption } from '@fluentui/react';
import * as pdfjsLib from 'pdfjs-dist';
import { PDFDocument } from 'pdf-lib';
import styles from './SplitMerge.module.scss';

pdfjsLib.GlobalWorkerOptions.workerSrc = require('pdfjs-dist/build/pdf.worker.min.js');

interface IPageInfo {
  id: string;
  sourceFileRef: string;
  sourceFileName: string;
  sourcePageNumber: number;
  selected: boolean;
}

interface IPdfSelection {
  fileRef: string;
  fileName: string;
}

interface IClassificationResult {
  key: string;
  text: string;
  confidence: number;
  pageNumbers: number[];
}

interface IAutoClassifyModalProps {
  isOpen: boolean;
  onDismiss: () => void;
  selectedPdfFile: IPdfSelection | null;
  documentTypes: IDropdownOption[];
  entityOptions: IDropdownOption[];
  sourceLibraryTitle: string;
  destinationLibraryTitle: string;
  destinationDocumentRepositoryTitle: string;
  context: WebPartContext;
  onUploadSuccess: () => Promise<void>;
}

const AutoClassifyModal: React.FC<IAutoClassifyModalProps> = ({
  isOpen,
  onDismiss,
  selectedPdfFile,
  documentTypes,
  entityOptions,
  sourceLibraryTitle,
  destinationLibraryTitle,
  destinationDocumentRepositoryTitle,
  context,
  onUploadSuccess
}) => {
  const [pages, setPages] = useState<IPageInfo[]>([]);
  const [currentPageNumber, setCurrentPageNumber] = useState(1);
  const [loading, setLoading] = useState(false);
  const [classificationEndpoint, setClassificationEndpoint] = useState('');
  const [classificationModelId, setClassificationModelId] = useState('');
  const [classificationLoading, setClassificationLoading] = useState(false);
  const [classificationError, setClassificationError] = useState('');
  const [classificationResults, setClassificationResults] = useState<IClassificationResult[]>([]);
  const [selectedDetectedDocumentType, setSelectedDetectedDocumentType] = useState('');
  const [newContractNumber, setNewContractNumber] = useState('');
  const [newDocumentType, setNewDocumentType] = useState('');
  const [selectedEntityKey, setSelectedEntityKey] = useState('');
  const [selectedEntitySiteUrl, setSelectedEntitySiteUrl] = useState('');
  const [uploading, setUploading] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');
  const [pdfDocument, setPdfDocument] = useState<any>(null);
  const [selectedPdfBytes, setSelectedPdfBytes] = useState<ArrayBuffer | null>(null);
  const previewCanvasRef = useRef<HTMLCanvasElement>(null);

  useEffect(() => {
    if (isOpen && selectedPdfFile) {
      loadPdf(selectedPdfFile);
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

  const resetState = () => {
    setPages([]);
    setCurrentPageNumber(1);
    setLoading(false);
    setClassificationLoading(false);
    setClassificationError('');
    setClassificationResults([]);
    setSelectedDetectedDocumentType('');
    setClassificationEndpoint('');
    setClassificationModelId('');
    setNewContractNumber('');
    setNewDocumentType('');
    setSelectedEntityKey('');
    setSelectedEntitySiteUrl('');
    setUploading(false);
    setErrorMessage('');
    setPdfDocument(null);
    setSelectedPdfBytes(null);
  };

  const loadPdf = async (selectedFile: IPdfSelection) => {
    setLoading(true);
    setErrorMessage('');
    setPages([]);
    setPdfDocument(null);

    try {
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

      const pdf = await pdfjsLib.getDocument({ data: pdfBytesForPdfJs }).promise;
      setPdfDocument(pdf);
      setSelectedPdfBytes(pdfBytes);

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
    } catch (error) {
      console.error('Error loading PDF for classification:', error);
      setErrorMessage(error instanceof Error ? error.message : 'Error loading PDF.');
    } finally {
      setLoading(false);
    }
  };

  const renderPdfPage = async (pageNumber: number) => {
    if (!previewCanvasRef.current || !pdfDocument) {
      return;
    }

    const pageInfo = pages[pageNumber - 1];
    if (!pageInfo) {
      return;
    }

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

  const handlePageNavigation = (pageNumber: number) => {
    if (pageNumber < 1 || pageNumber > pages.length) {
      return;
    }
    setCurrentPageNumber(pageNumber);
  };

  const handlePageSelect = (pageId: string, selected: boolean) => {
    setPages(prev => prev.map(page => page.id === pageId ? { ...page, selected } : page));
  };

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
      setPages(prev => prev.map(page => ({ ...page, selected: false })));
      setCurrentPageNumber(1);
      return;
    }

    setSelectedDetectedDocumentType(selectedKey);
    setPages(prev => prev.map(page => ({
      ...page,
      selected: result!.pageNumbers.indexOf(page.sourcePageNumber) !== -1
    })));
    
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

  const classifySelectedDocument = async () => {
    if (!selectedPdfFile) {
      setClassificationError('No file selected for classification.');
      return;
    }

    if (!selectedPdfBytes) {
      setClassificationError('The PDF bytes are not loaded yet. Please reopen the file.');
      return;
    }

    if (!classificationEndpoint.trim() || !classificationModelId.trim()) {
      setClassificationError('Please enter the Azure Function URL and model ID.');
      return;
    }

    setClassificationLoading(true);
    setClassificationError('');
    setClassificationResults([]);
    setSelectedDetectedDocumentType('');

    try {
      const functionUrl = getClassificationFunctionUrl(classificationEndpoint.trim());
      const modelId = classificationModelId.trim();
      const response = await fetch(functionUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json'
        },
        body: JSON.stringify({
          modelId,
          fileBase64: arrayBufferToBase64(selectedPdfBytes)
        })
      });

      if (!response.ok) {
        const body = await response.text();
        throw new Error(`Classification request failed: ${response.status} ${response.statusText} - ${body}`);
      }

      const resultJson = await response.json();

      const results = parseClassificationResults(resultJson);
      if (results.length === 0) {
        throw new Error('No document types were detected by the classification model.');
      }

      const firstKey = results[0].key;
      setClassificationResults(results);
      setSelectedDetectedDocumentType(firstKey);
      setPages(prev => prev.map(page => ({
        ...page,
        selected: firstKey !== '' && results[0].pageNumbers.indexOf(page.sourcePageNumber) !== -1
      })));
    } catch (error) {
      console.error('Classification error:', error);
      const rawMessage = error instanceof Error ? error.message : String(error);
      const corsMessage = getAzureCorsErrorMessage(rawMessage);
      setClassificationError(corsMessage || rawMessage || 'Unknown classification error');
    } finally {
      setClassificationLoading(false);
    }
  };

  const getAzureCorsErrorMessage = (message: string) => {
    const lower = message.toLowerCase();
    if (lower.indexOf('cors') !== -1 || lower.indexOf('access-control') !== -1 || lower.indexOf('failed to fetch') !== -1 || lower.indexOf('networkrequest failed') !== -1) {
      return 'Browser request blocked by CORS. Document Intelligence must be called through a server-side proxy or backend API, not directly from SharePoint client-side code with a subscription key.';
    }
    return null;
  };

  const getClassificationFunctionUrl = (endpoint: string) => {
    const trimmed = endpoint.replace(/\/api\/classify\/?$/i, '').replace(/\/+$/g, '');
    return `${trimmed}/api/classify`;
  };

  const arrayBufferToBase64 = (buffer: ArrayBuffer): string => {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      for (let j = 0; j < chunk.length; j++) {
        binary += String.fromCharCode(chunk[j]);
      }
    }
    return window.btoa(binary);
  };

  const parseClassificationResults = (classificationJson: any): IClassificationResult[] => {
    const rawDocuments = classificationJson.documents || classificationJson.analyzeResult?.documents || classificationJson.documentResults || [];
    const results: IClassificationResult[] = [];

    if (!Array.isArray(rawDocuments)) {
      return results;
    }

    for (let i = 0; i < rawDocuments.length; i++) {
      const doc = rawDocuments[i];
      const key = doc.docType || doc.documentType || doc.type || doc['documentType'] || doc.name || 'Unknown';
      const text = String(key);
      const confidence = typeof doc.confidence === 'number' ? doc.confidence : (typeof doc.confidence === 'string' ? parseFloat(doc.confidence) : 0);
      let pageNumbers: number[] = [];

      if (Array.isArray(doc.boundingRegions)) {
        for (let j = 0; j < doc.boundingRegions.length; j++) {
          const region = doc.boundingRegions[j];
          const pageValue = Number(region.pageNumber);
          if (!isNaN(pageValue)) {
            pageNumbers.push(pageValue);
          }
        }
      }

      if (pageNumbers.length === 0 && typeof doc.pageNumber === 'number') {
        pageNumbers = [doc.pageNumber];
      }

      const uniquePageNumbers: number[] = [];
      for (let j = 0; j < pageNumbers.length; j++) {
        const pageNumber = pageNumbers[j];
        if (uniquePageNumbers.indexOf(pageNumber) === -1) {
          uniquePageNumbers.push(pageNumber);
        }
      }
      uniquePageNumbers.sort((a, b) => a - b);
      pageNumbers = uniquePageNumbers;

      results.push({ key: text, text, confidence, pageNumbers });
    }

    return results;
  };

  const safeODataString = (value: string) => value.replace(/'/g, "''");

  const getServerRelativeUrl = (url: string) => {
    if (!url) {
      return url;
    }
    try {
      const parsed = new URL(url, window.location.origin);
      return parsed.pathname.replace(/\/$/, '');
    } catch {
      return url.replace(/\/$/, '');
    }
  };

  const getAbsoluteSiteUrl = (url: string) => {
    if (!url) {
      return '';
    }
    try {
      const parsed = new URL(url, window.location.origin);
      return `${parsed.protocol}//${parsed.host}${parsed.pathname}`.replace(/\/$/, '');
    } catch {
      return url.replace(/\/$/, '');
    }
  };

  const getRepositoryFolderUrl = async (siteUrl: string, repositoryName: string, documentSetName: string): Promise<{ siteApiBase: string; folderUrl: string } | null> => {
    const serverRelativeSiteUrl = getServerRelativeUrl(siteUrl).replace(/\/$/, '');
    const siteApiBase = getAbsoluteSiteUrl(siteUrl);
    const safeRepositoryName = repositoryName.trim().replace(/^\/+|\/+$/g, '');
    const safeDocumentSetName = documentSetName.trim();

    if (!siteApiBase || !serverRelativeSiteUrl || !safeRepositoryName || !safeDocumentSetName) {
      console.error('Invalid repository target values', { siteUrl, siteApiBase, serverRelativeSiteUrl, repositoryName, documentSetName });
      return null;
    }

    const folderUrl = `${serverRelativeSiteUrl}/${safeRepositoryName}/${safeDocumentSetName}`;
    const safeFolderUrl = folderUrl.replace(/'/g, "''");

    const response = await context.spHttpClient.get(
      `${siteApiBase}/_api/web/GetFolderByServerRelativeUrl('${safeFolderUrl}')`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      console.warn('Target folder not found', { siteApiBase, folderUrl, status: response.status, statusText: response.statusText });
      return null;
    }

    return { siteApiBase, folderUrl };
  };

  const checkInFile = async (serverRelativeUrl: string, siteApiBase: string): Promise<void> => {
    const fileUrlEncoded = encodeURIComponent(serverRelativeUrl);
    const checkInUrl = `${siteApiBase}/_api/web/GetFileByServerRelativeUrl('${fileUrlEncoded}')/CheckIn(comment='Uploaded by AutoClassify',checkintype=0)`;

    const response = await context.spHttpClient.post(checkInUrl, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata'
      }
    });

    if (!response.ok) {
      throw new Error(`File check-in failed: ${response.status} ${response.statusText}`);
    }
  };

  const handleCollateAndUpload = async () => {
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

    const repositoryInfo = await getRepositoryFolderUrl(selectedEntitySiteUrl, destinationDocumentRepositoryTitle, newContractNumber);
    if (!repositoryInfo) {
      alert(`Contract Number '${newContractNumber}' was not found in the selected entity repository '${destinationDocumentRepositoryTitle}'.`);
      return;
    }

    setUploading(true);

    try {
      const sourcePdfMap: { [fileRef: string]: PDFDocument } = {};
      const uniqueSourceFiles = selectedPages.reduce<{ [fileRef: string]: string }>((acc, pageInfo) => {
        acc[pageInfo.sourceFileRef] = pageInfo.sourceFileName;
        return acc;
      }, {});

      const newPdf = await PDFDocument.create();

      for (let i = 0; i < selectedPages.length; i++) {
        const pageInfo = selectedPages[i];
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

      const pdfBytes = await newPdf.save();
      const timestamp = new Date().toISOString().replace(/[T:.]/g, '-').substring(0, 19);
      const fileName = `${newContractNumber}_${newDocumentType}_${timestamp}.pdf`;
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

      await checkInFile(documentRepositoryServerRelativeUrl, repositoryInfo.siteApiBase);

      const currentUserId = context.pageContext.legacyPageContext?.userId;
      if (!currentUserId) {
        throw new Error('Unable to determine the current user ID for AssignedTo update.');
      }

      const sourceFileRefs = Object.keys(uniqueSourceFiles);
      const remainingFileRefs = new Set<string>();

      for (let i = 0; i < sourceFileRefs.length; i++) {
        const sourceFileRef = sourceFileRefs[i];
        const filePages = selectedPages.filter(p => p.sourceFileRef === sourceFileRef);
        const remainingPages = filePages.filter(p => !p.selected);

        if (remainingPages.length === 0) {
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
            AssignedToId: currentUserId
          })
        });

        if (!sourceMetadataResponse.ok) {
          throw new Error(`Source file assignment update failed for ${sourceFileRef}: ${sourceMetadataResponse.status} ${sourceMetadataResponse.statusText}`);
        }
      }

      alert('PDF collated and uploaded successfully.');
      await onUploadSuccess();
      handleDismiss();
    } catch (error) {
      console.error('Error collating, uploading, or updating source PDF:', error);
      alert('Error occurred. Please try again.');
      setUploading(false);
    }
  };

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
      onDismiss={handleDismiss}
      isBlocking={false}
      containerClassName={styles.modalContainer}
    >
      <div className={styles.modalHeader}>
        <h3>Auto Classify Document: {selectedPdfFile?.fileName || 'selected PDF'}</h3>
        <IconButton iconProps={{ iconName: 'Cancel' }} onClick={handleDismiss} title="Close" />
      </div>
      <div className={styles.modalBody}>
        <div className={styles.modalContent}>
          <div className={styles.previewPanel}>
            <div className={styles.previewControls}>
              <PrimaryButton text="Previous" onClick={() => handlePageNavigation(currentPageNumber - 1)} disabled={currentPageNumber <= 1 || loading} />
              <Label className={styles.currentPageLabel}>
                Page {currentPageNumber} of {pages.length}
                {pages.length > 0 && (() => {
                  const currentPage = pages[currentPageNumber - 1];
                  return currentPage ? ` (${currentPage.sourceFileName} page ${currentPage.sourcePageNumber})` : '';
                })()}
              </Label>
              <PrimaryButton text="Next" onClick={() => handlePageNavigation(currentPageNumber + 1)} disabled={currentPageNumber >= pages.length || loading} />
            </div>
            <div className={styles.previewCanvasWrapper}>
              <canvas ref={previewCanvasRef} className={styles.previewCanvas} />
              {loading && <Spinner size={SpinnerSize.large} label="Loading PDF..." />}
            </div>
          </div>

          <div className={styles.selectionPanel}>
            <Label>Document Classification</Label>
            <TextField
              label="Azure Function URL"
              value={classificationEndpoint}
              onChange={(ev, value) => setClassificationEndpoint(value || '')}
              placeholder="https://<your-function-app>.azurewebsites.net"
              description="Enter the Function App base URL (with or without /api/classify). The code will normalize it automatically."
            />
            <TextField
              label="Model ID"
              value={classificationModelId}
              onChange={(ev, value) => setClassificationModelId(value || '')}
            />
            <PrimaryButton text="Classify Document" onClick={classifySelectedDocument} disabled={classificationLoading || loading || !selectedPdfFile} />
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
                      />
                    </div>
                  ))}
                </div>

                <div className={styles.formSection}>
                  <TextField label="Contract Number" value={newContractNumber} onChange={(ev, value) => setNewContractNumber(value || '')} required />
                  <Dropdown
                    label="Document Type"
                    options={documentTypes}
                    selectedKey={newDocumentType || undefined}
                    onChange={(ev, option) => setNewDocumentType(option?.key as string || '')}
                    required
                    disabled={documentTypes.length === 0}
                    placeholder={documentTypes.length === 0 ? 'No document types available' : 'Select a document type'}
                  />
                  <Dropdown
                    label="Entity"
                    options={entityOptions}
                    selectedKey={selectedEntityKey || undefined}
                    onChange={(ev, option) => {
                      setSelectedEntityKey(option?.key as string || '');
                      setSelectedEntitySiteUrl(option?.data as string || '');
                    }}
                    required
                    disabled={entityOptions.length === 0}
                    placeholder={entityOptions.length === 0 ? 'No entities available' : 'Select an entity'}
                  />
                  <PrimaryButton
                    text="Merge and Upload"
                    onClick={handleCollateAndUpload}
                    disabled={uploading || !newContractNumber || !newDocumentType || !selectedEntityKey || selectedDocPages.filter(page => page.selected).length === 0}
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
