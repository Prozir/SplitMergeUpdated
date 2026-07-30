import { SPHttpClient } from '@microsoft/sp-http';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { IDropdownOption } from '@fluentui/react';
import { IClassificationResult } from './splitMergeTypes';

// Escape single quotes for SharePoint OData calls.
export const safeODataString = (value: string): string => value.replace(/'/g, "''");

// Convert a file URL into a server-relative path.
export const getServerRelativeUrl = (url: string): string => {
  if (!url) {
    return url;
  }

  try {
    // Parse the URL and strip the trailing slash.
    const parsed = new URL(url, window.location.origin);
    return parsed.pathname.replace(/\/$/, '');
  } catch {
    return url.replace(/\/$/, '');
  }
};

// Convert a site URL into an absolute site URL.
export const getAbsoluteSiteUrl = (url: string): string => {
  if (!url) {
    return '';
  }

  try {
    // Build the full site URL from the incoming path.
    const parsed = new URL(url, window.location.origin);
    return `${parsed.protocol}//${parsed.host}${parsed.pathname}`.replace(/\/$/, '');
  } catch {
    return url.replace(/\/$/, '');
  }
};

// Clean a value so it can be used safely in a file name.
export const sanitizeFileNamePart = (value: string): string => {
  // Remove invalid characters and normalize spaces.
  const cleaned = value
    .trim()
    .replace(/[\\/:*?"<>|]+/g, '-')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^-+|-+$/g, '');

  return cleaned || 'Unknown';
};

// Turn raw classification data into a simple result list.
export const parseClassificationResults = (classificationJson: any): IClassificationResult[] => {
  const rawDocuments = classificationJson.documents || classificationJson.analyzeResult?.documents || classificationJson.documentResults || [];
  const results: IClassificationResult[] = [];

  if (!Array.isArray(rawDocuments)) {
    return results;
  }

  for (let i = 0; i < rawDocuments.length; i++) {
    const doc = rawDocuments[i];
    // Pick the main document type from the raw response.
    const key = doc.docType || doc.documentType || doc.type || doc['documentType'] || doc.name || 'Unknown';
    const text = String(key);
    const confidence = typeof doc.confidence === 'number' ? doc.confidence : (typeof doc.confidence === 'string' ? parseFloat(doc.confidence) : 0);
    let pageNumbers: number[] = [];

    if (Array.isArray(doc.boundingRegions)) {
      // Collect all page numbers from the bounding regions.
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
    // Keep the page numbers in order.
    uniquePageNumbers.sort((a, b) => a - b);
    pageNumbers = uniquePageNumbers;

    results.push({ key: text, text, confidence, pageNumbers });
  }

  return results;
};

// Find the target folder for a contract in the repository site.
export const getRepositoryFolderUrl = async (
  context: WebPartContext,
  siteUrl: string,
  repositoryName: string,
  documentSetName: string
): Promise<{ siteApiBase: string; folderUrl: string } | null> => {
  const serverRelativeSiteUrl = getServerRelativeUrl(siteUrl).replace(/\/$/, '');
  const siteApiBase = getAbsoluteSiteUrl(siteUrl);
  const safeRepositoryName = repositoryName.trim().replace(/^\/+|\/+$/g, '');
  const safeDocumentSetName = documentSetName.trim();
  // Build the repository folder path from the site and contract name.

  if (!siteApiBase || !serverRelativeSiteUrl || !safeRepositoryName || !safeDocumentSetName) {
    console.error('Invalid repository target values', { siteUrl, siteApiBase, serverRelativeSiteUrl, repositoryName, documentSetName });
    return null;
  }

  const folderUrl = `${serverRelativeSiteUrl}/${safeRepositoryName}/${safeDocumentSetName}`;
  // Escape the folder path for the SharePoint API call.
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

// Check in a file with a short comment.
export const checkInFile = async (context: WebPartContext, serverRelativeUrl: string, siteApiBase: string, comment: string): Promise<void> => {
  // Encode the file path before calling the check-in endpoint.
  const fileUrlEncoded = encodeURIComponent(serverRelativeUrl);
  const safeComment = comment.replace(/'/g, "''");
  const checkInUrl = `${siteApiBase}/_api/web/GetFileByServerRelativeUrl('${fileUrlEncoded}')/CheckIn(comment='${safeComment}',checkintype=0)`;

  const response = await context.spHttpClient.post(checkInUrl, SPHttpClient.configurations.v1, {
    headers: {
      'Accept': 'application/json;odata=nometadata'
    }
  });

  if (!response.ok) {
    throw new Error(`File check-in failed: ${response.status} ${response.statusText}`);
  }
};

// Build a unique list of entity options for the dropdown.
export const buildDistinctEntityOptions = (items: any[]): IDropdownOption[] => {
  const seenEntities = new Set<string>();

  return items
    // Keep only items that have both entity and site information.
    .filter((item: any) => item.Entity && item.CMSSiteURL && item.CMSSiteURL.toString().trim() !== '')
    .map((item: any) => ({
      key: item.Entity,
      text: item.Entity,
      data: item.CMSSiteURL.toString().trim()
    }))
    .filter(option => {
      const key = option.key as string;
      if (seenEntities.has(key)) {
        return false;
      }
      seenEntities.add(key);
      return true;
    });
};
