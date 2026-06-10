import * as React from 'react';
import styles from './PdfESignature.module.scss';
import type { IPdfESignatureProps } from './IPdfESignatureProps';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { getDocument, GlobalWorkerOptions, PDFDocumentProxy } from 'pdfjs-dist/legacy/build/pdf';

GlobalWorkerOptions.workerSrc = require('pdfjs-dist/legacy/build/pdf.worker.min.js');

interface IPdfFile {
  name: string;
  serverRelativeUrl: string;
  uniqueId: string;
}

interface ILibraryTarget {
  folderServerRelativeUrl: string;
  webAbsoluteUrl: string;
}

interface IDirectPdfTarget {
  webAbsoluteUrl: string;
  fileServerRelativeUrl: string;
  folderAbsoluteUrl: string;
  fileName: string;
}

interface IPoint {
  x: number;
  y: number;
}

interface IPreviewInfo {
  scale: number;
  pageWidth: number;
  pageHeight: number;
}

interface ISignaturePlacement {
  x: number;
  y: number;
  width: number;
  height: number;
}

type ResizeHandle = 'l' | 'r';

const SIGNATURE_CANVAS_WIDTH = 250;
const SIGNATURE_CANVAS_HEIGHT = 120;
const SIGNATURE_LANDSCAPE_RATIO = SIGNATURE_CANVAS_WIDTH / SIGNATURE_CANVAS_HEIGHT;
const MIN_PLACEMENT_SIZE = SIGNATURE_CANVAS_WIDTH;

const createPlacementFromPoints = (start: IPoint, end: IPoint, canvas: HTMLCanvasElement): ISignaturePlacement => {
  const clamp = (value: number, min: number, max: number): number => Math.max(min, Math.min(max, value));

  const startX: number = clamp(start.x, 0, canvas.width);
  const startY: number = clamp(start.y, 0, canvas.height - MIN_PLACEMENT_SIZE / SIGNATURE_LANDSCAPE_RATIO);
  const endX: number = clamp(end.x, 0, canvas.width);

  let width: number = Math.max(MIN_PLACEMENT_SIZE, Math.abs(endX - startX));
  const draggingRight: boolean = endX >= startX;

  let x: number = draggingRight ? startX : startX - width;
  x = clamp(x, 0, canvas.width - MIN_PLACEMENT_SIZE);

  const maxWidthFromX: number = draggingRight ? canvas.width - x : startX;
  const maxWidthFromY: number = (canvas.height - startY) * SIGNATURE_LANDSCAPE_RATIO;
  width = clamp(width, MIN_PLACEMENT_SIZE, Math.max(MIN_PLACEMENT_SIZE, Math.min(maxWidthFromX, maxWidthFromY)));

  if (!draggingRight) {
    x = startX - width;
  }

  const height: number = width / SIGNATURE_LANDSCAPE_RATIO;
  const y: number = clamp(startY, 0, canvas.height - height);

  return { x, y, width, height };
};

const SIGNATURE_STROKE_COLOR = '#2563eb';
const MIN_PREVIEW_ZOOM = 0.7;
const MAX_PREVIEW_ZOOM = 3;
const SIGNATURE_OFFSET_X = 0;
const SIGNATURE_OFFSET_Y = 5;

const encodeODataValue = (value: string): string => encodeURIComponent(value);

const trimTrailingSlash = (value: string): string => value.replace(/\/$/, '');

const formatSignedAt = (date: Date): string => {
  const pad2 = (value: number): string => (`0${value}`).slice(-2);

  const year: number = date.getFullYear();
  const month: string = pad2(date.getMonth() + 1);
  const day: string = pad2(date.getDate());
  const hours: string = pad2(date.getHours());
  const minutes: string = pad2(date.getMinutes());

  return `${year}-${month}-${day}, ${hours}:${minutes}`;
};

const parseDirectPdfUrl = (input: string): IDirectPdfTarget | undefined => {
  try {
    const url = new URL(input.trim());
    const decodedPath: string = decodeURIComponent(url.pathname).replace(/\/$/, '');
    const segments: string[] = decodedPath.split('/').filter(Boolean);

    if (segments.length < 1) {
      return undefined;
    }

    const fileName: string = segments[segments.length - 1];
    if (!fileName.toLowerCase().endsWith('.pdf')) {
      return undefined;
    }

    const folderServerRelativeUrl: string = `/${segments.slice(0, -1).join('/')}`;
    const fileServerRelativeUrl: string = `/${segments.join('/')}`;

    let webServerRelativeUrl: string = '';
    if ((segments[0] === 'sites' || segments[0] === 'teams') && segments.length >= 2) {
      webServerRelativeUrl = `/${segments[0]}/${segments[1]}`;
    }

    const webAbsoluteUrl: string = `${url.origin}${webServerRelativeUrl}`;
    const folderAbsoluteUrl: string = `${url.origin}${folderServerRelativeUrl}`;

    return {
      webAbsoluteUrl,
      fileServerRelativeUrl,
      folderAbsoluteUrl,
      fileName
    };
  } catch {
    return undefined;
  }
};

const parseLibraryUrl = (input: string): ILibraryTarget | undefined => {
  try {
    const url = new URL(input.trim());
    const segments: string[] = decodeURIComponent(url.pathname)
      .replace(/\/$/, '')
      .split('/')
      .filter(Boolean);

    // Remove SharePoint list view paths such as /Forms/AllItems.aspx.
    while (segments.length > 0) {
      const lastSegment: string = segments[segments.length - 1].toLowerCase();
      if (lastSegment === 'forms' || lastSegment.endsWith('.aspx')) {
        segments.pop();
      } else {
        break;
      }
    }

    if (segments.length === 0) {
      return undefined;
    }

    let webServerRelativeUrl: string = '';
    if ((segments[0] === 'sites' || segments[0] === 'teams') && segments.length >= 2) {
      webServerRelativeUrl = `/${segments[0]}/${segments[1]}`;
    }

    return {
      folderServerRelativeUrl: `/${segments.join('/')}`,
      webAbsoluteUrl: `${url.origin}${webServerRelativeUrl}`
    };
  } catch {
    return undefined;
  }
};

const dataUrlToBytes = (dataUrl: string): Uint8Array => {
  const base64Data: string = dataUrl.split(',')[1] || '';
  const binary: string = atob(base64Data);
  const bytes: Uint8Array = new Uint8Array(binary.length);

  for (let index = 0; index < binary.length; index++) {
    bytes[index] = binary.charCodeAt(index);
  }

  return bytes;
};

const getRelativePoint = (evt: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>): IPoint => {
  const canvas: HTMLCanvasElement = evt.currentTarget;
  const rect: DOMRect = canvas.getBoundingClientRect();
  const scaleX: number = canvas.width / rect.width;
  const scaleY: number = canvas.height / rect.height;

  if ('touches' in evt && evt.touches.length > 0) {
    return {
      x: (evt.touches[0].clientX - rect.left) * scaleX,
      y: (evt.touches[0].clientY - rect.top) * scaleY
    };
  }

  if ('changedTouches' in evt && evt.changedTouches.length > 0) {
    return {
      x: (evt.changedTouches[0].clientX - rect.left) * scaleX,
      y: (evt.changedTouches[0].clientY - rect.top) * scaleY
    };
  }

  const mouseEvt: React.MouseEvent<HTMLCanvasElement> = evt as React.MouseEvent<HTMLCanvasElement>;
  return {
    x: (mouseEvt.clientX - rect.left) * scaleX,
    y: (mouseEvt.clientY - rect.top) * scaleY
  };
};

const createApiUrl = (webAbsoluteUrl: string, apiPath: string): string => `${trimTrailingSlash(webAbsoluteUrl)}${apiPath}`;

interface IClientPoint {
  clientX: number;
  clientY: number;
}

const getTouchDistance = (touchA: IClientPoint, touchB: IClientPoint): number => {
  const dx: number = touchA.clientX - touchB.clientX;
  const dy: number = touchA.clientY - touchB.clientY;
  return Math.sqrt(dx * dx + dy * dy);
};

const getTouchCenter = (touchA: IClientPoint, touchB: IClientPoint): IPoint => ({
  x: (touchA.clientX + touchB.clientX) / 2,
  y: (touchA.clientY + touchB.clientY) / 2
});

const PdfESignature: React.FC<IPdfESignatureProps> = (props) => {
  const sampleDirectPdfParamUrl: string = `${props.currentWebUrl}/SitePages/PdfESignature.aspx?pdf-url=${encodeURIComponent(`${props.currentWebUrl}/Shared Documents/Example.pdf`)}`;

  const [sourcePathInput, setSourcePathInput] = React.useState<string>('');
  const [sourceStatus, setSourceStatus] = React.useState<string>('');
  const [pdfFiles, setPdfFiles] = React.useState<IPdfFile[]>([]);
  const [selectedPdfUrl, setSelectedPdfUrl] = React.useState<string>('');
  const [selectedPdfName, setSelectedPdfName] = React.useState<string>('');
  const [isLoadingFiles, setIsLoadingFiles] = React.useState<boolean>(false);
  const [isLoadingPdf, setIsLoadingPdf] = React.useState<boolean>(false);
  const [isPreviewReady, setIsPreviewReady] = React.useState<boolean>(false);
  const [previewZoom, setPreviewZoom] = React.useState<number>(1);
  const [isSigning, setIsSigning] = React.useState<boolean>(false);
  const [isSignatureModalOpen, setIsSignatureModalOpen] = React.useState<boolean>(false);
  const [signaturePlacement, setSignaturePlacement] = React.useState<ISignaturePlacement | null>(null);
  const [draftPlacement, setDraftPlacement] = React.useState<ISignaturePlacement | null>(null);
  const [isSelectingPlacement, setIsSelectingPlacement] = React.useState<boolean>(false);
  const [signatureImageDataUrl, setSignatureImageDataUrl] = React.useState<string>('');
  const [destinationPathInput, setDestinationPathInput] = React.useState<string>('');
  const [successNotification, setSuccessNotification] = React.useState<string>('');
  const [saveErrorNotification, setSaveErrorNotification] = React.useState<string>('');
  const [isDirectPdfMode, setIsDirectPdfMode] = React.useState<boolean>(false);
  const [directPdfNotification, setDirectPdfNotification] = React.useState<string>('');
  const [isPostSaveModalOpen, setIsPostSaveModalOpen] = React.useState<boolean>(false);
  const [signerEmail, setSignerEmail] = React.useState<string>('');
  const [signerTimestamp, setSignerTimestamp] = React.useState<string>('');

  const signatureCanvasRef = React.useRef<HTMLCanvasElement | null>(null);
  const previewCanvasRef = React.useRef<HTMLCanvasElement | null>(null);
  const previewWrapRef = React.useRef<HTMLDivElement | null>(null);
  const previewInfoRef = React.useRef<IPreviewInfo | null>(null);
  const sourceBytesRef = React.useRef<ArrayBuffer | null>(null);
  const pdfDocumentRef = React.useRef<PDFDocumentProxy | null>(null);
  const isDrawingRef = React.useRef<boolean>(false);
  const lastDrawPointRef = React.useRef<IPoint | null>(null);
  const dragStartPointRef = React.useRef<IPoint | null>(null);
  const isPinchPanningRef = React.useRef<boolean>(false);
  const pinchStartDistanceRef = React.useRef<number>(0);
  const pinchStartZoomRef = React.useRef<number>(1);
  const pinchStartCenterRef = React.useRef<IPoint | null>(null);
  const pinchStartScrollRef = React.useRef<{ left: number; top: number } | null>(null);
  const [isResizingPlacement, setIsResizingPlacement] = React.useState<boolean>(false);
  const [useSameDestinationPath, setUseSameDestinationPath] = React.useState<boolean>(true);
  const resizeHandleRef = React.useRef<ResizeHandle | null>(null);
  const resizeStartPointRef = React.useRef<IPoint | null>(null);
  const resizeStartPlacementRef = React.useRef<ISignaturePlacement | null>(null);
  const hasTriedDirectPdfRef = React.useRef<boolean>(false);

  const ensureLibraryExists = React.useCallback(async (target: ILibraryTarget): Promise<void> => {
    const folderApiUrl: string = createApiUrl(
      target.webAbsoluteUrl,
      `/_api/web/GetFolderByServerRelativePath(decodedUrl='${encodeODataValue(target.folderServerRelativeUrl)}')?$select=Exists,ServerRelativeUrl,Name`
    );

    const folderResponse: SPHttpClientResponse = await props.spHttpClient.get(folderApiUrl, SPHttpClient.configurations.v1);
    if (!folderResponse.ok) {
      const errorText = await folderResponse.text().catch(() => folderResponse.statusText);
      throw new Error(`Path validation failed (${folderResponse.status}): ${errorText || 'Access denied or path not found'}`);
    }
  }, [props.spHttpClient]);

  const queryPdfFiles = React.useCallback(async (target: ILibraryTarget): Promise<IPdfFile[]> => {
    interface ISpFileEntry {
      Name: string;
      ServerRelativeUrl: string;
      UniqueId: string;
    }

    interface ISpFolderEntry {
      Name: string;
      ServerRelativeUrl: string;
    }

    const getCollection = <T,>(json: unknown): T[] => {
      const payload = json as {
        value?: T[];
        d?: {
          results?: T[];
        };
      };

      if (Array.isArray(payload.value)) {
        return payload.value;
      }

      if (Array.isArray(payload.d?.results)) {
        return payload.d.results;
      }

      return [];
    };

    const getFolderFilesAndFolders = async (folderServerRelativeUrl: string): Promise<{ files: ISpFileEntry[]; folders: ISpFolderEntry[] }> => {
      const filesApiUrl: string = createApiUrl(
        target.webAbsoluteUrl,
        `/_api/web/GetFolderByServerRelativePath(decodedUrl='${encodeODataValue(folderServerRelativeUrl)}')/Files?$select=Name,ServerRelativeUrl,UniqueId`
      );
      const foldersApiUrl: string = createApiUrl(
        target.webAbsoluteUrl,
        `/_api/web/GetFolderByServerRelativePath(decodedUrl='${encodeODataValue(folderServerRelativeUrl)}')/Folders?$select=Name,ServerRelativeUrl`
      );

      const [filesResponse, foldersResponse] = await Promise.all([
        props.spHttpClient.get(filesApiUrl, SPHttpClient.configurations.v1, {
          headers: { Accept: 'application/json;odata=nometadata' }
        }),
        props.spHttpClient.get(foldersApiUrl, SPHttpClient.configurations.v1, {
          headers: { Accept: 'application/json;odata=nometadata' }
        })
      ]);

      if (!filesResponse.ok) {
        const errorText = await filesResponse.text().catch(() => filesResponse.statusText);
        throw new Error(`Failed to read files (${filesResponse.status}): ${errorText || 'Check permissions'}`);
      }

      if (!foldersResponse.ok) {
        const errorText = await foldersResponse.text().catch(() => foldersResponse.statusText);
        throw new Error(`Failed to read folders (${foldersResponse.status}): ${errorText || 'Check permissions'}`);
      }

      const filesJson: unknown = await filesResponse.json();
      const foldersJson: unknown = await foldersResponse.json();

      return {
        files: getCollection<ISpFileEntry>(filesJson),
        folders: getCollection<ISpFolderEntry>(foldersJson)
      };
    };

    const visitedFolders: Set<string> = new Set<string>();
    const queue: string[] = [target.folderServerRelativeUrl];
    const collectedFiles: ISpFileEntry[] = [];

    while (queue.length > 0) {
      const currentFolderUrl: string | undefined = queue.shift();
      if (!currentFolderUrl) {
        continue;
      }

      const normalizedFolderUrl: string = currentFolderUrl.toLowerCase();
      if (visitedFolders.has(normalizedFolderUrl)) {
        continue;
      }
      visitedFolders.add(normalizedFolderUrl);

      const { files, folders } = await getFolderFilesAndFolders(currentFolderUrl);
      collectedFiles.push(...files);

      folders
        .filter((folder) => folder.Name.toLowerCase() !== 'forms')
        .forEach((folder) => {
          if (!visitedFolders.has(folder.ServerRelativeUrl.toLowerCase())) {
            queue.push(folder.ServerRelativeUrl);
          }
        });
    }

    const pdfFiles: IPdfFile[] = collectedFiles
      .filter((file) => file.Name.toLowerCase().endsWith('.pdf'))
      .map((file) => ({
        name: file.Name,
        serverRelativeUrl: file.ServerRelativeUrl,
        uniqueId: file.UniqueId
      }));

    console.log(`Scanned ${visitedFolders.size} folder(s), found ${collectedFiles.length} total files and ${pdfFiles.length} PDF(s).`);
    return pdfFiles;
  }, [props.spHttpClient]);

  const loadPdfPreview = React.useCallback(async (fileServerRelativeUrl: string, overrideWebAbsoluteUrl?: string, displayName?: string): Promise<void> => {
    const sourceTarget: ILibraryTarget | undefined = overrideWebAbsoluteUrl ? undefined : parseLibraryUrl(sourcePathInput);
    const webAbsoluteUrl: string | undefined = overrideWebAbsoluteUrl || sourceTarget?.webAbsoluteUrl;
    if (!webAbsoluteUrl) {
      setSourceStatus('Invalid source URL format.');
      return;
    }

    const downloadUrl: string = createApiUrl(
      webAbsoluteUrl,
      `/_api/web/GetFileByServerRelativePath(decodedUrl='${encodeODataValue(fileServerRelativeUrl)}')/$value`
    );

    setIsLoadingPdf(true);
    setIsPreviewReady(false);
    setSourceStatus('Loading selected PDF...');

    try {
      const pdfResponse: SPHttpClientResponse = await props.spHttpClient.get(downloadUrl, SPHttpClient.configurations.v1);
      if (!pdfResponse.ok) {
        throw new Error('Unable to load selected PDF.');
      }

      const bytes: ArrayBuffer = await pdfResponse.arrayBuffer();
      sourceBytesRef.current = bytes;

      const loadingTask = getDocument({ data: new Uint8Array(bytes) });
      const pdfDoc: PDFDocumentProxy = await loadingTask.promise;
      pdfDocumentRef.current = pdfDoc;

      const firstPage = await pdfDoc.getPage(1);
      const viewportAtScaleOne = firstPage.getViewport({ scale: 1 });
      const targetWidth: number = 1600;
      const renderScale: number = Math.min(3, Math.max(1, targetWidth / viewportAtScaleOne.width));
      const viewport = firstPage.getViewport({ scale: renderScale });

      const canvas = previewCanvasRef.current;
      if (!canvas) {
        throw new Error('Preview canvas is unavailable.');
      }

      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);

      const context = canvas.getContext('2d');
      if (!context) {
        throw new Error('Cannot get preview canvas context.');
      }

      const isSafariBrowser: boolean = /^((?!chrome|android).)*safari/i.test(window.navigator.userAgent);
      if (isSafariBrowser) {
        canvas.style.transform = 'none';
        canvas.style.webkitTransform = 'none';
        await new Promise<void>((resolve) => {
          window.requestAnimationFrame(() => resolve());
        });
      }

      context.setTransform(1, 0, 0, 1, 0, 0);
      context.clearRect(0, 0, canvas.width, canvas.height);

      await firstPage.render({ canvasContext: context, viewport }).promise;

      previewInfoRef.current = {
        scale: renderScale,
        pageWidth: viewportAtScaleOne.width,
        pageHeight: viewportAtScaleOne.height
      };

      setSignaturePlacement(null);
      setDraftPlacement(null);
      setIsSelectingPlacement(false);
      setPreviewZoom(1);
      setIsPreviewReady(true);
      setSourceStatus(`Loaded: ${displayName || selectedPdfName}. Click the preview to place signature coordinates.`);
    } catch (error) {
      setSourceStatus(error instanceof Error ? error.message : 'Unexpected error while loading PDF.');
      throw error;
    } finally {
      setIsLoadingPdf(false);
    }
  }, [props.spHttpClient, selectedPdfName, sourcePathInput]);

  React.useEffect(() => {
    if (hasTriedDirectPdfRef.current) {
      return;
    }
    hasTriedDirectPdfRef.current = true;

    const params: URLSearchParams = new URLSearchParams(window.location.search);
    const directPdfParam: string = params.get('pdf-url') || '';

    if (!directPdfParam) {
      return;
    }

    const parsedDirect: IDirectPdfTarget | undefined = parseDirectPdfUrl(directPdfParam);
    if (!parsedDirect) {
      setDirectPdfNotification('Direct PDF URL parameter is invalid. Please use Source document library full URL.');
      setIsDirectPdfMode(false);
      return;
    }

    setSourcePathInput(parsedDirect.folderAbsoluteUrl);
    setSelectedPdfUrl(parsedDirect.fileServerRelativeUrl);
    setSelectedPdfName(parsedDirect.fileName);
    setSourceStatus('Loading PDF from URL parameter...');

    loadPdfPreview(parsedDirect.fileServerRelativeUrl, parsedDirect.webAbsoluteUrl, parsedDirect.fileName)
      .then(() => {
        setIsDirectPdfMode(true);
        setDirectPdfNotification('');
      })
      .catch(() => {
        setIsDirectPdfMode(false);
        setDirectPdfNotification('File not exists from URL parameter. You can use Source document library full URL.');
        setSelectedPdfUrl('');
        setSelectedPdfName('');
      });
  }, [loadPdfPreview]);

  React.useEffect(() => {
    if (!isSignatureModalOpen) {
      return;
    }

    const canvas = signatureCanvasRef.current;
    if (!canvas) {
      return;
    }

    canvas.width = SIGNATURE_CANVAS_WIDTH;
    canvas.height = SIGNATURE_CANVAS_HEIGHT;
    const context = canvas.getContext('2d');
    if (!context) {
      return;
    }

    context.clearRect(0, 0, canvas.width, canvas.height);
    context.strokeStyle = SIGNATURE_STROKE_COLOR;
    context.lineWidth = 2;
    context.lineCap = 'round';
    context.lineJoin = 'round';
  }, [isSignatureModalOpen]);

  React.useEffect(() => {
    if (!saveErrorNotification) {
      return;
    }

    const timeout = window.setTimeout(() => {
      setSaveErrorNotification('');
    }, 4500);

    return () => {
      window.clearTimeout(timeout);
    };
  }, [saveErrorNotification]);

  const handleLoadLibrary = React.useCallback(async (): Promise<void> => {
    setPdfFiles([]);
    setSelectedPdfUrl('');
    setSelectedPdfName('');
    setSignaturePlacement(null);
    setDraftPlacement(null);
    setIsSelectingPlacement(false);
    setIsPreviewReady(false);
    setPreviewZoom(1);
    setIsSignatureModalOpen(false);
    setUseSameDestinationPath(true);
    setDestinationPathInput('');
    sourceBytesRef.current = null;

    const trimmedInput: string = sourcePathInput.trim();
    const target: ILibraryTarget | undefined = parseLibraryUrl(trimmedInput);
    if (!trimmedInput || !target) {
      setSourceStatus('Please enter a valid full URL path for document library.');
      return;
    }

    setIsLoadingFiles(true);
    setSourceStatus(`Validating path: ${target.folderServerRelativeUrl}...`);

    try {
      await ensureLibraryExists(target);
      const files: IPdfFile[] = await queryPdfFiles(target);

      if (files.length === 0) {
        setSourceStatus(`Path exists but has no PDF documents. (Total files in this path: check library contents)`);
        setPdfFiles([]);
        return;
      }

      setPdfFiles(files);
      setSelectedPdfUrl(files[0].serverRelativeUrl);
      setSelectedPdfName(files[0].name);
      setSourceStatus(`Path exists. Found ${files.length} PDF file(s). Selected: ${files[0].name}`);
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : 'Path not exists.';
      console.error('Library loading error:', errorMsg);
      setSourceStatus(errorMsg);
    } finally {
      setIsLoadingFiles(false);
    }
  }, [ensureLibraryExists, queryPdfFiles, sourcePathInput]);

  React.useEffect(() => {
    if (!selectedPdfUrl) {
      return;
    }

    loadPdfPreview(selectedPdfUrl).catch((error) => {
      setSourceStatus(error instanceof Error ? error.message : 'Failed to load selected PDF.');
    });
  }, [loadPdfPreview, selectedPdfUrl]);

  const clearSignaturePad = React.useCallback((): void => {
    const canvas = signatureCanvasRef.current;
    if (!canvas) {
      return;
    }

    const context = canvas.getContext('2d');
    if (!context) {
      return;
    }

    context.clearRect(0, 0, canvas.width, canvas.height);
    context.strokeStyle = SIGNATURE_STROKE_COLOR;
    context.lineWidth = 2;
    setSignatureImageDataUrl('');
    setSignerTimestamp('');
    setSignerEmail('');
  }, []);

  const resolveCurrentUserEmail = React.useCallback(async (): Promise<string> => {
    const currentUserApiUrl: string = createApiUrl(
      props.currentWebUrl,
      '/_api/web/currentuser?$select=Email,LoginName'
    );

    try {
      const response: SPHttpClientResponse = await props.spHttpClient.get(
        currentUserApiUrl,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        return 'Email unavailable';
      }

      const payload: unknown = await response.json();
      const user = payload as { Email?: string; LoginName?: string };
      const emailFromProfile: string = (user.Email || '').trim();
      if (emailFromProfile) {
        return emailFromProfile;
      }

      const loginName: string = (user.LoginName || '').trim();
      const loginMatch: RegExpMatchArray | null = loginName.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
      return loginMatch?.[0] || 'Email unavailable';
    } catch {
      return 'Email unavailable';
    }
  }, [props.currentWebUrl, props.spHttpClient]);

  const handleSignatureDone = React.useCallback(async (): Promise<void> => {
    const canvas = signatureCanvasRef.current;
    if (!canvas) {
      return;
    }

    const resolvedEmail: string = await resolveCurrentUserEmail();
    const signedAt: string = formatSignedAt(new Date());
    setSignerEmail(resolvedEmail);
    setSignerTimestamp(signedAt);

    const signatureDataUrl: string = canvas.toDataURL('image/png');
    setSignatureImageDataUrl(signatureDataUrl);
    setIsSignatureModalOpen(false);
  }, [resolveCurrentUserEmail]);

  const startDraw = React.useCallback((evt: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>): void => {
    evt.preventDefault();
    const canvas = signatureCanvasRef.current;
    if (!canvas) {
      return;
    }

    isDrawingRef.current = true;
    lastDrawPointRef.current = getRelativePoint(evt);
  }, []);

  const draw = React.useCallback((evt: React.MouseEvent<HTMLCanvasElement> | React.TouchEvent<HTMLCanvasElement>): void => {
    evt.preventDefault();
    if (!isDrawingRef.current) {
      return;
    }

    const canvas = signatureCanvasRef.current;
    if (!canvas || !lastDrawPointRef.current) {
      return;
    }

    const context = canvas.getContext('2d');
    if (!context) {
      return;
    }

    const currentPoint: IPoint = getRelativePoint(evt);
    context.beginPath();
    context.moveTo(lastDrawPointRef.current.x, lastDrawPointRef.current.y);
    context.lineTo(currentPoint.x, currentPoint.y);
    context.stroke();
    lastDrawPointRef.current = currentPoint;
  }, []);

  const stopDraw = React.useCallback((): void => {
    isDrawingRef.current = false;
    lastDrawPointRef.current = null;
  }, []);

  const startResizePlacement = React.useCallback((handle: ResizeHandle, point: IPoint): void => {
    if (!signaturePlacement || !previewCanvasRef.current) {
      return;
    }

    resizeHandleRef.current = handle;
    resizeStartPointRef.current = point;
    resizeStartPlacementRef.current = { ...signaturePlacement };
    setIsResizingPlacement(true);
  }, [signaturePlacement]);

  const getPreviewPointFromClient = React.useCallback((clientX: number, clientY: number): IPoint | undefined => {
    const previewCanvas: HTMLCanvasElement | null = previewCanvasRef.current;
    if (!previewCanvas) {
      return undefined;
    }

    const rect: DOMRect = previewCanvas.getBoundingClientRect();
    const scaleX: number = previewCanvas.width / rect.width;
    const scaleY: number = previewCanvas.height / rect.height;

    return {
      x: (clientX - rect.left) * scaleX,
      y: (clientY - rect.top) * scaleY
    };
  }, []);

  const applyResizePlacement = React.useCallback((currentPoint: IPoint): boolean => {
    if (!(isResizingPlacement && resizeHandleRef.current && resizeStartPointRef.current && resizeStartPlacementRef.current)) {
      return false;
    }

    const previewCanvas: HTMLCanvasElement | null = previewCanvasRef.current;
    if (!previewCanvas) {
      return true;
    }

    const clamp = (value: number, min: number, max: number): number => Math.max(min, Math.min(max, value));
    const dx: number = currentPoint.x - resizeStartPointRef.current.x;
    const initial: ISignaturePlacement = resizeStartPlacementRef.current;

    const initialLeft: number = initial.x;
    const initialTop: number = initial.y;
    const initialRight: number = initial.x + initial.width;

    let left: number = initialLeft;
    let right: number = initialRight;

    switch (resizeHandleRef.current) {
      case 'l':
        left = clamp(initialLeft + dx, 0, initialRight - MIN_PLACEMENT_SIZE);
        break;
      case 'r':
        right = clamp(initialRight + dx, initialLeft + MIN_PLACEMENT_SIZE, previewCanvas.width);
        break;
    }

    let width: number = right - left;
    width = clamp(width, MIN_PLACEMENT_SIZE, previewCanvas.width);
    let height: number = width / SIGNATURE_LANDSCAPE_RATIO;

    const maxHeight: number = previewCanvas.height - initialTop;
    if (height > maxHeight) {
      height = maxHeight;
      width = height * SIGNATURE_LANDSCAPE_RATIO;
      if (resizeHandleRef.current === 'l') {
        left = initialRight - width;
      } else {
        right = initialLeft + width;
      }
    }

    setSignaturePlacement({
      x: left,
      y: initialTop,
      width,
      height
    });

    return true;
  }, [isResizingPlacement]);

  const applyDraftPlacement = React.useCallback((currentPoint: IPoint): void => {
    if (!isSelectingPlacement || !dragStartPointRef.current) {
      return;
    }

    const previewCanvas: HTMLCanvasElement | null = previewCanvasRef.current;
    if (!previewCanvas) {
      return;
    }

    const placement: ISignaturePlacement = createPlacementFromPoints(dragStartPointRef.current, currentPoint, previewCanvas);
    setDraftPlacement(placement);
  }, [isSelectingPlacement]);

  const handleResizeMouseDown = React.useCallback((handle: ResizeHandle, evt: React.MouseEvent<HTMLButtonElement>): void => {
    evt.preventDefault();
    evt.stopPropagation();

    const point: IPoint | undefined = getPreviewPointFromClient(evt.clientX, evt.clientY);
    if (!point) {
      return;
    }

    startResizePlacement(handle, point);
  }, [getPreviewPointFromClient, startResizePlacement]);

  const handleResizeTouchStart = React.useCallback((handle: ResizeHandle, evt: React.TouchEvent<HTMLButtonElement>): void => {
    evt.preventDefault();
    evt.stopPropagation();

    if (evt.touches.length === 0) {
      return;
    }

    const touch = evt.touches[0];
    const point: IPoint | undefined = getPreviewPointFromClient(touch.clientX, touch.clientY);
    if (!point) {
      return;
    }

    startResizePlacement(handle, point);
  }, [getPreviewPointFromClient, startResizePlacement]);

  const finishResizePlacement = React.useCallback((): void => {
    if (!isResizingPlacement) {
      return;
    }

    setIsResizingPlacement(false);
    resizeHandleRef.current = null;
    resizeStartPointRef.current = null;
    resizeStartPlacementRef.current = null;
  }, [isResizingPlacement]);

  const handlePreviewMouseDown = React.useCallback((evt: React.MouseEvent<HTMLCanvasElement>): void => {
    if (signatureImageDataUrl) {
      return;
    }

    if (isResizingPlacement) {
      return;
    }

    if (!isPreviewReady || !previewInfoRef.current) {
      return;
    }

    const previewCanvas: HTMLCanvasElement | null = previewCanvasRef.current;
    if (!previewCanvas) {
      return;
    }

    const startPoint: IPoint = getRelativePoint(evt);
    dragStartPointRef.current = startPoint;
    setIsSelectingPlacement(true);
    setDraftPlacement({ x: startPoint.x, y: startPoint.y, width: 0, height: 0 });
    setSignatureImageDataUrl('');
  }, [isPreviewReady, isResizingPlacement, signatureImageDataUrl]);

  const handlePreviewTouchStart = React.useCallback((evt: React.TouchEvent<HTMLCanvasElement>): void => {
    if (signatureImageDataUrl) {
      return;
    }

    if (evt.touches.length >= 2) {
      const wrap = previewWrapRef.current;
      if (!wrap) {
        return;
      }

      const touchA = evt.touches[0];
      const touchB = evt.touches[1];
      isPinchPanningRef.current = true;
      pinchStartDistanceRef.current = getTouchDistance(touchA, touchB);
      pinchStartZoomRef.current = previewZoom;
      pinchStartCenterRef.current = getTouchCenter(touchA, touchB);
      pinchStartScrollRef.current = { left: wrap.scrollLeft, top: wrap.scrollTop };
      return;
    }

    evt.preventDefault();

    if (isResizingPlacement) {
      return;
    }

    if (!isPreviewReady || !previewInfoRef.current) {
      return;
    }

    const previewCanvas: HTMLCanvasElement | null = previewCanvasRef.current;
    if (!previewCanvas) {
      return;
    }

    const startPoint: IPoint = getRelativePoint(evt);
    dragStartPointRef.current = startPoint;
    setIsSelectingPlacement(true);
    setDraftPlacement({ x: startPoint.x, y: startPoint.y, width: 0, height: 0 });
    setSignatureImageDataUrl('');
  }, [isPreviewReady, isResizingPlacement, previewZoom, signatureImageDataUrl]);

  const handlePreviewMouseMove = React.useCallback((evt: React.MouseEvent<HTMLCanvasElement>): void => {
    const currentPoint: IPoint = getRelativePoint(evt);
    if (applyResizePlacement(currentPoint)) {
      return;
    }

    applyDraftPlacement(currentPoint);
  }, [applyDraftPlacement, applyResizePlacement]);

  const handlePreviewTouchMove = React.useCallback((evt: React.TouchEvent<HTMLCanvasElement>): void => {
    if (isPinchPanningRef.current && evt.touches.length >= 2) {
      evt.preventDefault();
      const wrap = previewWrapRef.current;
      if (!wrap || !pinchStartCenterRef.current || !pinchStartScrollRef.current || pinchStartDistanceRef.current <= 0) {
        return;
      }

      const touchA = evt.touches[0];
      const touchB = evt.touches[1];
      const currentDistance = getTouchDistance(touchA, touchB);
      const ratio = currentDistance / pinchStartDistanceRef.current;
      const nextZoom = Math.max(MIN_PREVIEW_ZOOM, Math.min(MAX_PREVIEW_ZOOM, pinchStartZoomRef.current * ratio));
      setPreviewZoom(nextZoom);

      const center = getTouchCenter(touchA, touchB);
      const dx = center.x - pinchStartCenterRef.current.x;
      const dy = center.y - pinchStartCenterRef.current.y;
      wrap.scrollLeft = pinchStartScrollRef.current.left - dx;
      wrap.scrollTop = pinchStartScrollRef.current.top - dy;
      return;
    }

    evt.preventDefault();

    const currentPoint: IPoint = getRelativePoint(evt);
    if (applyResizePlacement(currentPoint)) {
      return;
    }

    applyDraftPlacement(currentPoint);
  }, [applyDraftPlacement, applyResizePlacement]);

  const finalizePreviewSelection = React.useCallback((placementOverride?: ISignaturePlacement): void => {
    if (!isSelectingPlacement) {
      return;
    }

    setIsSelectingPlacement(false);
    dragStartPointRef.current = null;

    const placement: ISignaturePlacement | null = placementOverride || draftPlacement;
    if (!placement) {
      return;
    }

    setSignaturePlacement(placement);
    setDraftPlacement(null);
    setIsSignatureModalOpen(true);
  }, [draftPlacement, isSelectingPlacement]);

  const handlePreviewMouseUp = React.useCallback((evt: React.MouseEvent<HTMLCanvasElement>): void => {
    if (isResizingPlacement) {
      finishResizePlacement();
      return;
    }

    if (isSelectingPlacement && dragStartPointRef.current) {
      const previewCanvas: HTMLCanvasElement | null = previewCanvasRef.current;
      if (previewCanvas) {
        const currentPoint: IPoint = getRelativePoint(evt);
        const placement: ISignaturePlacement = createPlacementFromPoints(dragStartPointRef.current, currentPoint, previewCanvas);
        setDraftPlacement(placement);
        finalizePreviewSelection(placement);
        return;
      }
    }

    finalizePreviewSelection();
  }, [finalizePreviewSelection, finishResizePlacement, isResizingPlacement]);

  const handlePreviewTouchEnd = React.useCallback((evt: React.TouchEvent<HTMLCanvasElement>): void => {
    if (isPinchPanningRef.current) {
      isPinchPanningRef.current = false;
      pinchStartCenterRef.current = null;
      pinchStartScrollRef.current = null;
      return;
    }

    if (isResizingPlacement) {
      finishResizePlacement();
      return;
    }

    if (evt.changedTouches.length > 0) {
      const lastTouch = evt.changedTouches[0];
      const currentPoint: IPoint | undefined = getPreviewPointFromClient(lastTouch.clientX, lastTouch.clientY);
      if (currentPoint) {
        const previewCanvas: HTMLCanvasElement | null = previewCanvasRef.current;
        if (previewCanvas && dragStartPointRef.current) {
          const placement: ISignaturePlacement = createPlacementFromPoints(dragStartPointRef.current, currentPoint, previewCanvas);
          setDraftPlacement(placement);
          finalizePreviewSelection(placement);
          return;
        }

        applyDraftPlacement(currentPoint);
      }
    }

    finalizePreviewSelection();
  }, [applyDraftPlacement, finalizePreviewSelection, finishResizePlacement, getPreviewPointFromClient, isResizingPlacement]);

  const handlePreviewMouseLeave = React.useCallback((): void => {
    if (isResizingPlacement) {
      finishResizePlacement();
      return;
    }

    finalizePreviewSelection();
  }, [finalizePreviewSelection, finishResizePlacement, isResizingPlacement]);

  const handleClearPlacedSignature = React.useCallback((): void => {
    setSignatureImageDataUrl('');
    setSignerEmail('');
    setSignerTimestamp('');
    setSignaturePlacement(null);
    setDraftPlacement(null);
    setIsSelectingPlacement(false);
    dragStartPointRef.current = null;
    finishResizePlacement();
  }, [finishResizePlacement]);

  const uploadSignedPdf = React.useCallback(async (destinationTarget: ILibraryTarget, bytes: Uint8Array, sourceFileName: string): Promise<void> => {
    const uploadBody: ArrayBuffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
    const uploadApiUrl: string = createApiUrl(
      destinationTarget.webAbsoluteUrl,
      `/_api/web/GetFolderByServerRelativePath(decodedUrl='${encodeODataValue(destinationTarget.folderServerRelativeUrl)}')/Files/add(overwrite=true,url='${encodeODataValue(sourceFileName)}')`
    );

    const uploadResponse: SPHttpClientResponse = await props.spHttpClient.post(
      uploadApiUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/octet-stream'
        },
        body: uploadBody
      }
    );

    if (!uploadResponse.ok) {
      throw new Error('Signed PDF created but failed to upload to destination path.');
    }
  }, [props.spHttpClient]);

  const handleSignAndSave = React.useCallback(async (): Promise<void> => {
    setSuccessNotification('');
    setSaveErrorNotification('');
    setIsPostSaveModalOpen(false);

    if (!sourceBytesRef.current || !previewInfoRef.current || !signaturePlacement || !selectedPdfName) {
      setSaveErrorNotification('Please load a PDF and choose a signature coordinate area first.');
      return;
    }

    const destinationPathValue: string = useSameDestinationPath ? sourcePathInput : destinationPathInput;
    const destinationTarget: ILibraryTarget | undefined = parseLibraryUrl(destinationPathValue);
    if (!destinationPathValue.trim() || !destinationTarget) {
      setSaveErrorNotification('Please provide a valid destination full URL path.');
      return;
    }

    if (!signatureImageDataUrl) {
      setSaveErrorNotification('Please draw your signature and click Done in the popup first.');
      return;
    }

    const trimmedEmail: string = (signerEmail || 'Email unavailable').trim();
    const trimmedTimestamp: string = (signerTimestamp || formatSignedAt(new Date())).trim();

    setIsSigning(true);

    try {
      await ensureLibraryExists(destinationTarget);

      const signatureBytes: Uint8Array = dataUrlToBytes(signatureImageDataUrl);
      const pdfDoc = await PDFDocument.load(sourceBytesRef.current);
      const signatureImage = await pdfDoc.embedPng(signatureBytes);
      const emailFont = await pdfDoc.embedFont(StandardFonts.Helvetica);
      const firstPage = pdfDoc.getPage(0);

      const previewCanvasForMapping: HTMLCanvasElement | null = previewCanvasRef.current;
      if (!previewCanvasForMapping || previewCanvasForMapping.width <= 0 || previewCanvasForMapping.height <= 0) {
        throw new Error('Preview canvas is unavailable for coordinate mapping.');
      }

      const pageWidth: number = firstPage.getWidth();
      const pageHeight: number = firstPage.getHeight();
      const xRatio: number = signaturePlacement.x / previewCanvasForMapping.width;
      const yRatio: number = signaturePlacement.y / previewCanvasForMapping.height;
      const widthRatio: number = signaturePlacement.width / previewCanvasForMapping.width;
      const heightRatio: number = signaturePlacement.height / previewCanvasForMapping.height;

      const pdfX: number = xRatio * pageWidth;
      const pdfWidth: number = widthRatio * pageWidth;
      const pdfHeight: number = heightRatio * pageHeight;
      const pdfYFromTop: number = yRatio * pageHeight;
      const drawBoxY: number = pageHeight - pdfYFromTop - pdfHeight + SIGNATURE_OFFSET_Y;

      // Match preview behavior (object-fit: contain) so the drawn signature is not stretched.
      const imageAspectRatio: number = signatureImage.width / signatureImage.height;
      const boxAspectRatio: number = pdfWidth / pdfHeight;
      const drawSignatureWidth: number = imageAspectRatio >= boxAspectRatio ? pdfWidth : pdfHeight * imageAspectRatio;
      const drawSignatureHeight: number = imageAspectRatio >= boxAspectRatio ? pdfWidth / imageAspectRatio : pdfHeight;
      const drawSignatureX: number = pdfX + SIGNATURE_OFFSET_X + Math.max(0, (pdfWidth - drawSignatureWidth) / 2);
      const drawSignatureY: number = drawBoxY + Math.max(0, (pdfHeight - drawSignatureHeight) / 2);

      firstPage.drawImage(signatureImage, {
        x: drawSignatureX,
        y: Math.max(0, drawSignatureY),
        width: drawSignatureWidth,
        height: drawSignatureHeight
      });

      const emailText: string = `Email: ${trimmedEmail}`;
      const signedAtText: string = `Signed at: ${trimmedTimestamp}`;
      const baseEmailFontSize: number = Math.min(12, Math.max(8, pdfHeight * 0.18));
      const maxTextWidth: number = Math.max(10, pdfWidth);
      let emailFontSize: number = baseEmailFontSize;
      const widestMetaWidthAtBase: number = Math.max(
        emailFont.widthOfTextAtSize(emailText, baseEmailFontSize),
        emailFont.widthOfTextAtSize(signedAtText, baseEmailFontSize)
      );

      if (widestMetaWidthAtBase > maxTextWidth) {
        emailFontSize = Math.max(7, (maxTextWidth / widestMetaWidthAtBase) * baseEmailFontSize);
      }

      const emailGap: number = Math.max(3, emailFontSize * 0.3);
      const lineGap: number = Math.max(2, emailFontSize * 0.2);
      const blockHeight: number = emailFontSize * 2 + lineGap;
      const emailTextWidth: number = emailFont.widthOfTextAtSize(emailText, emailFontSize);
      const signedAtTextWidth: number = emailFont.widthOfTextAtSize(signedAtText, emailFontSize);
      const centeredEmailX: number = pdfX + SIGNATURE_OFFSET_X + Math.max(0, (pdfWidth - emailTextWidth) / 2);
      const centeredSignedAtX: number = pdfX + SIGNATURE_OFFSET_X + Math.max(0, (pdfWidth - signedAtTextWidth) / 2);
      const underSignatureY: number = drawBoxY - blockHeight - emailGap;
      const fallbackAboveSignatureY: number = drawBoxY + pdfHeight + emailGap;
      const baseMetaY: number = underSignatureY >= 0 ? underSignatureY : Math.min(pageHeight - blockHeight, fallbackAboveSignatureY);
      const signedAtY: number = baseMetaY;
      const emailY: number = signedAtY + emailFontSize + lineGap;

      firstPage.drawText(emailText, {
        x: centeredEmailX,
        y: Math.max(0, emailY),
        size: emailFontSize,
        font: emailFont,
        color: rgb(0.11, 0.11, 0.11)
      });

      firstPage.drawText(signedAtText, {
        x: centeredSignedAtX,
        y: Math.max(0, signedAtY),
        size: emailFontSize,
        font: emailFont,
        color: rgb(0.11, 0.11, 0.11)
      });

      const signedBytes: Uint8Array = await pdfDoc.save();
      await uploadSignedPdf(destinationTarget, signedBytes, selectedPdfName);

      setSourcePathInput('');
      setSourceStatus('Provide the source document library URL to begin.');
      setPdfFiles([]);
      setSelectedPdfUrl('');
      setSelectedPdfName('');
      setIsLoadingFiles(false);
      setIsLoadingPdf(false);
      setIsPreviewReady(false);
      setPreviewZoom(1);
      setSignaturePlacement(null);
      setDraftPlacement(null);
      setIsSelectingPlacement(false);
      setIsResizingPlacement(false);
      resizeHandleRef.current = null;
      resizeStartPointRef.current = null;
      resizeStartPlacementRef.current = null;
      setSignatureImageDataUrl('');
      setSignerEmail('');
      setSignerTimestamp('');
      setIsSignatureModalOpen(false);
      setUseSameDestinationPath(true);
      setDestinationPathInput('');
      sourceBytesRef.current = null;
      pdfDocumentRef.current = null;
      previewInfoRef.current = null;

      const previewCanvas = previewCanvasRef.current;
      if (previewCanvas) {
        const previewContext = previewCanvas.getContext('2d');
        if (previewContext) {
          previewContext.clearRect(0, 0, previewCanvas.width, previewCanvas.height);
        }
        previewCanvas.width = 0;
        previewCanvas.height = 0;
      }

      setSuccessNotification('Signed PDF generated successfully. Please refresh this page before signing another PDF.');
      setIsPostSaveModalOpen(true);
      setSaveErrorNotification('');
    } catch (error) {
      setSuccessNotification('');
      setIsPostSaveModalOpen(false);
      setSaveErrorNotification(error instanceof Error ? error.message : 'Failed to sign and upload PDF.');
    } finally {
      setIsSigning(false);
    }
  }, [destinationPathInput, ensureLibraryExists, selectedPdfName, signatureImageDataUrl, signaturePlacement, signerEmail, signerTimestamp, sourcePathInput, uploadSignedPdf, useSameDestinationPath]);

  return (
    <section className={styles.pdfESignature}>
      <header className={styles.header}>
        <h2>PDF E-Signature</h2>
        <p className={styles.statusText}>Note: Open a PDF directly with URL parameter `pdf-url`.</p>
        <p className={styles.statusText}>Example: https://yourdomain.sharepoint.com/sites/yoursite/Pages/pdf-signature.aspx?pdf-url=https://yourdomain.sharepoint.com/sites/yoursite/Shared%20Documents/sample.pdf</p>
        {directPdfNotification && (
          <p className={styles.statusText}>{directPdfNotification}</p>
        )}
      </header>

      {!isDirectPdfMode && (
        <div className={styles.card}>
          <label htmlFor="sourceLibraryPath">Source document library full URL</label>
          <div className={styles.inlineRow}>
            <input
              id="sourceLibraryPath"
              type="text"
              value={sourcePathInput}
              onChange={(evt) => setSourcePathInput(evt.target.value)}
              placeholder={`${props.currentWebUrl}/Shared Documents`}
            />
            <button type="button" onClick={handleLoadLibrary} disabled={isLoadingFiles}>
              {isLoadingFiles ? 'Loading...' : 'Load PDF List'}
            </button>
          </div>
          <p className={styles.statusText}>{sourceStatus}</p>
        </div>
      )}

      {isDirectPdfMode && (
        <div className={styles.card}>
          <p className={styles.statusText}>{sourceStatus}</p>
        </div>
      )}

      {pdfFiles.length > 0 && (
        <div className={styles.card}>
          <label htmlFor="pdfSelector">PDF documents in source path</label>
          <select
            id="pdfSelector"
            value={selectedPdfUrl}
            onChange={(evt) => {
              const selectedUrl: string = evt.target.value;
              const selectedFile: IPdfFile | undefined = pdfFiles.find((file) => file.serverRelativeUrl === selectedUrl);

              // Reset current selection/signature state when switching to another source PDF.
              setSignaturePlacement(null);
              setDraftPlacement(null);
              setIsSelectingPlacement(false);
              setIsResizingPlacement(false);
              resizeHandleRef.current = null;
              resizeStartPointRef.current = null;
              resizeStartPlacementRef.current = null;
              dragStartPointRef.current = null;
              setSignatureImageDataUrl('');
              setSignerEmail('');
              setSignerTimestamp('');
              setIsSignatureModalOpen(false);

              setSelectedPdfUrl(selectedUrl);
              setSelectedPdfName(selectedFile ? selectedFile.name : '');
            }}
          >
            {pdfFiles.map((file) => (
              <option key={file.uniqueId} value={file.serverRelativeUrl}>{file.name}</option>
            ))}
          </select>
          {isLoadingPdf && <p className={styles.statusText}>Rendering PDF preview...</p>}
        </div>
      )}

      <div className={styles.previewPanel}>
        <div className={`${styles.card} ${styles.previewCard}`}>
          <h3>PDF Preview</h3>
          <p>Drag on the PDF to draw the signature box. Signature image will be shown inside the selected box after you click Done.</p>
          <p className={styles.statusText}>Minimum draw box size: {SIGNATURE_CANVAS_WIDTH} x {SIGNATURE_CANVAS_HEIGHT}</p>
          <div className={styles.previewWrap} ref={previewWrapRef}>
            <canvas
              ref={previewCanvasRef}
              onMouseDown={handlePreviewMouseDown}
              onMouseMove={handlePreviewMouseMove}
              onMouseUp={handlePreviewMouseUp}
              onMouseLeave={handlePreviewMouseLeave}
              onTouchStart={handlePreviewTouchStart}
              onTouchMove={handlePreviewTouchMove}
              onTouchEnd={handlePreviewTouchEnd}
              onTouchCancel={handlePreviewTouchEnd}
              className={`${styles.previewCanvas} ${signaturePlacement ? styles.previewCanvasWithPlacement : ''}`.trim()}
              style={{
                width: `${(previewCanvasRef.current?.width || 0) * previewZoom}px`,
                height: `${(previewCanvasRef.current?.height || 0) * previewZoom}px`
              }}
            />
            {draftPlacement && (
              <div
                className={`${styles.selectedCoordinateBox} ${styles.draftCoordinateBox}`}
                style={{
                  left: `${draftPlacement.x * (((previewCanvasRef.current?.clientWidth || 1) / (previewCanvasRef.current?.width || 1)))}px`,
                  top: `${draftPlacement.y * (((previewCanvasRef.current?.clientHeight || 1) / (previewCanvasRef.current?.height || 1)))}px`,
                  width: `${draftPlacement.width * (((previewCanvasRef.current?.clientWidth || 1) / (previewCanvasRef.current?.width || 1)))}px`,
                  height: `${draftPlacement.height * (((previewCanvasRef.current?.clientHeight || 1) / (previewCanvasRef.current?.height || 1)))}px`
                }}
              />
            )}
            {signaturePlacement && (
              <div
                className={styles.selectedCoordinateBox}
                style={{
                  left: `${signaturePlacement.x * (((previewCanvasRef.current?.clientWidth || 1) / (previewCanvasRef.current?.width || 1)))}px`,
                  top: `${signaturePlacement.y * (((previewCanvasRef.current?.clientHeight || 1) / (previewCanvasRef.current?.height || 1)))}px`,
                  width: `${signaturePlacement.width * (((previewCanvasRef.current?.clientWidth || 1) / (previewCanvasRef.current?.width || 1)))}px`,
                  height: `${signaturePlacement.height * (((previewCanvasRef.current?.clientHeight || 1) / (previewCanvasRef.current?.height || 1)))}px`
                }}
              />
            )}
            {signaturePlacement && signatureImageDataUrl && (
              <>
                <img
                  src={signatureImageDataUrl}
                  alt="Signature preview"
                  className={styles.signaturePreviewImage}
                  style={{
                    left: `${signaturePlacement.x * (((previewCanvasRef.current?.clientWidth || 1) / (previewCanvasRef.current?.width || 1)))}px`,
                    top: `${signaturePlacement.y * (((previewCanvasRef.current?.clientHeight || 1) / (previewCanvasRef.current?.height || 1)))}px`,
                    width: `${signaturePlacement.width * (((previewCanvasRef.current?.clientWidth || 1) / (previewCanvasRef.current?.width || 1)))}px`,
                    height: `${signaturePlacement.height * (((previewCanvasRef.current?.clientHeight || 1) / (previewCanvasRef.current?.height || 1)))}px`
                  }}
                />
                {(signerEmail || signerTimestamp) && (
                  <div
                    className={styles.signatureMetaPreview}
                    style={{
                      left: `${signaturePlacement.x * (((previewCanvasRef.current?.clientWidth || 1) / (previewCanvasRef.current?.width || 1)))}px`,
                      top: `${(signaturePlacement.y + signaturePlacement.height) * (((previewCanvasRef.current?.clientHeight || 1) / (previewCanvasRef.current?.height || 1)))}px`,
                      width: `${signaturePlacement.width * (((previewCanvasRef.current?.clientWidth || 1) / (previewCanvasRef.current?.width || 1)))}px`
                    }}
                  >
                    <div>Email: {signerEmail || 'Email unavailable'}</div>
                    <div>Signed at: {signerTimestamp || '-'}</div>
                  </div>
                )}
              </>
            )}
          </div>
          {signaturePlacement && (
            <div className={styles.previewActions}>
              <button type="button" className={styles.secondaryButton} onClick={handleClearPlacedSignature}>Clear Placed Signature</button>
            </div>
          )}
        </div>
      </div>

      {signatureImageDataUrl && (
        <div className={styles.card}>
          <label htmlFor="useSameDestinationPath" className={styles.checkboxRow}>
            <input
              id="useSameDestinationPath"
              type="checkbox"
              checked={useSameDestinationPath}
              onChange={(evt) => setUseSameDestinationPath(evt.target.checked)}
            />
            <span>Use same destination document library full URL as source</span>
          </label>

          {useSameDestinationPath ? (
            <div className={styles.destinationSummary}>
              {sourcePathInput || '-'}
            </div>
          ) : (
            <>
              <label htmlFor="destinationLibraryPath">Destination document library full URL</label>
              <div className={styles.inlineRow}>
                <input
                  id="destinationLibraryPath"
                  type="text"
                  value={destinationPathInput}
                  onChange={(evt) => setDestinationPathInput(evt.target.value)}
                  placeholder={`${props.currentWebUrl}/Shared Documents/Signed`}
                />
              </div>
            </>
          )}

          <div className={styles.inlineRow}>
            <button type="button" onClick={handleSignAndSave} disabled={isSigning}>
              {isSigning ? 'Saving...' : 'Generate Signed PDF'}
            </button>
          </div>

          {saveErrorNotification && (
            <p className={styles.warningNotification} role="alert" aria-live="assertive">
              {saveErrorNotification}
            </p>
          )}

          <p className={styles.statusText}><small>Note: If a file with the same name already exists in the destination path, it will be replaced.</small></p>

          
        </div>
      )}

      {isSignatureModalOpen && (
        <div className={styles.modalOverlay} role="presentation">
          <div className={styles.modalCard} role="dialog" aria-modal="true" aria-label="Draw signature" onClick={(evt) => evt.stopPropagation()}>
            <div className={styles.signatureModalFrame}>
              <div className={styles.modalHeader}>
                <h3>Draw Signature</h3>
                <button type="button" className={styles.iconCloseButton} onClick={() => setIsSignatureModalOpen(false)}>Close</button>
              </div>
              <p>Use mouse or touch to draw your signature. Use Close or Done to dismiss this popup.</p>
              <canvas
                ref={signatureCanvasRef}
                className={styles.signatureCanvas}
                onMouseDown={startDraw}
                onMouseMove={draw}
                onMouseUp={stopDraw}
                onMouseLeave={stopDraw}
                onTouchStart={startDraw}
                onTouchMove={draw}
                onTouchEnd={stopDraw}
              />
              <div className={styles.modalActions}>
                <button type="button" className={`${styles.secondaryButton} ${styles.modalActionButton}`} onClick={clearSignaturePad}>Clear Signature</button>
                <button type="button" className={`${styles.primaryButton} ${styles.modalActionButton}`} onClick={handleSignatureDone}>Done</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {isPostSaveModalOpen && (
        <div className={styles.modalOverlay} role="presentation">
          <div
            className={styles.modalCard}
            role="dialog"
            aria-modal="true"
            aria-label="Signed PDF generated"
            onClick={(evt) => evt.stopPropagation()}
            style={{ width: 'min(380px, 92vw)', textAlign: 'center' }}
          >
            <h3>Success</h3>
            <p style={{ margin: '10px auto 12px', maxWidth: '320px' }}>{successNotification}</p>
            <div className={styles.modalActions} style={{ display: 'flex', justifyContent: 'center' }}>
              <button
                type="button"
                className={`${styles.primaryButton} ${styles.modalActionButton}`}
                style={{ width: 'auto', minWidth: '120px', paddingInline: '18px' }}
                onClick={() => {
                  setIsPostSaveModalOpen(false);
                  setSuccessNotification('');
                }}
              >
                OK
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
};

export default PdfESignature;
