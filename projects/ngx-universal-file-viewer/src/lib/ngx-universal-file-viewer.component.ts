import {
  Component,
  Input,
  OnInit,
  ViewChild,
  ElementRef,
  OnChanges,
  SimpleChanges,
  Output,
  EventEmitter,
  Inject,
  PLATFORM_ID,
} from '@angular/core';
import { CommonModule, isPlatformBrowser } from '@angular/common';
import { DomSanitizer, SafeHtml, SafeResourceUrl } from '@angular/platform-browser';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx';

// High-fidelity DOCX rendering
import { renderAsync } from 'docx-preview';

// PDF.js types
declare global {
  interface Window { pdfjsLib: any; }
}

export type ViewMode = 'continuous' | 'page';

export interface ToolbarConfig {
  showDownload?: boolean;
  showPrint?: boolean;
  showZoom?: boolean;
  showRotation?: boolean;
  showNavigation?: boolean;
  showPageInput?: boolean;
  showFitToWidth?: boolean;
  showFullscreen?: boolean;
  showSearch?: boolean;
  showThumbnails?: boolean;
  showViewModeToggle?: boolean;
}

export interface ViewerConfig {
  defaultViewMode?: ViewMode;
  enableDownload?: boolean;
  enablePrint?: boolean;
  enableZoom?: boolean;
  enableNavigation?: boolean;
  enableViewModeToggle?: boolean;
  pdfWorkerSrc?: string;
  theme?: 'light' | 'dark';
  height?: string;
  useGoogleDocsViewer?: boolean;
  forceChapterStartOnNewPage?: boolean; // NEW: every Chapter (H1) starts new page
}

export interface PageChangeEvent {
  page: number;
  totalPages: number;
  type: 'pdf' | 'word' | 'excel' | 'ppt';
  viewMode: ViewMode;
}

@Component({
  selector: 'ngx-universal-file-viewer',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div class="file-viewer-container"
         [class.loading]="isLoading"
         [class.continuous-mode]="viewMode === 'continuous'"
         [class.page-mode]="viewMode === 'page'"
         [style.height]="viewerConfig.height || '100vh'">

      <!-- Loader -->
      <div class="loader-wrapper" *ngIf="isLoading">
        <div class="loader">
          <div class="spinner"></div>
          <p>{{ loadingMessage }}</p>
          <div class="progress-bar" *ngIf="loadingProgress > 0">
            <div class="progress-fill" [style.width.%]="loadingProgress"></div>
          </div>
        </div>
      </div>

      <!-- Error -->
      <div class="error-wrapper" *ngIf="errorMessage && !isLoading">
        <div class="error-content">
          <svg class="error-icon" viewBox="0 0 24 24">
            <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>
          </svg>
          <h3>Error Loading File</h3>
          <p>{{ errorMessage }}</p>
          <button (click)="retry()" class="retry-btn">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
              <path d="M17.65 6.35C16.2 4.9 14.21 4 12 4c-4.42 0-7.99 3.58-7.99 8s3.57 8 7.99 8c3.73 0 6.84-2.55 7.73-6h-2.08c-.82 2.33-3.04 4-5.65 4-3.31 0-6-2.69-6-6s2.69-6 6-6c1.66 0 3.14.69 4.22 1.78L13 11h7V4l-2.35 2.35z"/>
            </svg>
            Retry
          </button>
        </div>
      </div>

      <!-- PDF -->
      <div class="pdf-viewer" *ngIf="fileType==='pdf' && !isLoading && !errorMessage">
        <div class="pdf-controls" *ngIf="showToolbar">
          <div class="control-group" *ngIf="toolbarConfig.showViewModeToggle !== false">
            <button (click)="toggleViewMode()" class="view-mode-btn"
                    [title]="viewMode==='continuous' ? 'Switch to Page View' : 'Switch to Continuous View'">
              <svg *ngIf="viewMode==='continuous'" width="20" height="20" viewBox="0 0 24 24"><path d="M3 3v8h8V3H3zm6 6H5V5h4v4zm-6 4v8h8v-8H3zm6 6H5v-4h4v4zm4-16v8h8V3h-8zm6 6h-4V5h4v4zm-6 4v8h8v-8h-8zm6 6h-4v-4h4v4z"/></svg>
              <svg *ngIf="viewMode==='page'" width="20" height="20" viewBox="0 0 24 24"><path d="M3 9h18v2H3V9zm0 4h18v2H3v-2z"/></svg>
              {{ viewMode==='continuous' ? 'Page View' : 'Continuous' }}
            </button>
          </div>

          <div class="control-group" *ngIf="toolbarConfig.showNavigation !== false && viewMode==='page'">
            <button (click)="firstPage()" [disabled]="currentPage<=1" title="First Page">
              <svg width="20" height="20" viewBox="0 0 24 24"><path d="M18.41 16.59L13.82 12l4.59-4.59L17 6l-6 6 6 6zM6 6h2v12H6z"/></svg>
            </button>
            <button (click)="previousPage()" [disabled]="currentPage<=1" title="Previous">
              <svg width="20" height="20" viewBox="0 0 24 24"><path d="M15.41 7.41L14 6l-6 6 6 6 1.41-1.41L10.83 12z"/></svg>
            </button>
            <span class="page-info">
              <input *ngIf="toolbarConfig.showPageInput !== false" type="number" [(ngModel)]="currentPage"
                     (ngModelChange)="goToPage()" [min]="1" [max]="totalPages" class="page-input"/>
              <span *ngIf="toolbarConfig.showPageInput===false">{{ currentPage }}</span> / {{ totalPages }}
            </span>
            <button (click)="nextPage()" [disabled]="currentPage>=totalPages" title="Next">
              <svg width="20" height="20" viewBox="0 0 24 24"><path d="M10 6L8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z"/></svg>
            </button>
            <button (click)="lastPage()" [disabled]="currentPage>=totalPages" title="Last Page">
              <svg width="20" height="20" viewBox="0 0 24 24"><path d="M5.59 7.41L10.18 12l-4.59 4.59L7 18l6-6-6-6zM16 6h2v12h-2z"/></svg>
            </button>
          </div>

          <div class="control-group" *ngIf="toolbarConfig.showZoom !== false">
            <button (click)="zoomOut()" [disabled]="scale<=0.5" title="Zoom Out"><svg width="20" height="20" viewBox="0 0 24 24"><path d="M19 13H5v-2h14v2z"/></svg></button>
            <select [(ngModel)]="scale" (ngModelChange)="changeZoom()" class="zoom-select">
              <option [value]="0.5">50%</option><option [value]="0.75">75%</option><option [value]="1">100%</option>
              <option [value]="1.25">125%</option><option [value]="1.5">150%</option><option [value]="2">200%</option><option [value]="3">300%</option>
            </select>
            <button (click)="zoomIn()" [disabled]="scale>=3" title="Zoom In"><svg width="20" height="20" viewBox="0 0 24 24"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg></button>
            <button (click)="fitToWidth()" title="Fit to Width" *ngIf="toolbarConfig.showFitToWidth !== false">
              <svg width="20" height="20" viewBox="0 0 24 24"><path d="M9 3L5 7l4 4V8h8v3l4-4-4-4v3H9V3zm0 18l4-4-4-4v3H1v2h8v3zm10-7v3h-8v-3l-4 4 4 4v-3h8v3l4-4-4-4z"/></svg>
            </button>
          </div>

          <div class="control-group">
            <button (click)="rotate(-90)" title="Rotate Left" *ngIf="toolbarConfig.showRotation !== false">
              <svg width="20" height="20" viewBox="0 0 24 24"><path d="M7.11 8.53L5.7 7.11C4.8 8.27 4.24 9.61 4.07 11h2.02c.14-.87.49-1.72 1.02-2.47zM6.09 13H4.07c.17 1.39.72 2.73 1.62 3.89l1.41-1.42c-.52-.75-.88-1.59-1.01-2.47zm1.01 5.32c1.16.9 2.51 1.44 3.9 1.61V17.9c-.87-.15-1.71-.49-2.46-1.03L7.1 18.32zM13 4.07V1L8.45 5.55 13 10V6.09c2.84.48 5 2.94 5 5.91s-2.16 5.43-5 5.91v2.02c3.95-.49 7-3.85 7-7.93s-3.05-7.44-7-7.93z"/></svg>
            </button>
            <button (click)="rotate(90)" title="Rotate Right" *ngIf="toolbarConfig.showRotation !== false">
              <svg width="20" height="20" viewBox="0 0 24 24"><path d="M15.55 5.55L11 1v3.07C7.06 4.56 4 7.92 4 12s3.05 7.44 7 7.93v-2.02c-2.84-.48-5-2.94-5-5.91s2.16-5.43 5-5.91V10l4.55-4.45zM19.93 11c-.17-1.39-.72-2.73-1.62-3.89l-1.42 1.42c.54.75.88 1.6 1.02 2.47h2.02zM13 17.9v2.02c1.39-.17 2.74-.71 3.9-1.61l-1.44-1.44c-.75.54-1.59.89-2.46 1.03zm3.89-2.42l1.42 1.41c.9-1.16 1.45-2.5 1.62-3.89h-2.02c-.14.87-.48 1.72-1.02 2.48z"/></svg>
            </button>
          </div>
        </div>

        <div class="pdf-canvas-container" #canvasContainer *ngIf="viewMode==='page'">
          <canvas #pdfCanvas></canvas>
        </div>

        <div class="pdf-continuous-container" *ngIf="viewMode==='continuous'" #continuousContainer (scroll)="onContinuousScroll($event)">
          <div *ngFor="let pageNum of pdfPagesArray" class="pdf-page-wrapper" [attr.data-page]="pageNum">
            <div class="page-number">Page {{ pageNum }}</div>
            <canvas [id]="'pdf-page-' + pageNum"></canvas>
          </div>
        </div>
      </div>

      <!-- WORD (docx-preview) -->
      <div class="word-viewer" *ngIf="fileType==='word' && !isLoading && !errorMessage">
        <div class="word-controls" *ngIf="showToolbar">
          <div class="control-group" *ngIf="toolbarConfig.showViewModeToggle !== false">
            <button (click)="toggleViewMode()" class="view-mode-btn">
              <svg *ngIf="viewMode==='continuous'" width="20" height="20" viewBox="0 0 24 24"><path d="M3 3v8h8V3H3zm6 6H5V5h4v4zm-6 4v8h8v-8H3zm6 6H5v-4h4v4zm4-16v8h8V3h-8zm6 6h-4V5h4v4zm-6 4v8h8v-8h-8zm6 6h-4v-4h4v4z"/></svg>
              <svg *ngIf="viewMode==='page'" width="20" height="20" viewBox="0 0 24 24"><path d="M3 9h18v2H3V9zm0 4h18v2H3v-2z"/></svg>
              {{ viewMode==='continuous' ? 'Page View' : 'Continuous View' }}
            </button>
          </div>

          <div class="control-group" *ngIf="toolbarConfig.showZoom !== false">
            <button (click)="zoomOutWord()" [disabled]="wordZoom<=0.5" title="Zoom Out">
              <svg width="20" height="20" viewBox="0 0 24 24"><path d="M19 13H5v-2h14v2z"/></svg>
            </button>
            <select [(ngModel)]="wordZoom" (ngModelChange)="applyWordZoom()" class="zoom-select">
              <option [value]="0.5">50%</option><option [value]="0.75">75%</option><option [value]="1">100%</option>
              <option [value]="1.25">125%</option><option [value]="1.5">150%</option><option [value]="1.75">175%</option><option [value]="2">200%</option>
            </select>
            <button (click)="zoomInWord()" [disabled]="wordZoom>=2" title="Zoom In">
              <svg width="20" height="20" viewBox="0 0 24 24"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
            </button>
            <button (click)="fitToPageWidth()" title="Fit to Width">
              <svg width="20" height="20" viewBox="0 0 24 24"><path d="M9 3L5 7l4 4V8h8v3l4-4-4-4v3H9V3zm0 18l4-4-4-4v3H1v2h8v3zm10-7v3h-8v-3l-4 4 4 4v-3h8v3l4-4-4-4z"/></svg>
            </button>
          </div>

         
        </div>

        <!-- If public URL + Google viewer explicitly wanted -->
        <iframe *ngIf="useGoogleViewer && isPublicUrl"
                [src]="googleDocsViewerUrl"
                class="word-iframe-viewer">
        </iframe>

        <!-- High-fidelity DOCX rendering -->
        <div class="word-document-container" #wordRoot *ngIf="!useGoogleViewer">
          <!-- docx-preview will render pages into this root -->
          <div class="docx-wrapper"
               [style.transform]="'scale(' + wordZoom + ')'">
          </div>
        </div>
      </div>

      <!-- EXCEL -->
      <div class="excel-viewer" *ngIf="fileType==='excel' && !isLoading && !errorMessage">
        <div class="excel-controls" *ngIf="showToolbar">
          <select [(ngModel)]="currentSheet" (ngModelChange)="onSheetChange()" class="sheet-select">
            <option *ngFor="let sheet of excelSheets; let i = index" [value]="i">{{ sheet }}</option>
          </select>
          <button (click)="downloadExcel()" *ngIf="viewerConfig.enableDownload !== false">Download</button>
        </div>
        <div class="table-wrapper">
          <div [innerHTML]="excelContent"></div>
        </div>
      </div>

      <!-- PPT (placeholder) -->
      <div class="ppt-viewer" *ngIf="fileType==='ppt' && !isLoading && !errorMessage">
        <div class="ppt-controls" *ngIf="showToolbar">
          <button (click)="previousSlide()" [disabled]="currentSlide<=1">Previous</button>
          <span>Slide {{ currentSlide }} / {{ totalSlides }}</span>
          <button (click)="nextSlide()" [disabled]="currentSlide>=totalSlides">Next</button>
          <button (click)="downloadPPT()" *ngIf="viewerConfig.enableDownload !== false">Download</button>
        </div>
        <div class="slide-content" [innerHTML]="slideContent"></div>
      </div>
    </div>
  `,
  styles: [`
    /* Load metric-compatible fonts so layout matches Word closely */
    @font-face {
      font-family: 'Carlito';
      src: url('https://fonts.cdnfonts.com/s/17883/Carlito-Regular.woff') format('woff');
      font-weight: 400; font-style: normal; font-display: swap;
    }
    @font-face {
      font-family: 'Carlito';
      src: url('https://fonts.cdnfonts.com/s/17883/Carlito-Bold.woff') format('woff');
      font-weight: 700; font-style: normal; font-display: swap;
    }
    @font-face {
      font-family: 'Caladea';
      src: url('https://fonts.cdnfonts.com/s/15362/Caladea-Regular.woff') format('woff');
      font-weight: 400; font-style: normal; font-display: swap;
    }
    @font-face {
      font-family: 'Caladea';
      src: url('https://fonts.cdnfonts.com/s/15362/Caladea-Bold.woff') format('woff');
      font-weight: 700; font-style: normal; font-display: swap;
    }

    .file-viewer-container{width:100%;height:100vh;max-height:100vh;background:#f5f5f5;border-radius:8px;overflow:hidden;position:relative;display:flex;flex-direction:column;box-sizing:border-box}
    /* ===== WORD (docx-preview) ===== */
    .word-document-container{flex:1;overflow:auto;background:#525659;padding:20px;height:calc(100vh - 56px);box-sizing:border-box;display:flex;justify-content:center;align-items:flex-start}
    .docx-wrapper{transform-origin:top center;transition:transform .2s ease}

    /* docx-preview injects .docx, .docx-wrapper, .docx-page */
    .docx { font-family: "Carlito","Caladea","Arial",sans-serif; line-height: 1.15; }
    .docx .docx-page{box-shadow:0 0 10px rgba(0,0,0,0.3);background:#fff;margin:12px auto; position: relative;}

    /* --- High-fidelity image/table sizing --- */
    .docx img, .docx .image { max-width: 100%; height: auto; image-rendering: -webkit-optimize-contrast; }
    .docx table { border-collapse: collapse; }
    .docx table td, .docx table th { vertical-align: top; }

    /* --- Page-break hygiene --- */
    .docx h1, .docx h2, .docx h3 { break-after: avoid; page-break-after: avoid; }
    .docx .chapter-block { break-inside: avoid; page-break-inside: avoid; }
    .docx table, .docx figure, .docx img, .docx .docx-table { break-inside: avoid; page-break-inside: avoid; }
    .docx p { orphans: 2; widows: 2; }

    /* ===== PDF ===== */
    .pdf-viewer{flex:1;display:flex;flex-direction:column;overflow:hidden;height:100%}
    .pdf-canvas-container{flex:1;overflow:auto;display:flex;justify-content:center;align-items:flex-start;padding:20px;background:#525659;height:calc(100vh - 56px);box-sizing:border-box}
    canvas{box-shadow:0 4px 20px rgba(0,0,0,.3);background:#fff;max-width:100%;height:auto;display:block;image-rendering:crisp-edges;image-rendering:-webkit-optimize-contrast}
    .pdf-continuous-container{flex:1;overflow-y:auto;overflow-x:auto;padding:20px;background:#525659;height:calc(100vh - 56px);box-sizing:border-box;scroll-behavior:smooth}
    .pdf-page-wrapper{margin-bottom:20px;position:relative;display:flex;justify-content:center}
    .page-number{position:absolute;top:10px;right:10px;background:rgba(0,0,0,.7);color:#fff;padding:4px 8px;border-radius:4px;font-size:12px;z-index:10}
    /* ===== Controls ===== */
    .pdf-controls,.word-controls,.excel-controls,.ppt-controls{background:#fff;padding:12px 20px;border-bottom:1px solid #e0e0e0;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:15px;box-shadow:0 2px 4px rgba(0,0,0,.05);flex-shrink:0;height:56px;min-height:56px;max-height:56px}
    .control-group{display:flex;align-items:center;gap:8px}
    .pdf-controls button,.word-controls button,.excel-controls button{display:inline-flex;align-items:center;justify-content:center;min-width:36px;height:36px;padding:0 12px;background:#fff;color:#333;border:1px solid #ddd;border-radius:6px;cursor:pointer;transition:all .2s ease}
    .pdf-controls button:hover:not(:disabled),.word-controls button:hover:not(:disabled){background:#f8f9fa;border-color:#3498db;color:#3498db}
    button:disabled{opacity:.4;cursor:not-allowed}
    .page-info{display:flex;align-items:center;gap:8px;font-size:14px;color:#666;font-weight:500}
    .page-input{width:50px;padding:4px 8px;border:1px solid #ddd;border-radius:4px;text-align:center;font-size:14px}
    .zoom-select,.sheet-select{padding:6px 10px;border:1px solid #ddd;border-radius:6px;background:#fff;font-size:14px;cursor:pointer}
    .view-mode-btn{display:flex;align-items:center;gap:6px;padding:6px 12px;background:#3498db;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:14px;transition:all .2s ease}
    .view-mode-btn:hover{background:#2980b9}
    /* Excel */
    .excel-viewer{height:100%;display:flex;flex-direction:column;overflow:hidden}
    .table-wrapper{flex:1;overflow:auto;background:#fff;padding:20px;height:calc(100vh - 56px);box-sizing:border-box}
    .table-wrapper table{border-collapse:collapse;width:100%;font-size:13px}
    .table-wrapper th,.table-wrapper td{border:1px solid #ddd;padding:10px;text-align:left}
    .table-wrapper th{background:#3498db;color:#fff;font-weight:600;position:sticky;top:0;z-index:10}
    /* PPT */
    .ppt-viewer{height:100%;display:flex;flex-direction:column;overflow:hidden}
    .slide-content{flex:1;background:#fff;padding:40px;overflow:auto;display:flex;align-items:center;justify-content:center;height:calc(100vh - 56px);box-sizing:border-box}
    /* Loading */
    .loader-wrapper{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;background:rgba(255,255,255,.98);z-index:1000}
    .loader{text-align:center;padding:40px}
    .spinner{border:4px solid #e0e0e0;border-top:4px solid #3498db;border-radius:50%;width:50px;height:50px;animation:spin 1s linear infinite;margin:0 auto 20px}
    @keyframes spin{0%{transform:rotate(0)}100%{transform:rotate(360deg)}}
    .progress-bar{width:200px;height:4px;background:#e0e0e0;border-radius:2px;overflow:hidden;margin-top:20px}
    .progress-fill{height:100%;background:#3498db;transition:width .3s ease}
    /* Error */
    .error-wrapper{display:flex;align-items:center;justify-content:center;height:100%;padding:40px}
    .error-content{text-align:center;max-width:400px}
    .error-icon{width:60px;height:60px;fill:#e74c3c}
    .retry-btn{display:inline-flex;align-items:center;gap:8px;padding:10px 20px;background:#3498db;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:14px;transition:.3s}
    .retry-btn:hover{background:#2980b9}
    /* Iframe (Google Docs) */
    .word-iframe-viewer{width:100%;height:calc(100vh - 56px);border:none;background:#fff}
    /* Scrollbars */
    .pdf-continuous-container::-webkit-scrollbar,
    .pdf-canvas-container::-webkit-scrollbar,
    .word-document-container::-webkit-scrollbar{width:12px;height:12px}
    .pdf-continuous-container::-webkit-scrollbar-track,
    .pdf-canvas-container::-webkit-scrollbar-track,
    .word-document-container::-webkit-scrollbar-track{background:#3a3d41;border-radius:6px}
    .pdf-continuous-container::-webkit-scrollbar-thumb,
    .pdf-canvas-container::-webkit-scrollbar-thumb,
    .word-document-container::-webkit-scrollbar-thumb{background:#6b6f75;border-radius:6px;border:2px solid #3a3d41}
    .pdf-continuous-container::-webkit-scrollbar-thumb:hover,
    .pdf-canvas-container::-webkit-scrollbar-thumb:hover,
    .word-document-container::-webkit-scrollbar-thumb:hover{background:#888}
    @media (max-width:768px){.file-viewer-container{height:100vh;border-radius:0}}
    @media print{
      .word-controls,.pdf-controls{display:none!important}
      .docx .docx-page{box-shadow:none;margin:0 auto}
      .chapter-block { break-inside: avoid; page-break-inside: avoid; }
    }
  `]
})
export class NgxUniversalFileViewerComponent implements OnInit, OnChanges {
  @ViewChild('pdfCanvas', { static: false }) pdfCanvas!: ElementRef<HTMLCanvasElement>;
  @ViewChild('canvasContainer', { static: false }) canvasContainer!: ElementRef<HTMLDivElement>;
  @ViewChild('continuousContainer', { static: false }) continuousContainer!: ElementRef<HTMLDivElement>;
  @ViewChild('wordRoot', { static: false }) wordRoot!: ElementRef<HTMLDivElement>;

  @Input() src!: string | ArrayBuffer | Blob;
  @Input() fileType: 'auto' | 'pdf' | 'word' | 'excel' | 'ppt' = 'auto';
  @Input() showToolbar: boolean = true;
  @Input() toolbarConfig: ToolbarConfig = {};
  @Input() viewerConfig: ViewerConfig = {};
  @Input() viewMode: ViewMode = 'continuous';

  @Output() onLoad = new EventEmitter<any>();
  @Output() onError = new EventEmitter<any>();
  @Output() pageChange = new EventEmitter<PageChangeEvent>();
  @Output() viewModeChange = new EventEmitter<ViewMode>();

  isLoading = false;
  loadingMessage = 'Loading file...';
  loadingProgress = 0;
  errorMessage = '';

  // PDF
  pdfDocument: any = null;
  currentPage = 1;
  totalPages = 0;
  scale = 1.5;
  rotation = 0;
  pdfPagesArray: number[] = [];
  private pdfLib: any = null;

  // WORD
  useGoogleViewer = false;
  isPublicUrl = false;
  googleDocsViewerUrl: SafeResourceUrl = '';
  wordZoom = 1;
  private docxRendered = false;

  // EXCEL
  excelContent: SafeHtml = '';
  excelSheets: string[] = [];
  currentSheet = 0;
  workbook: any;

  // PPT (placeholder)
  currentSlide = 1;
  totalSlides = 1;
  slideContent: SafeHtml = '';
  slides: string[] = [];

  private isBrowser: boolean;

  constructor(
    public sanitizer: DomSanitizer,
    @Inject(PLATFORM_ID) private platformId: Object
  ) {
    this.isBrowser = isPlatformBrowser(this.platformId);
  }

  ngOnInit() {
    if (this.viewerConfig.defaultViewMode) this.viewMode = this.viewerConfig.defaultViewMode;
    if (this.isBrowser) {
      this.injectFontPreloads();
      this.initializePdfJs().then(() => this.loadFile()).catch(err => console.error('PDF.js init failed:', err));
    }
  }

  ngOnChanges(changes: SimpleChanges) {
    if (changes['src'] && !changes['src'].firstChange && this.isBrowser) {
      this.loadFile();
    }
  }

  /* ======= Helpers ======= */
  private injectFontPreloads() {
    // Preload metric-compatible fonts to reduce reflow
    const links = [
      'https://fonts.cdnfonts.com/s/17883/Carlito-Regular.woff',
      'https://fonts.cdnfonts.com/s/17883/Carlito-Bold.woff',
      'https://fonts.cdnfonts.com/s/15362/Caladea-Regular.woff',
      'https://fonts.cdnfonts.com/s/15362/Caladea-Bold.woff',
    ];
    links.forEach(href => {
      const l = document.createElement('link');
      l.rel = 'preload';
      l.as = 'font';
      l.href = href;
      l.type = 'font/woff';
      l.crossOrigin = 'anonymous';
      document.head.appendChild(l);
    });
  }

  async initializePdfJs(): Promise<void> {
    if (!this.isBrowser) return;
    return new Promise((resolve, reject) => {
      if (typeof window === 'undefined') return reject(new Error('Window not defined'));
      if (window.pdfjsLib) { this.pdfLib = window.pdfjsLib; return resolve(); }

      const script = document.createElement('script');
      script.src = this.viewerConfig.pdfWorkerSrc || 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
      script.onload = () => {
        if (window.pdfjsLib) {
          this.pdfLib = window.pdfjsLib;
          this.pdfLib.GlobalWorkerOptions.workerSrc =
            this.viewerConfig.pdfWorkerSrc?.replace('pdf.min.js', 'pdf.worker.min.js') ||
            'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
          resolve();
        } else reject(new Error('PDF.js failed to load'));
      };
      script.onerror = () => reject(new Error('Failed to load PDF.js'));
      document.head.appendChild(script);
    });
  }

  async loadFile() {
    if (!this.src) { this.errorMessage = 'No file source provided'; return; }
    this.isLoading = true; this.errorMessage = ''; this.loadingProgress = 0;

    try {
      if (this.fileType === 'auto') this.detectFileType();
      this.loadingMessage = `Loading ${this.fileType} file...`;

      switch (this.fileType) {
        case 'pdf': await this.loadPDF(); break;
        case 'word': await this.loadWordHighFidelity(); break;
        case 'excel': await this.loadExcel(); break;
        case 'ppt': await this.loadPPT(); break;
        default: throw new Error('Unsupported file type');
      }
      this.onLoad.emit({ type: this.fileType, status: 'success' });
    } catch (e: any) {
      console.error('File load error:', e);
      this.errorMessage = e?.message || `Failed to load ${this.fileType} file`;
      this.onError.emit(e);
    } finally {
      this.isLoading = false;
      this.loadingProgress = 0;
    }
  }

  detectFileType() {
    if (typeof this.src === 'string') {
      const ext = this.src.split('.').pop()?.toLowerCase();
      if (['pdf'].includes(ext!)) this.fileType = 'pdf';
      else if (['doc','docx'].includes(ext!)) this.fileType = 'word';
      else if (['xls','xlsx'].includes(ext!)) this.fileType = 'excel';
      else if (['ppt','pptx'].includes(ext!)) this.fileType = 'ppt';
    }
  }

  /* ========================= PDF ========================= */
  async loadPDF() {
    if (!this.isBrowser || !this.pdfLib) throw new Error('PDF.js is not available');
    this.loadingMessage = 'Loading PDF document...'; this.loadingProgress = 25;

    let pdfData: any;
    if (typeof this.src === 'string') {
      if (this.src.startsWith('data:')) {
        const base64 = this.src.split(',')[1]; const bin = atob(base64);
        const bytes = new Uint8Array(bin.length); for (let i=0;i<bin.length;i++) bytes[i]=bin.charCodeAt(i);
        pdfData = { data: bytes };
      } else {
        pdfData = this.src;
      }
    } else if (this.src instanceof ArrayBuffer) {
      pdfData = { data: new Uint8Array(this.src) };
    } else if (this.src instanceof Blob) {
      const buf = await this.src.arrayBuffer(); pdfData = { data: new Uint8Array(buf) };
    }

    const task = this.pdfLib.getDocument(pdfData);
    (task as any).onProgress = (p: any) => { if (p.total>0) this.loadingProgress = 25 + (p.loaded/p.total)*60; };
    this.pdfDocument = await task.promise;
    this.totalPages = this.pdfDocument.numPages; this.currentPage = 1; this.rotation = 0; this.loadingProgress = 90;
    this.pdfPagesArray = Array.from({length:this.totalPages}, (_,i)=>i+1);

    setTimeout(() => {
      if (this.viewMode==='continuous') this.renderAllPDFPages();
      else this.renderPDFPage(1);
    }, 50);
  }

  async renderPDFPage(pageNum: number) {
    if (!this.pdfDocument || !this.pdfCanvas?.nativeElement) return;
    const page = await this.pdfDocument.getPage(pageNum);
    const canvas = this.pdfCanvas.nativeElement;
    const ctx = canvas.getContext('2d')!;
    const dpr = (this.isBrowser && typeof window!=='undefined') ? (window.devicePixelRatio || 1) : 1;

    const viewport = page.getViewport({ scale: this.scale * dpr, rotation: this.rotation });
    canvas.width = Math.floor(viewport.width); canvas.height = Math.floor(viewport.height);
    canvas.style.width = Math.floor(viewport.width / dpr) + 'px';
    canvas.style.height = Math.floor(viewport.height / dpr) + 'px';
    ctx.imageSmoothingEnabled = false;

    await page.render({ canvasContext: ctx, viewport, enableWebGL:true, renderInteractiveForms:true }).promise;

    this.pageChange.emit({ page: this.currentPage, totalPages: this.totalPages, type:'pdf', viewMode:this.viewMode });
  }

  async renderPDFPageToContinuous(pageNum: number) {
    if (!this.pdfDocument) return;
    const page = await this.pdfDocument.getPage(pageNum);
    const canvas = document.getElementById(`pdf-page-${pageNum}`) as HTMLCanvasElement;
    if (!canvas) return;
    const ctx = canvas.getContext('2d')!;
    const dpr = (this.isBrowser && typeof window!=='undefined') ? (window.devicePixelRatio || 1) : 1;
    const viewport = page.getViewport({ scale: this.scale * dpr, rotation: this.rotation });
    canvas.width = Math.floor(viewport.width); canvas.height = Math.floor(viewport.height);
    canvas.style.width = Math.floor(viewport.width/dpr)+'px'; canvas.style.height = Math.floor(viewport.height/dpr)+'px';
    ctx.imageSmoothingEnabled = false;
    await page.render({ canvasContext: ctx, viewport, enableWebGL:true, renderInteractiveForms:true }).promise;
  }

  async renderAllPDFPages() {
    if (!this.pdfDocument || !this.continuousContainer) return;
    setTimeout(async () => { for (let p=1;p<=this.totalPages;p++) await this.renderPDFPageToContinuous(p); }, 20);
  }

  onContinuousScroll(event: Event) {
    const container = event.target as HTMLElement;
    const scrollTop = container.scrollTop;
    const pages = container.querySelectorAll('.pdf-page-wrapper');
    for (let i=0;i<pages.length;i++){
      const el = pages[i] as HTMLElement;
      const top = el.offsetTop - container.offsetTop;
      const bottom = top + el.offsetHeight;
      if (scrollTop >= top && scrollTop < bottom){
        const num = parseInt(el.getAttribute('data-page') || '1', 10);
        if (this.currentPage !== num){
          this.currentPage = num;
          this.pageChange.emit({ page:num, totalPages:this.totalPages, type:'pdf', viewMode:'continuous' });
        }
        break;
      }
    }
  }

  toggleViewMode() {
    this.viewMode = (this.viewMode === 'continuous') ? 'page' : 'continuous';
    this.viewModeChange.emit(this.viewMode);
    if (this.fileType === 'pdf') {
      if (this.viewMode==='continuous') setTimeout(()=>this.renderAllPDFPages(), 50);
      else setTimeout(()=>this.renderPDFPage(this.currentPage), 50);
    }
  }

  firstPage(){ if (this.currentPage>1){ this.currentPage=1; this.renderPDFPage(1);} }
  lastPage(){ if (this.currentPage<this.totalPages){ this.currentPage=this.totalPages; this.renderPDFPage(this.totalPages);} }
  previousPage(){ if (this.currentPage>1){ this.currentPage--; this.renderPDFPage(this.currentPage);} }
  nextPage(){ if (this.currentPage<this.totalPages){ this.currentPage++; this.renderPDFPage(this.currentPage);} }
  goToPage(){ if (this.currentPage<1) this.currentPage=1; else if (this.currentPage>this.totalPages) this.currentPage=this.totalPages; this.renderPDFPage(this.currentPage); }

  zoomIn(){ if (this.scale<3){ this.scale+=0.25; (this.viewMode==='continuous')?this.renderAllPDFPages():this.renderPDFPage(this.currentPage);} }
  zoomOut(){ if (this.scale>0.5){ this.scale-=0.25; (this.viewMode==='continuous')?this.renderAllPDFPages():this.renderPDFPage(this.currentPage);} }
  changeZoom(){ (this.viewMode==='continuous')?this.renderAllPDFPages():this.renderPDFPage(this.currentPage); }
  fitToWidth(){
    if (this.canvasContainer?.nativeElement && this.pdfDocument){
      const containerWidth = this.canvasContainer.nativeElement.clientWidth - 40;
      this.pdfDocument.getPage(this.currentPage).then((page: any) => {
        const viewport = page.getViewport({ scale: 1 });
        this.scale = containerWidth / viewport.width;
        (this.viewMode==='continuous')?this.renderAllPDFPages():this.renderPDFPage(this.currentPage);
      });
    }
  }
  rotate(deg:number){ this.rotation = (this.rotation + deg) % 360; (this.viewMode==='continuous')?this.renderAllPDFPages():this.renderPDFPage(this.currentPage); }

  printPDF(){
    if (this.viewerConfig.enablePrint===false) return;
    if (!this.isBrowser) return;
    if (this.pdfCanvas?.nativeElement){
      const w = window.open('', '_blank');
      if (w){
        w.document.write('<html><head><title>Print PDF</title></head><body>');
        w.document.write('<img src="'+this.pdfCanvas.nativeElement.toDataURL()+'" style="width:100%;">');
        w.document.write('</body></html>');
        w.document.close();
        w.print();
      }
    }
  }

  async downloadPDF(){
    if (this.viewerConfig.enableDownload===false) return;
    if (!this.isBrowser) return;
    if (typeof this.src === 'string') window.open(this.src, '_blank');
    else {
      const blob = (this.src instanceof Blob) ? this.src : new Blob([this.src as ArrayBuffer], { type:'application/pdf' });
      const url = URL.createObjectURL(blob); const a = document.createElement('a');
      a.href = url; a.download = 'document.pdf'; a.click(); URL.revokeObjectURL(url);
    }
  }

  /* ========================= WORD (High Fidelity) ========================= */
  async loadWordHighFidelity() {
    this.loadingMessage = 'Rendering Word document...'; this.loadingProgress = 20;

    // Optionally use Google Viewer for PUBLIC links
    if (typeof this.src === 'string' && this.src.startsWith('http') && this.viewerConfig.useGoogleDocsViewer) {
      this.isPublicUrl = true; this.useGoogleViewer = true;
      const encoded = encodeURIComponent(this.src);
      this.googleDocsViewerUrl = this.sanitizer.bypassSecurityTrustResourceUrl(
        `https://docs.google.com/viewer?url=${encoded}&embedded=true`
      );
      this.loadingProgress = 100;
      return;
    }

    const arrayBuffer = await this.getArrayBuffer();
    this.loadingProgress = 60;

    setTimeout(async () => {
      try {
        const container = this.wordRoot?.nativeElement?.querySelector('.docx-wrapper') as HTMLElement | null;
        if (!container) throw new Error('Word container not ready');

        // Clean previous renders
        container.innerHTML = '';

        const options = {
          className: 'docx',
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          experimental: true,
          renderHeaders: true,
          renderFooters: true,
          renderFootnotes: true,
          renderEndnotes: true,
          useBase64URL: true,
          trimXmlDeclaration: true,
        } as any;

        await renderAsync(arrayBuffer, container, undefined, options);

        // --- Group chapters and apply keep-together ---
        this.wrapChapters(container);

        // (Optional) Force every Chapter (H1) to start on a new page
        if (this.viewerConfig.forceChapterStartOnNewPage) {
          const chapters = Array.from(container.querySelectorAll('.chapter-block')) as HTMLElement[];
          chapters.forEach((ch, idx) => {
            if (idx === 0) return;
            ch.style.breakBefore = 'page';
            (ch.style as any)['page-break-before'] = 'always';
          });
        }

        this.docxRendered = true;
        this.wordZoom = 1; // reset zoom
        this.loadingProgress = 100;

        // page count (best-effort): count rendered .docx-page
        const pages = container.querySelectorAll('.docx-page');
        const totalWordPages = Math.max(1, pages.length);
        this.pageChange.emit({ page: 1, totalPages: totalWordPages, type:'word', viewMode:this.viewMode });
      } catch (err:any) {
        console.error('DOCX render error:', err);
        throw new Error('Failed to render Word (DOCX) document');
      }
    }, 0);
  }

  // === Chapter wrapper: keep Chapter content together ===
 private wrapChapters(rootEl: HTMLElement) {
  const pagesRoot = rootEl.querySelector('.docx') as HTMLElement;
  if (!pagesRoot) return;

  const headings = Array.from(pagesRoot.querySelectorAll('h1')) as HTMLElement[];
  if (headings.length === 0) return;

  const wrapRange = (start: Node, endExclusive: Node | null) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'chapter-block';
    let node: Node | null = start;
    const parent = start.parentNode;
    if (!parent) return;

    parent.insertBefore(wrapper, start);
    while (node && node !== endExclusive) {
      const nextNode: Node | null = node.nextSibling; // 👈 typed + renamed
      wrapper.appendChild(node);
      node = nextNode as Node;
    }
  };

  for (let i = 0; i < headings.length; i++) {
    const h = headings[i];
    const nextHeading: Node | null = (headings[i + 1] ?? null); // 👈 typed + renamed
    if (h.parentElement?.classList.contains('chapter-block')) continue;
    wrapRange(h, nextHeading);
  }
}


  zoomInWord(){ if (this.wordZoom < 2){ this.wordZoom += 0.25; this.applyWordZoom(); } }
  zoomOutWord(){ if (this.wordZoom > 0.5){ this.wordZoom -= 0.25; this.applyWordZoom(); } }
  applyWordZoom(){ /* CSS transform on .docx-wrapper handles this */ }

  fitToPageWidth() {
    if (!this.isBrowser) return;
    const container = this.wordRoot?.nativeElement as HTMLElement;
    if (!container) return;
    const page = container.querySelector('.docx-page') as HTMLElement;
    if (!page) return;
    const containerWidth = container.clientWidth - 40;
    const pageWidth = page.clientWidth || 794;
    this.wordZoom = containerWidth / pageWidth;
  }

  downloadWord() {
    if (this.viewerConfig.enableDownload===false) return;
    if (!this.isBrowser) return;
    if (typeof this.src === 'string') window.open(this.src, '_blank');
    else {
      const blob = (this.src instanceof Blob) ? this.src : new Blob([this.src as ArrayBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
      const url = URL.createObjectURL(blob); const a = document.createElement('a');
      a.href = url; a.download = 'document.docx'; a.click(); URL.revokeObjectURL(url);
    }
  }

  printWord() {
    if (this.viewerConfig.enablePrint===false) return;
    if (!this.isBrowser) return;
    const root = this.wordRoot?.nativeElement as HTMLElement;
    if (!root) return;
    const w = window.open('', '_blank');
    if (!w) return;
    const content = root.querySelector('.docx')?.outerHTML || '';
    w.document.write(`
      <html>
        <head>
          <title>Print Document</title>
          <style>
            @media print {
              body { margin: 0; background: white; }
              .docx .docx-page{ page-break-after: always; box-shadow:none; margin:0 auto; }
              .chapter-block { break-inside: avoid; page-break-inside: avoid; }
            }
          </style>
        </head>
        <body>${content}</body>
      </html>
    `);
    w.document.close();
    w.focus();
    w.print();
  }

  /* ========================= EXCEL ========================= */
  async loadExcel() {
    this.loadingMessage = 'Processing Excel spreadsheet...'; this.loadingProgress = 30;
    const arrayBuffer = await this.getArrayBuffer(); this.loadingProgress = 60;
    this.workbook = XLSX.read(arrayBuffer, { type: 'array' });
    this.excelSheets = this.workbook.SheetNames; this.loadingProgress = 90;
    this.renderExcelSheet(0);
  }

  renderExcelSheet(sheetIndex:number){
    if (!this.workbook || !this.excelSheets[sheetIndex]) return;
    const ws = this.workbook.Sheets[this.excelSheets[sheetIndex]];
    const html = XLSX.utils.sheet_to_html(ws, { editable:false, header:'<table class="excel-table">', footer:'</table>' });
    const sanitized = this.sanitizer.sanitize(1, html); this.excelContent = sanitized || '';
  }

  onSheetChange(){ this.renderExcelSheet(this.currentSheet); }

  downloadExcel(){
    if (this.viewerConfig.enableDownload===false) return;
    if (!this.isBrowser) return;
    if (typeof this.src === 'string') window.open(this.src, '_blank');
    else {
      const blob = (this.src instanceof Blob) ? this.src : new Blob([this.src as ArrayBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      const url = URL.createObjectURL(blob); const a = document.createElement('a');
      a.href = url; a.download = 'spreadsheet.xlsx'; a.click(); URL.revokeObjectURL(url);
    }
  }

  /* ========================= PPT (placeholder) ========================= */
  async loadPPT() {
    this.loadingMessage = 'Processing PowerPoint presentation...';
    this.totalSlides = 5; this.currentSlide = 1;
    this.slides = [
      '<div class="ppt-slide"><h1>Slide 1</h1></div>',
      '<div class="ppt-slide"><h1>Slide 2</h1></div>',
      '<div class="ppt-slide"><h1>Slide 3</h1></div>',
      '<div class="ppt-slide"><h1>Slide 4</h1></div>',
      '<div class="ppt-slide"><h1>Slide 5</h1></div>',
    ];
    this.renderSlide(1);
  }

  renderSlide(n:number){
    if (this.slides[n-1]) {
      const sanitized = this.sanitizer.sanitize(1, this.slides[n-1]);
      this.slideContent = sanitized || '';
      this.pageChange.emit({ page:n, totalPages:this.totalSlides, type:'ppt', viewMode:this.viewMode });
    }
  }
  previousSlide(){ if (this.currentSlide>1){ this.currentSlide--; this.renderSlide(this.currentSlide);} }
  nextSlide(){ if (this.currentSlide<this.totalSlides){ this.currentSlide++; this.renderSlide(this.currentSlide);} }
  downloadPPT(){ if (this.viewerConfig.enableDownload===false) return; console.log('Download PPT not implemented'); }

  /* ========================= COMMON ========================= */
  async getArrayBuffer(): Promise<ArrayBuffer> {
    if (this.src instanceof ArrayBuffer) return this.src;
    if (this.src instanceof Blob) return await this.src.arrayBuffer();
    if (typeof this.src === 'string') {
      if (this.src.startsWith('data:')) {
        const base64 = this.src.split(',')[1];
        const bin = atob(base64); const bytes = new Uint8Array(bin.length);
        for (let i=0;i<bin.length;i++) bytes[i]=bin.charCodeAt(i);
        return bytes.buffer;
      } else {
        const res = await fetch(this.src);
        return await res.arrayBuffer();
      }
    }
    throw new Error('Invalid source type');
  }

  retry(){ this.errorMessage=''; this.loadFile(); }
}
