import { Component, Input, OnInit, ViewChild, ElementRef, OnChanges, SimpleChanges, Output, EventEmitter, Inject, PLATFORM_ID } from '@angular/core';
import { CommonModule, isPlatformBrowser } from '@angular/common';
import { DomSanitizer, SafeHtml } from '@angular/platform-browser';
import { FormsModule } from '@angular/forms';
import * as mammoth from 'mammoth';
import * as XLSX from 'xlsx';

// PDF.js types
declare global {
  interface Window {
    pdfjsLib: any;
  }
}

// View mode types
export type ViewMode = 'continuous' | 'page';

// Toolbar configuration interface
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

// Viewer configuration
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
}

// Page change event interface
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
    <!-- Template remains the same as previous -->
    <div class="file-viewer-container" 
         [class.loading]="isLoading" 
         [class.continuous-mode]="viewMode === 'continuous'" 
         [class.page-mode]="viewMode === 'page'"
         [style.height]="viewerConfig.height || '100vh'">
      
      <!-- Loading Spinner -->
      <div class="loader-wrapper" *ngIf="isLoading">
        <div class="loader">
          <div class="spinner"></div>
          <p>{{ loadingMessage }}</p>
          <div class="progress-bar" *ngIf="loadingProgress > 0">
            <div class="progress-fill" [style.width.%]="loadingProgress"></div>
          </div>
        </div>
      </div>

      <!-- Error Message -->
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

      <!-- PDF Viewer -->
      <div class="pdf-viewer" *ngIf="fileType === 'pdf' && !isLoading && !errorMessage">
        
        <!-- Top Controls -->
        <div class="pdf-controls" *ngIf="showToolbar">
          
          <!-- View Mode Toggle -->
          <div class="control-group" *ngIf="toolbarConfig.showViewModeToggle !== false">
            <button (click)="toggleViewMode()" class="view-mode-btn" [title]="viewMode === 'continuous' ? 'Switch to Page View' : 'Switch to Continuous View'">
              <svg *ngIf="viewMode === 'continuous'" width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M3 3v8h8V3H3zm6 6H5V5h4v4zm-6 4v8h8v-8H3zm6 6H5v-4h4v4zm4-16v8h8V3h-8zm6 6h-4V5h4v4zm-6 4v8h8v-8h-8zm6 6h-4v-4h4v4z"/>
              </svg>
              <svg *ngIf="viewMode === 'page'" width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M3 9h18v2H3V9zm0 4h18v2H3v-2z"/>
              </svg>
              {{ viewMode === 'continuous' ? 'Page View' : 'Continuous' }}
            </button>
          </div>

          <div class="control-group" *ngIf="toolbarConfig.showNavigation !== false && viewMode === 'page'">
            <button (click)="firstPage()" [disabled]="currentPage <= 1" title="First Page">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M18.41 16.59L13.82 12l4.59-4.59L17 6l-6 6 6 6zM6 6h2v12H6z"/>
              </svg>
            </button>
            <button (click)="previousPage()" [disabled]="currentPage <= 1" title="Previous Page">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M15.41 7.41L14 6l-6 6 6 6 1.41-1.41L10.83 12z"/>
              </svg>
            </button>
            <span class="page-info">
              <input *ngIf="toolbarConfig.showPageInput !== false"
                     type="number"
                     [(ngModel)]="currentPage"
                     (ngModelChange)="goToPage()"
                     [min]="1"
                     [max]="totalPages"
                     class="page-input" />
              <span *ngIf="toolbarConfig.showPageInput === false">{{ currentPage }}</span>
              / {{ totalPages }}
            </span>
            <button (click)="nextPage()" [disabled]="currentPage >= totalPages" title="Next Page">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M10 6L8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z"/>
              </svg>
            </button>
            <button (click)="lastPage()" [disabled]="currentPage >= totalPages" title="Last Page">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M5.59 7.41L10.18 12l-4.59 4.59L7 18l6-6-6-6zM16 6h2v12h-2z"/>
              </svg>
            </button>
          </div>

          <div class="control-group" *ngIf="toolbarConfig.showZoom !== false">
            <button (click)="zoomOut()" [disabled]="scale <= 0.5" title="Zoom Out">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M19 13H5v-2h14v2z"/>
              </svg>
            </button>
            <select [(ngModel)]="scale" (ngModelChange)="changeZoom()" class="zoom-select">
              <option [value]="0.5">50%</option>
              <option [value]="0.75">75%</option>
              <option [value]="1">100%</option>
              <option [value]="1.25">125%</option>
              <option [value]="1.5">150%</option>
              <option [value]="2">200%</option>
              <option [value]="3">300%</option>
            </select>
            <button (click)="zoomIn()" [disabled]="scale >= 3" title="Zoom In">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/>
              </svg>
            </button>
            <button (click)="fitToWidth()" title="Fit to Width" *ngIf="toolbarConfig.showFitToWidth !== false">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M9 3L5 7l4 4V8h8v3l4-4-4-4v3H9V3zm0 18l4-4-4-4v3H1v2h8v3zm10-7v3h-8v-3l-4 4 4 4v-3h8v3l4-4-4-4z"/>
              </svg>
            </button>
          </div>

          <div class="control-group">
            <button (click)="rotate(-90)" title="Rotate Left" *ngIf="toolbarConfig.showRotation !== false">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M7.11 8.53L5.7 7.11C4.8 8.27 4.24 9.61 4.07 11h2.02c.14-.87.49-1.72 1.02-2.47zM6.09 13H4.07c.17 1.39.72 2.73 1.62 3.89l1.41-1.42c-.52-.75-.88-1.59-1.01-2.47zm1.01 5.32c1.16.9 2.51 1.44 3.9 1.61V17.9c-.87-.15-1.71-.49-2.46-1.03L7.1 18.32zM13 4.07V1L8.45 5.55 13 10V6.09c2.84.48 5 2.94 5 5.91s-2.16 5.43-5 5.91v2.02c3.95-.49 7-3.85 7-7.93s-3.05-7.44-7-7.93z"/>
              </svg>
            </button>
            <button (click)="rotate(90)" title="Rotate Right" *ngIf="toolbarConfig.showRotation !== false">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M15.55 5.55L11 1v3.07C7.06 4.56 4 7.92 4 12s3.05 7.44 7 7.93v-2.02c-2.84-.48-5-2.94-5-5.91s2.16-5.43 5-5.91V10l4.55-4.45zM19.93 11c-.17-1.39-.72-2.73-1.62-3.89l-1.42 1.42c.54.75.88 1.6 1.02 2.47h2.02zM13 17.9v2.02c1.39-.17 2.74-.71 3.9-1.61l-1.44-1.44c-.75.54-1.59.89-2.46 1.03zm3.89-2.42l1.42 1.41c.9-1.16 1.45-2.5 1.62-3.89h-2.02c-.14.87-.48 1.72-1.02 2.48z"/>
              </svg>
            </button>
          </div>
        </div>

        <!-- Canvas Container for Page Mode -->
        <div class="pdf-canvas-container" #canvasContainer *ngIf="viewMode === 'page'">
          <canvas #pdfCanvas></canvas>
        </div>

        <!-- Continuous Scroll Container -->
        <div class="pdf-continuous-container" *ngIf="viewMode === 'continuous'" #continuousContainer (scroll)="onContinuousScroll($event)">
          <div *ngFor="let pageNum of pdfPagesArray" class="pdf-page-wrapper" [attr.data-page]="pageNum">
            <div class="page-number">Page {{ pageNum }}</div>
            <canvas [id]="'pdf-page-' + pageNum"></canvas>
          </div>
        </div>
      </div>

      <!-- Word Document Viewer -->
      <div class="word-viewer" *ngIf="fileType === 'word' && !isLoading && !errorMessage">
        <div class="word-controls" *ngIf="showToolbar">
          <div class="control-group" *ngIf="toolbarConfig.showViewModeToggle !== false">
            <button (click)="toggleViewMode()" class="view-mode-btn">
              {{ viewMode === 'continuous' ? 'Page View' : 'Continuous View' }}
            </button>
          </div>

          <div class="control-group" *ngIf="viewMode === 'page'">
            <button (click)="firstWordPage()" [disabled]="currentWordPage <= 1" title="First Page">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M18.41 16.59L13.82 12l4.59-4.59L17 6l-6 6 6 6zM6 6h2v12H6z"/>
              </svg>
            </button>
            <button (click)="previousWordPage()" [disabled]="currentWordPage <= 1" title="Previous Page">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M15.41 7.41L14 6l-6 6 6 6 1.41-1.41L10.83 12z"/>
              </svg>
            </button>
            <span class="page-info">
              {{ currentWordPage }} / {{ totalWordPages }}
            </span>
            <button (click)="nextWordPage()" [disabled]="currentWordPage >= totalWordPages" title="Next Page">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M10 6L8.59 7.41 13.17 12l-4.58 4.59L10 18l6-6z"/>
              </svg>
            </button>
            <button (click)="lastWordPage()" [disabled]="currentWordPage >= totalWordPages" title="Last Page">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M5.59 7.41L10.18 12l-4.59 4.59L7 18l6-6-6-6zM16 6h2v12h-2z"/>
              </svg>
            </button>
          </div>

          <div class="control-group" *ngIf="toolbarConfig.showZoom !== false">
            <button (click)="zoomOutWord()" [disabled]="wordZoom <= 0.5" title="Zoom Out">-</button>
            <select [(ngModel)]="wordZoom" class="zoom-select">
              <option [value]="0.5">50%</option>
              <option [value]="0.75">75%</option>
              <option [value]="1">100%</option>
              <option [value]="1.25">125%</option>
              <option [value]="1.5">150%</option>
            </select>
            <button (click)="zoomInWord()" [disabled]="wordZoom >= 2" title="Zoom In">+</button>
          </div>
        </div>

        <!-- Page Mode -->
        <div class="word-document-container" [style.zoom]="wordZoom" *ngIf="viewMode === 'page'">
          <div class="a4-page" [innerHTML]="currentWordPageContent"></div>
        </div>

        <!-- Continuous Mode -->
        <div class="word-continuous-container" [style.zoom]="wordZoom" *ngIf="viewMode === 'continuous'">
          <div *ngFor="let page of wordPages; let i = index" class="a4-page continuous-page">
            <div class="page-number">Page {{ i + 1 }}</div>
            <div [innerHTML]="sanitizer.sanitize(1, page)"></div>
          </div>
        </div>
      </div>

      <!-- Excel Viewer -->
      <div class="excel-viewer" *ngIf="fileType === 'excel' && !isLoading && !errorMessage">
        <div class="excel-controls" *ngIf="showToolbar">
          <select [(ngModel)]="currentSheet" (ngModelChange)="onSheetChange()" class="sheet-select">
            <option *ngFor="let sheet of excelSheets; let i = index" [value]="i">
              {{ sheet }}
            </option>
          </select>
          <button (click)="downloadExcel()" *ngIf="viewerConfig.enableDownload !== false">
            Download
          </button>
        </div>
        <div class="table-wrapper">
          <div [innerHTML]="excelContent"></div>
        </div>
      </div>

      <!-- PowerPoint Viewer -->
      <div class="ppt-viewer" *ngIf="fileType === 'ppt' && !isLoading && !errorMessage">
        <div class="ppt-controls" *ngIf="showToolbar">
          <button (click)="previousSlide()" [disabled]="currentSlide <= 1">Previous</button>
          <span>Slide {{ currentSlide }} / {{ totalSlides }}</span>
          <button (click)="nextSlide()" [disabled]="currentSlide >= totalSlides">Next</button>
          <button (click)="downloadPPT()" *ngIf="viewerConfig.enableDownload !== false">Download</button>
        </div>
        <div class="slide-content" [innerHTML]="slideContent"></div>
      </div>
    </div>
  `,
  styles: [`
    /* Styles remain the same as previous */
    .file-viewer-container {
      width: 100%;
      height: 100vh;
      max-height: 100vh;
      background: #f5f5f5;
      border-radius: 8px;
      overflow: hidden;
      position: relative;
      display: flex;
      flex-direction: column;
      box-sizing: border-box;
    }

    .view-mode-btn {
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 6px 12px;
      background: #3498db;
      color: white;
      border: none;
      border-radius: 6px;
      cursor: pointer;
      font-size: 14px;
      transition: all 0.2s ease;
    }

    .view-mode-btn:hover {
      background: #2980b9;
    }

    .loader-wrapper {
      position: absolute;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      display: flex;
      align-items: center;
      justify-content: center;
      background: rgba(255, 255, 255, 0.98);
      z-index: 1000;
    }

    .loader {
      text-align: center;
      padding: 40px;
    }

    .spinner {
      border: 4px solid #e0e0e0;
      border-top: 4px solid #3498db;
      border-radius: 50%;
      width: 50px;
      height: 50px;
      animation: spin 1s linear infinite;
      margin: 0 auto 20px;
    }

    @keyframes spin {
      0% { transform: rotate(0deg); }
      100% { transform: rotate(360deg); }
    }

    .progress-bar {
      width: 200px;
      height: 4px;
      background: #e0e0e0;
      border-radius: 2px;
      overflow: hidden;
      margin-top: 20px;
    }

    .progress-fill {
      height: 100%;
      background: #3498db;
      transition: width 0.3s ease;
    }

    .error-wrapper {
      display: flex;
      align-items: center;
      justify-content: center;
      height: 100%;
      padding: 40px;
    }

    .error-content {
      text-align: center;
      max-width: 400px;
    }

    .error-icon {
      width: 60px;
      height: 60px;
      fill: #e74c3c;
    }

    .retry-btn {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 10px 20px;
      background: #3498db;
      color: white;
      border: none;
      border-radius: 6px;
      cursor: pointer;
      font-size: 14px;
      transition: all 0.3s ease;
    }

    .retry-btn:hover {
      background: #2980b9;
    }

    .pdf-controls, .word-controls, .excel-controls, .ppt-controls {
      background: white;
      padding: 12px 20px;
      border-bottom: 1px solid #e0e0e0;
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 15px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.05);
      flex-shrink: 0;
      min-height: 56px;
    }

    .control-group {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .pdf-controls button,
    .word-controls button,
    .excel-controls button {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 36px;
      height: 36px;
      padding: 0 12px;
      background: white;
      color: #333;
      border: 1px solid #ddd;
      border-radius: 6px;
      cursor: pointer;
      transition: all 0.2s ease;
    }

    .pdf-controls button:hover:not(:disabled),
    .word-controls button:hover:not(:disabled) {
      background: #f8f9fa;
      border-color: #3498db;
      color: #3498db;
    }

    button:disabled {
      opacity: 0.4;
      cursor: not-allowed;
    }

    .page-info {
      display: flex;
      align-items: center;
      gap: 8px;
      font-size: 14px;
      color: #666;
      font-weight: 500;
    }

    .page-input {
      width: 50px;
      padding: 4px 8px;
      border: 1px solid #ddd;
      border-radius: 4px;
      text-align: center;
      font-size: 14px;
    }

    .zoom-select, .sheet-select {
      padding: 6px 10px;
      border: 1px solid #ddd;
      border-radius: 6px;
      background: white;
      font-size: 14px;
      cursor: pointer;
    }

    .pdf-viewer {
      flex: 1;
      display: flex;
      flex-direction: column;
      overflow: hidden;
      height: 100%;
    }

    .pdf-canvas-container {
      flex: 1;
      overflow: auto;
      display: flex;
      justify-content: center;
      align-items: flex-start;
      padding: 20px;
      background: #525659;
      height: calc(100% - 56px);
      box-sizing: border-box;
    }

    canvas {
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
      background: white;
      max-width: 100%;
      height: auto;
      display: block;
      image-rendering: crisp-edges;
      image-rendering: -webkit-optimize-contrast;
      -ms-interpolation-mode: nearest-neighbor;
    }

    .pdf-continuous-container {
      flex: 1;
      overflow-y: auto;
      overflow-x: auto;
      padding: 20px;
      background: #525659;
      height: calc(100% - 56px);
      box-sizing: border-box;
      scroll-behavior: smooth;
    }

    .pdf-continuous-container::-webkit-scrollbar,
    .pdf-canvas-container::-webkit-scrollbar {
      width: 12px;
      height: 12px;
    }

    .pdf-continuous-container::-webkit-scrollbar-track,
    .pdf-canvas-container::-webkit-scrollbar-track {
      background: #3a3d41;
      border-radius: 6px;
    }

    .pdf-continuous-container::-webkit-scrollbar-thumb,
    .pdf-canvas-container::-webkit-scrollbar-thumb {
      background: #6b6f75;
      border-radius: 6px;
      border: 2px solid #3a3d41;
    }

    .pdf-continuous-container::-webkit-scrollbar-thumb:hover,
    .pdf-canvas-container::-webkit-scrollbar-thumb:hover {
      background: #888;
    }

    .pdf-page-wrapper {
      margin-bottom: 20px;
      position: relative;
      display: flex;
      justify-content: center;
    }

    .pdf-page-wrapper canvas {
      display: block;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
      background: white;
      image-rendering: crisp-edges;
      image-rendering: -webkit-optimize-contrast;
    }

    .page-number {
      position: absolute;
      top: 10px;
      right: 10px;
      background: rgba(0, 0, 0, 0.7);
      color: white;
      padding: 4px 8px;
      border-radius: 4px;
      font-size: 12px;
      z-index: 10;
    }

    .word-viewer {
      height: 100%;
      display: flex;
      flex-direction: column;
      background: #e5e5e5;
      overflow: hidden;
    }

    .word-document-container {
      flex: 1;
      overflow: auto;
      padding: 20px;
      display: flex;
      justify-content: center;
      align-items: flex-start;
      background: #e5e5e5;
      height: calc(100% - 56px);
      box-sizing: border-box;
    }

    .word-continuous-container {
      flex: 1;
      overflow-y: auto;
      padding: 20px;
      background: #e5e5e5;
      height: calc(100% - 56px);
      box-sizing: border-box;
    }

    .word-document-container::-webkit-scrollbar,
    .word-continuous-container::-webkit-scrollbar {
      width: 10px;
    }

    .word-document-container::-webkit-scrollbar-track,
    .word-continuous-container::-webkit-scrollbar-track {
      background: #ddd;
      border-radius: 5px;
    }

    .word-document-container::-webkit-scrollbar-thumb,
    .word-continuous-container::-webkit-scrollbar-thumb {
      background: #999;
      border-radius: 5px;
    }

    .word-document-container::-webkit-scrollbar-thumb:hover,
    .word-continuous-container::-webkit-scrollbar-thumb:hover {
      background: #777;
    }

    .a4-page {
      width: 794px;
      min-height: 1123px;
      padding: 72px;
      background: white;
      box-shadow: 0 0 10px rgba(0, 0, 0, 0.2);
      margin: 0 auto;
      box-sizing: border-box;
      font-family: 'Calibri', 'Arial', sans-serif;
      font-size: 11pt;
      line-height: 1.6;
      color: #000;
    }

    .continuous-page {
      margin-bottom: 20px;
      position: relative;
    }

    .excel-viewer {
      height: 100%;
      display: flex;
      flex-direction: column;
      overflow: hidden;
    }

    .table-wrapper {
      flex: 1;
      overflow: auto;
      background: white;
      padding: 20px;
      height: calc(100% - 56px);
      box-sizing: border-box;
    }

    .table-wrapper table {
      border-collapse: collapse;
      width: 100%;
      font-size: 13px;
    }

    .table-wrapper th,
    .table-wrapper td {
      border: 1px solid #ddd;
      padding: 10px;
      text-align: left;
    }

    .table-wrapper th {
      background: #3498db;
      color: white;
      font-weight: 600;
      position: sticky;
      top: 0;
      z-index: 10;
    }

    .ppt-viewer {
      height: 100%;
      display: flex;
      flex-direction: column;
      overflow: hidden;
    }

    .slide-content {
      flex: 1;
      background: white;
      padding: 40px;
      overflow: auto;
      display: flex;
      align-items: center;
      justify-content: center;
      height: calc(100% - 56px);
      box-sizing: border-box;
    }

    @media (max-width: 850px) {
      .a4-page {
        width: calc(100vw - 40px);
        min-height: calc((100vw - 40px) * 1.414);
        padding: 40px;
      }

      .pdf-canvas-container,
      .pdf-continuous-container {
        padding: 10px;
      }
    }

    @media (max-width: 768px) {
      .control-group {
        flex-wrap: wrap;
      }

      .file-viewer-container {
        height: 100vh;
        border-radius: 0;
      }
    }
  `]
})
export class NgxUniversalFileViewerComponent implements OnInit, OnChanges {
  @ViewChild('pdfCanvas', { static: false }) pdfCanvas!: ElementRef<HTMLCanvasElement>;
  @ViewChild('canvasContainer', { static: false }) canvasContainer!: ElementRef<HTMLDivElement>;
  @ViewChild('continuousContainer', { static: false }) continuousContainer!: ElementRef<HTMLDivElement>;

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
  
  documentContent: SafeHtml = '';
  currentWordPageContent: SafeHtml = '';
  excelContent: SafeHtml = '';
  slideContent: SafeHtml = '';

  pdfDocument: any = null;
  currentPage = 1;
  totalPages = 0;
  scale = 1.5;
  rotation = 0;
  pdfPagesArray: number[] = [];
  private pdfLib: any = null;
  private originalFileData: any = null;
  private devicePixelRatio = 1; // Default value for SSR

  wordContent: string = '';
  wordPages: string[] = [];
  currentWordPage = 1;
  totalWordPages = 1;
  wordZoom = 1;

  excelSheets: string[] = [];
  currentSheet = 0;
  workbook: any;

  currentSlide = 1;
  totalSlides = 1;
  slides: string[] = [];

  private isBrowser: boolean;

  constructor(
    public sanitizer: DomSanitizer,
    @Inject(PLATFORM_ID) private platformId: Object
  ) {
    this.isBrowser = isPlatformBrowser(this.platformId);
    
    // Set devicePixelRatio only in browser environment
    if (this.isBrowser && typeof window !== 'undefined') {
      this.devicePixelRatio = window.devicePixelRatio || 1;
    }
  }

  ngOnInit() {
    if (this.viewerConfig.defaultViewMode) {
      this.viewMode = this.viewerConfig.defaultViewMode;
    }

    if (this.isBrowser) {
      this.initializePdfJs().then(() => {
        this.loadFile();
      });
    }
  }

  ngOnChanges(changes: SimpleChanges) {
    if (changes['src'] && !changes['src'].firstChange && this.isBrowser) {
      this.loadFile();
    }
  }

  async initializePdfJs(): Promise<void> {
    if (!this.isBrowser) return;

    return new Promise((resolve, reject) => {
      // Check if window exists (for SSR)
      if (typeof window === 'undefined') {
        reject(new Error('Window is not defined'));
        return;
      }

      if (window.pdfjsLib) {
        this.pdfLib = window.pdfjsLib;
        resolve();
        return;
      }

      const script = document.createElement('script');
      script.src = this.viewerConfig.pdfWorkerSrc || 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
      script.onload = () => {
        if (window.pdfjsLib) {
          this.pdfLib = window.pdfjsLib;
          this.pdfLib.GlobalWorkerOptions.workerSrc = 
            this.viewerConfig.pdfWorkerSrc?.replace('pdf.min.js', 'pdf.worker.min.js') || 
            'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
          resolve();
        } else {
          reject(new Error('PDF.js failed to load'));
        }
      };
      script.onerror = () => {
        reject(new Error('Failed to load PDF.js script'));
      };
      document.head.appendChild(script);
    });
  }

  async loadFile() {
    if (!this.src) {
      this.errorMessage = 'No file source provided';
      return;
    }

    this.isLoading = true;
    this.errorMessage = '';
    this.loadingProgress = 0;

    try {
      if (this.fileType === 'auto') {
        this.detectFileType();
      }

      this.loadingMessage = `Loading ${this.fileType} file...`;

      switch (this.fileType) {
        case 'pdf':
          await this.loadPDF();
          break;
        case 'word':
          await this.loadWord();
          break;
        case 'excel':
          await this.loadExcel();
          break;
        case 'ppt':
          await this.loadPPT();
          break;
        default:
          throw new Error('Unsupported file type');
      }

      this.onLoad.emit({ type: this.fileType, status: 'success' });
    } catch (error: any) {
      console.error('File load error:', error);
      this.errorMessage = error.message || `Failed to load ${this.fileType} file`;
      this.onError.emit(error);
    } finally {
      this.isLoading = false;
      this.loadingProgress = 0;
    }
  }

  detectFileType() {
    if (typeof this.src === 'string') {
      const extension = this.src.split('.').pop()?.toLowerCase();
      switch (extension) {
        case 'pdf':
          this.fileType = 'pdf';
          break;
        case 'doc':
        case 'docx':
          this.fileType = 'word';
          break;
        case 'xls':
        case 'xlsx':
          this.fileType = 'excel';
          break;
        case 'ppt':
        case 'pptx':
          this.fileType = 'ppt';
          break;
      }
    }
  }

  async loadPDF() {
    if (!this.isBrowser || !this.pdfLib) {
      throw new Error('PDF.js is not available');
    }

    try {
      this.loadingMessage = 'Loading PDF document...';
      this.loadingProgress = 20;

      let pdfData: any;
      if (typeof this.src === 'string') {
        if (this.src.startsWith('data:')) {
          const base64 = this.src.split(',')[1];
          const binaryString = atob(base64);
          const bytes = new Uint8Array(binaryString.length);
          for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
          }
          pdfData = { data: bytes };
        } else {
          pdfData = this.src;
        }
      } else if (this.src instanceof ArrayBuffer) {
        pdfData = { data: new Uint8Array(this.src) };
      } else if (this.src instanceof Blob) {
        const arrayBuffer = await this.src.arrayBuffer();
        pdfData = { data: new Uint8Array(arrayBuffer) };
      }

      this.originalFileData = pdfData;
      this.loadingProgress = 40;

      const loadingTask = this.pdfLib.getDocument(pdfData);
      loadingTask.onProgress = (progress: any) => {
        if (progress.total > 0) {
          this.loadingProgress = 40 + (progress.loaded / progress.total) * 40;
        }
      };

      this.pdfDocument = await loadingTask.promise;
      this.totalPages = this.pdfDocument.numPages;
      this.currentPage = 1;
      this.rotation = 0;
      this.loadingProgress = 90;

      this.pdfPagesArray = Array.from({ length: this.totalPages }, (_, i) => i + 1);

      setTimeout(() => {
        if (this.viewMode === 'continuous') {
          this.renderAllPDFPages();
        } else {
          this.renderPDFPage(1);
        }
      }, 100);

    } catch (error: any) {
      console.error('PDF loading error:', error);
      throw new Error(`Failed to load PDF: ${error.message}`);
    }
  }

  async renderPDFPage(pageNum: number) {
    if (!this.pdfDocument || !this.pdfCanvas?.nativeElement) {
      console.error('PDF document or canvas not ready');
      return;
    }

    try {
      const page = await this.pdfDocument.getPage(pageNum);
      const canvas = this.pdfCanvas.nativeElement;
      const context = canvas.getContext('2d');

      if (!context) {
        throw new Error('Could not get canvas context');
      }

      // Get device pixel ratio (safe for SSR)
      const outputScale = this.isBrowser && typeof window !== 'undefined' 
        ? (window.devicePixelRatio || 1) 
        : 1;

      const viewport = page.getViewport({ 
        scale: this.scale * outputScale, 
        rotation: this.rotation 
      });

      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);

      canvas.style.width = Math.floor(viewport.width / outputScale) + 'px';
      canvas.style.height = Math.floor(viewport.height / outputScale) + 'px';

      // Enable high quality rendering
      context.imageSmoothingEnabled = false;
      // Use type assertion for vendor-specific properties
      (context as any).webkitImageSmoothingEnabled = false;
      (context as any).mozImageSmoothingEnabled = false;
      (context as any).msImageSmoothingEnabled = false;

      const renderContext = {
        canvasContext: context,
        viewport: viewport,
        enableWebGL: true,
        renderInteractiveForms: true
      };

      await page.render(renderContext).promise;

      this.pageChange.emit({
        page: this.currentPage,
        totalPages: this.totalPages,
        type: 'pdf',
        viewMode: this.viewMode
      });

    } catch (error) {
      console.error('Error rendering PDF page:', error);
      this.errorMessage = 'Failed to render PDF page';
    }
  }

  async renderPDFPageToContinuous(pageNum: number) {
    if (!this.pdfDocument) return;

    try {
      const page = await this.pdfDocument.getPage(pageNum);
      const canvas = document.getElementById(`pdf-page-${pageNum}`) as HTMLCanvasElement;
      
      if (!canvas) {
        console.error(`Canvas not found for page ${pageNum}`);
        return;
      }

      const context = canvas.getContext('2d');
      if (!context) return;

      // Get device pixel ratio (safe for SSR)
      const outputScale = this.isBrowser && typeof window !== 'undefined' 
        ? (window.devicePixelRatio || 1) 
        : 1;

      const viewport = page.getViewport({ 
        scale: this.scale * outputScale, 
        rotation: this.rotation 
      });
      
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);

      canvas.style.width = Math.floor(viewport.width / outputScale) + 'px';
      canvas.style.height = Math.floor(viewport.height / outputScale) + 'px';

      // Enable high quality rendering
      context.imageSmoothingEnabled = false;
      (context as any).webkitImageSmoothingEnabled = false;
      (context as any).mozImageSmoothingEnabled = false;
      (context as any).msImageSmoothingEnabled = false;

      const renderContext = {
        canvasContext: context,
        viewport: viewport,
        enableWebGL: true,
        renderInteractiveForms: true
      };

      await page.render(renderContext).promise;
    } catch (error) {
      console.error(`Error rendering page ${pageNum}:`, error);
    }
  }

  async renderAllPDFPages() {
    if (!this.pdfDocument || !this.continuousContainer) {
      return;
    }

    setTimeout(async () => {
      for (let pageNum = 1; pageNum <= this.totalPages; pageNum++) {
        await this.renderPDFPageToContinuous(pageNum);
      }
    }, 100);
  }

  // Rest of the methods remain the same...
  async loadWord() {
    try {
      this.loadingMessage = 'Processing Word document...';
      this.loadingProgress = 30;

      const arrayBuffer = await this.getArrayBuffer();
      this.loadingProgress = 60;

      const result = await mammoth.convertToHtml({ arrayBuffer });
      this.loadingProgress = 90;

      this.wordContent = result.value;
      this.splitWordIntoPages();
      
      if (this.viewMode === 'page') {
        this.renderWordPage(1);
      }

    } catch (error) {
      console.error('Word loading error:', error);
      throw new Error('Failed to load Word document');
    }
  }

  splitWordIntoPages() {
    // Only access document in browser environment
    if (!this.isBrowser) {
      this.wordPages = [this.wordContent];
      this.totalWordPages = 1;
      this.currentWordPage = 1;
      return;
    }

    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = this.wordContent;
    
    const elements = Array.from(tempDiv.children);
    const pages: string[] = [];
    let currentPage = '';
    let currentHeight = 0;
    const maxHeight = 900;

    elements.forEach(element => {
      const elementHtml = element.outerHTML;
      const estimatedHeight = 50;

      if (currentHeight + estimatedHeight > maxHeight && currentPage) {
        pages.push(currentPage);
        currentPage = elementHtml;
        currentHeight = estimatedHeight;
      } else {
        currentPage += elementHtml;
        currentHeight += estimatedHeight;
      }
    });

    if (currentPage) {
      pages.push(currentPage);
    }

    this.wordPages = pages.length > 0 ? pages : [this.wordContent];
    this.totalWordPages = this.wordPages.length;
    this.currentWordPage = 1;
  }

  renderWordPage(pageNum: number) {
    if (this.wordPages[pageNum - 1]) {
      const sanitizedHtml = this.sanitizer.sanitize(1, this.wordPages[pageNum - 1]);
      this.currentWordPageContent = sanitizedHtml || '';
      
      this.pageChange.emit({
        page: pageNum,
        totalPages: this.totalWordPages,
        type: 'word',
        viewMode: this.viewMode
      });
    }
  }

  async loadExcel() {
    try {
      this.loadingMessage = 'Processing Excel spreadsheet...';
      this.loadingProgress = 30;

      const arrayBuffer = await this.getArrayBuffer();
      this.loadingProgress = 60;

      this.workbook = XLSX.read(arrayBuffer, { type: 'array' });
      this.excelSheets = this.workbook.SheetNames;
      this.loadingProgress = 90;

      this.renderExcelSheet(0);

    } catch (error) {
      console.error('Excel loading error:', error);
      throw new Error('Failed to load Excel file');
    }
  }

  renderExcelSheet(sheetIndex: number) {
    if (!this.workbook || !this.excelSheets[sheetIndex]) return;

    const worksheet = this.workbook.Sheets[this.excelSheets[sheetIndex]];
    const html = XLSX.utils.sheet_to_html(worksheet, {
      editable: false,
      header: '<table class="excel-table">',
      footer: '</table>'
    });

    const sanitizedHtml = this.sanitizer.sanitize(1, html);
    this.excelContent = sanitizedHtml || '';
  }

  async loadPPT() {
    this.loadingMessage = 'Processing PowerPoint presentation...';
    this.totalSlides = 5;
    this.currentSlide = 1;
    this.slides = [
      '<div class="ppt-slide"><h1>Slide 1</h1></div>',
      '<div class="ppt-slide"><h1>Slide 2</h1></div>',
      '<div class="ppt-slide"><h1>Slide 3</h1></div>',
      '<div class="ppt-slide"><h1>Slide 4</h1></div>',
      '<div class="ppt-slide"><h1>Slide 5</h1></div>'
    ];
    this.renderSlide(1);
  }

  renderSlide(slideNum: number) {
    if (this.slides[slideNum - 1]) {
      const sanitizedHtml = this.sanitizer.sanitize(1, this.slides[slideNum - 1]);
      this.slideContent = sanitizedHtml || '';
      
      this.pageChange.emit({
        page: slideNum,
        totalPages: this.totalSlides,
        type: 'ppt',
        viewMode: this.viewMode
      });
    }
  }

  async getArrayBuffer(): Promise<ArrayBuffer> {
    if (this.src instanceof ArrayBuffer) {
      return this.src;
    } else if (this.src instanceof Blob) {
      return await this.src.arrayBuffer();
    } else if (typeof this.src === 'string') {
      if (this.src.startsWith('data:')) {
        const base64 = this.src.split(',')[1];
        const binaryString = atob(base64);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }
        return bytes.buffer;
      } else {
        const response = await fetch(this.src);
        return await response.arrayBuffer();
      }
    }
    throw new Error('Invalid source type');
  }

  onContinuousScroll(event: Event) {
    const container = event.target as HTMLElement;
    const scrollPosition = container.scrollTop;
    const pages = container.querySelectorAll('.pdf-page-wrapper');
    
    for (let i = 0; i < pages.length; i++) {
      const page = pages[i] as HTMLElement;
      const pageTop = page.offsetTop - container.offsetTop;
      const pageBottom = pageTop + page.offsetHeight;
      
      if (scrollPosition >= pageTop && scrollPosition < pageBottom) {
        const pageNum = parseInt(page.getAttribute('data-page') || '1');
        if (this.currentPage !== pageNum) {
          this.currentPage = pageNum;
          this.pageChange.emit({
            page: pageNum,
            totalPages: this.totalPages,
            type: 'pdf',
            viewMode: 'continuous'
          });
        }
        break;
      }
    }
  }

  toggleViewMode() {
    this.viewMode = this.viewMode === 'continuous' ? 'page' : 'continuous';
    this.viewModeChange.emit(this.viewMode);

    if (this.fileType === 'pdf') {
      if (this.viewMode === 'continuous') {
        setTimeout(() => this.renderAllPDFPages(), 100);
      } else {
        setTimeout(() => this.renderPDFPage(this.currentPage), 100);
      }
    } else if (this.fileType === 'word' && this.viewMode === 'page') {
      this.renderWordPage(this.currentWordPage);
    }
  }

  firstPage() {
    if (this.currentPage > 1) {
      this.currentPage = 1;
      this.renderPDFPage(1);
    }
  }

  lastPage() {
    if (this.currentPage < this.totalPages) {
      this.currentPage = this.totalPages;
      this.renderPDFPage(this.totalPages);
    }
  }

  previousPage() {
    if (this.currentPage > 1) {
      this.currentPage--;
      this.renderPDFPage(this.currentPage);
    }
  }

  nextPage() {
    if (this.currentPage < this.totalPages) {
      this.currentPage++;
      this.renderPDFPage(this.currentPage);
    }
  }

  goToPage() {
    if (this.currentPage < 1) {
      this.currentPage = 1;
    } else if (this.currentPage > this.totalPages) {
      this.currentPage = this.totalPages;
    }
    this.renderPDFPage(this.currentPage);
  }

  firstWordPage() {
    if (this.currentWordPage > 1) {
      this.currentWordPage = 1;
      this.renderWordPage(1);
    }
  }

  lastWordPage() {
    if (this.currentWordPage < this.totalWordPages) {
      this.currentWordPage = this.totalWordPages;
      this.renderWordPage(this.totalWordPages);
    }
  }

  previousWordPage() {
    if (this.currentWordPage > 1) {
      this.currentWordPage--;
      this.renderWordPage(this.currentWordPage);
    }
  }

  nextWordPage() {
    if (this.currentWordPage < this.totalWordPages) {
      this.currentWordPage++;
      this.renderWordPage(this.currentWordPage);
    }
  }

  zoomIn() {
    if (this.scale < 3) {
      this.scale += 0.25;
      if (this.viewMode === 'continuous') {
        this.renderAllPDFPages();
      } else {
        this.renderPDFPage(this.currentPage);
      }
    }
  }

  zoomOut() {
    if (this.scale > 0.5) {
      this.scale -= 0.25;
      if (this.viewMode === 'continuous') {
        this.renderAllPDFPages();
      } else {
        this.renderPDFPage(this.currentPage);
      }
    }
  }

  changeZoom() {
    if (this.viewMode === 'continuous') {
      this.renderAllPDFPages();
    } else {
      this.renderPDFPage(this.currentPage);
    }
  }

  fitToWidth() {
    if (this.canvasContainer?.nativeElement) {
      const containerWidth = this.canvasContainer.nativeElement.clientWidth - 40;
      this.pdfDocument.getPage(this.currentPage).then((page: any) => {
        const viewport = page.getViewport({ scale: 1 });
        this.scale = containerWidth / viewport.width;
        if (this.viewMode === 'continuous') {
          this.renderAllPDFPages();
        } else {
          this.renderPDFPage(this.currentPage);
        }
      });
    }
  }

  zoomInWord() {
    if (this.wordZoom < 2) {
      this.wordZoom += 0.25;
    }
  }

  zoomOutWord() {
    if (this.wordZoom > 0.5) {
      this.wordZoom -= 0.25;
    }
  }

  rotate(degrees: number) {
    this.rotation = (this.rotation + degrees) % 360;
    if (this.viewMode === 'continuous') {
      this.renderAllPDFPages();
    } else {
      this.renderPDFPage(this.currentPage);
    }
  }

  async downloadPDF() {
    if (!this.viewerConfig.enableDownload && this.viewerConfig.enableDownload !== undefined) {
      return;
    }

    // Only allow downloads in browser environment
    if (!this.isBrowser || typeof window === 'undefined') {
      return;
    }

    if (typeof this.src === 'string') {
      window.open(this.src, '_blank');
    } else {
      let blob: Blob;
      if (this.src instanceof Blob) {
        blob = this.src;
      } else {
        blob = new Blob([this.src], { type: 'application/pdf' });
      }
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'document.pdf';
      a.click();
      URL.revokeObjectURL(url);
    }
  }

  downloadWord() {
    if (!this.viewerConfig.enableDownload && this.viewerConfig.enableDownload !== undefined) {
      return;
    }

    // Only allow downloads in browser environment
    if (!this.isBrowser || typeof window === 'undefined') {
      return;
    }

    if (typeof this.src === 'string') {
      window.open(this.src, '_blank');
    } else {
      let blob: Blob;
      if (this.src instanceof Blob) {
        blob = this.src;
      } else {
        blob = new Blob([this.src], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      }
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'document.docx';
      a.click();
      URL.revokeObjectURL(url);
    }
  }

  downloadExcel() {
    if (!this.viewerConfig.enableDownload && this.viewerConfig.enableDownload !== undefined) {
      return;
    }

    // Only allow downloads in browser environment
    if (!this.isBrowser || typeof window === 'undefined') {
      return;
    }

    if (typeof this.src === 'string') {
      window.open(this.src, '_blank');
    } else {
      let blob: Blob;
      if (this.src instanceof Blob) {
        blob = this.src;
      } else {
        blob = new Blob([this.src], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      }
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'spreadsheet.xlsx';
      a.click();
      URL.revokeObjectURL(url);
    }
  }

  downloadPPT() {
    if (!this.viewerConfig.enableDownload && this.viewerConfig.enableDownload !== undefined) {
      return;
    }
    console.log('Download PowerPoint presentation');
  }

  printPDF() {
    if (!this.viewerConfig.enablePrint && this.viewerConfig.enablePrint !== undefined) {
      return;
    }

    // Only allow printing in browser environment
    if (!this.isBrowser || typeof window === 'undefined') {
      return;
    }

    if (this.pdfCanvas?.nativeElement) {
      const printWindow = window.open('', '_blank');
      if (printWindow) {
        printWindow.document.write('<html><head><title>Print PDF</title></head><body>');
        printWindow.document.write('<img src="' + this.pdfCanvas.nativeElement.toDataURL() + '" style="width:100%;">');
        printWindow.document.write('</body></html>');
        printWindow.document.close();
        printWindow.print();
      }
    }
  }

  printWord() {
    if (!this.viewerConfig.enablePrint && this.viewerConfig.enablePrint !== undefined) {
      return;
    }

    // Only allow printing in browser environment
    if (!this.isBrowser || typeof window === 'undefined') {
      return;
    }

    const printWindow = window.open('', '_blank');
    if (printWindow) {
      printWindow.document.write(`
        <html>
        <head>
          <title>Print Document</title>
          <style>
            @media print {
              body { margin: 0; }
              .page-break { page-break-after: always; }
            }
          </style>
        </head>
        <body>
      `);

      this.wordPages.forEach((page, index) => {
        printWindow.document.write(page);
        if (index < this.wordPages.length - 1) {
          printWindow.document.write('<div class="page-break"></div>');
        }
      });

      printWindow.document.write('</body></html>');
      printWindow.document.close();
      printWindow.print();
    }
  }

  onSheetChange() {
    this.renderExcelSheet(this.currentSheet);
  }

  previousSlide() {
    if (this.currentSlide > 1) {
      this.currentSlide--;
      this.renderSlide(this.currentSlide);
    }
  }

  nextSlide() {
    if (this.currentSlide < this.totalSlides) {
      this.currentSlide++;
      this.renderSlide(this.currentSlide);
    }
  }

  retry() {
    this.errorMessage = '';
    this.loadFile();
  }
}