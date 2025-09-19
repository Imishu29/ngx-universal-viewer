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
}

// Page change event interface
export interface PageChangeEvent {
  page: number;
  totalPages: number;
  type: 'pdf' | 'word' | 'excel' | 'ppt';
}

@Component({
  selector: 'ngx-universal-file-viewer',
  standalone: true,
  imports: [CommonModule, FormsModule],
  template: `
    <div class="file-viewer-container" [class.loading]="isLoading">
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
          <div class="control-group" *ngIf="toolbarConfig.showNavigation !== false">
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
              <input 
                *ngIf="toolbarConfig.showPageInput !== false"
                type="number" 
                [(ngModel)]="currentPage" 
                (ngModelChange)="goToPage()"
                [min]="1"
                [max]="totalPages"
                class="page-input"
              />
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
            <button (click)="downloadPDF()" title="Download PDF" *ngIf="toolbarConfig.showDownload !== false">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
              </svg>
            </button>
            <button (click)="printPDF()" title="Print PDF" *ngIf="toolbarConfig.showPrint">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M19 8H5c-1.66 0-3 1.34-3 3v6h4v4h12v-4h4v-6c0-1.66-1.34-3-3-3zm-3 11H8v-5h8v5zm3-7c-.55 0-1-.45-1-1s.45-1 1-1 1 .45 1 1-.45 1-1 1zm-1-9H6v4h12V3z"/>
              </svg>
            </button>
          </div>
        </div>

        <!-- Canvas Container -->
        <div class="pdf-canvas-container" #canvasContainer>
          <canvas #pdfCanvas></canvas>
        </div>
      </div>

      <!-- Enhanced Word Document Viewer with A4 Layout -->
      <div class="word-viewer" *ngIf="fileType === 'word' && !isLoading && !errorMessage">
        <div class="word-controls" *ngIf="showToolbar">
          <div class="control-group" *ngIf="toolbarConfig.showNavigation !== false">
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
              <input 
                *ngIf="toolbarConfig.showPageInput !== false"
                type="number" 
                [(ngModel)]="currentWordPage" 
                (ngModelChange)="goToWordPage()"
                [min]="1"
                [max]="totalWordPages"
                class="page-input"
              />
              <span *ngIf="toolbarConfig.showPageInput === false">{{ currentWordPage }}</span>
              / {{ totalWordPages }}
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
            <button (click)="zoomOutWord()" [disabled]="wordZoom <= 0.5" title="Zoom Out">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M19 13H5v-2h14v2z"/>
              </svg>
            </button>
            <select [(ngModel)]="wordZoom" (ngModelChange)="changeWordZoom()" class="zoom-select">
              <option [value]="0.5">50%</option>
              <option [value]="0.75">75%</option>
              <option [value]="1">100%</option>
              <option [value]="1.25">125%</option>
              <option [value]="1.5">150%</option>
              <option [value]="2">200%</option>
            </select>
            <button (click)="zoomInWord()" [disabled]="wordZoom >= 2" title="Zoom In">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/>
              </svg>
            </button>
          </div>

          <div class="control-group">
            <button (click)="downloadWord()" title="Download" *ngIf="toolbarConfig.showDownload !== false">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
              </svg>
            </button>
            <button (click)="printWord()" title="Print" *ngIf="toolbarConfig.showPrint">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
                <path d="M19 8H5c-1.66 0-3 1.34-3 3v6h4v4h12v-4h4v-6c0-1.66-1.34-3-3-3zm-3 11H8v-5h8v5zm3-7c-.55 0-1-.45-1-1s.45-1 1-1 1 .45 1 1-.45 1-1 1zm-1-9H6v4h12V3z"/>
              </svg>
            </button>
          </div>
        </div>
        
        <!-- A4 Page Container -->
        <div class="word-document-container" [style.zoom]="wordZoom">
          <div class="a4-page" [innerHTML]="currentWordPageContent"></div>
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
          <button (click)="downloadExcel()" *ngIf="toolbarConfig.showDownload !== false">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor">
              <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
            </svg>
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
          <button (click)="previousSlide()" [disabled]="currentSlide <= 1" *ngIf="toolbarConfig.showNavigation !== false">Previous Slide</button>
          <span>Slide {{ currentSlide }} of {{ totalSlides }}</span>
          <button (click)="nextSlide()" [disabled]="currentSlide >= totalSlides" *ngIf="toolbarConfig.showNavigation !== false">Next Slide</button>
          <button (click)="downloadPPT()" *ngIf="toolbarConfig.showDownload !== false">Download</button>
        </div>
        <div class="slide-content" [innerHTML]="slideContent"></div>
      </div>
    </div>
  `,
  styles: [`
    .file-viewer-container {
      width: 100%;
      height: 100%;
      min-height: 600px;
      background: #f5f5f5;
      border-radius: 8px;
      overflow: hidden;
      position: relative;
      display: flex;
      flex-direction: column;
    }

    /* Loading Styles */
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

    /* Error Styles */
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

    .error-content h3 {
      color: #e74c3c;
      margin: 20px 0 10px;
    }

    .error-content p {
      color: #666;
      margin-bottom: 20px;
      font-size: 14px;
      line-height: 1.5;
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
      transform: translateY(-2px);
      box-shadow: 0 4px 8px rgba(0,0,0,0.15);
    }

    /* PDF Viewer Styles */
    .pdf-viewer {
      height: 100%;
      display: flex;
      flex-direction: column;
    }

    .pdf-controls {
      background: white;
      padding: 12px 20px;
      border-bottom: 1px solid #e0e0e0;
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 20px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }

    .control-group {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .pdf-controls button, .word-controls button, .excel-controls button {
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
      gap: 6px;
    }

    .pdf-controls button:hover:not(:disabled),
    .word-controls button:hover:not(:disabled),
    .excel-controls button:hover:not(:disabled) {
      background: #f8f9fa;
      border-color: #3498db;
      color: #3498db;
      transform: translateY(-1px);
    }

    .pdf-controls button:disabled,
    .word-controls button:disabled {
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

    .pdf-canvas-container {
      flex: 1;
      overflow: auto;
      display: flex;
      justify-content: center;
      padding: 20px;
      background: #e8e8e8;
      background-image: 
        linear-gradient(45deg, #f0f0f0 25%, transparent 25%),
        linear-gradient(-45deg, #f0f0f0 25%, transparent 25%),
        linear-gradient(45deg, transparent 75%, #f0f0f0 75%),
        linear-gradient(-45deg, transparent 75%, #f0f0f0 75%);
      background-size: 20px 20px;
      background-position: 0 0, 0 10px, 10px -10px, -10px 0px;
    }

    canvas {
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
      background: white;
      max-width: 100%;
      height: auto;
      display: block;
    }

    /* Enhanced Word Viewer Styles for A4 Layout */
    .word-viewer {
      height: 100%;
      display: flex;
      flex-direction: column;
      background: #e5e5e5;
    }

    .word-controls {
      background: white;
      padding: 12px 20px;
      border-bottom: 1px solid #e0e0e0;
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 20px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }

    .word-document-container {
      flex: 1;
      overflow: auto;
      padding: 20px;
      display: flex;
      justify-content: center;
      align-items: flex-start;
      background: #e5e5e5;
    }

    /* A4 Page Styling - Fixed Height for proper page display */
    .a4-page {
      width: 794px; /* A4 width at 96 DPI */
      height: 1123px; /* Fixed A4 height at 96 DPI */
      padding: 72px; /* 1 inch margins */
      background: white;
      box-shadow: 0 0 10px rgba(0, 0, 0, 0.2);
      margin: 0 auto;
      box-sizing: border-box;
      font-family: 'Calibri', 'Arial', sans-serif;
      font-size: 11pt;
      line-height: 1.6;
      color: #000;
      position: relative;
      overflow-y: auto;
      overflow-x: hidden;
    }

    /* Word Document Typography */
    .a4-page h1 {
      font-size: 16pt;
      font-weight: bold;
      margin: 0 0 12pt 0;
      color: #2E74B5;
    }

    .a4-page h2 {
      font-size: 14pt;
      font-weight: bold;
      margin: 12pt 0 6pt 0;
      color: #2E74B5;
    }

    .a4-page h3 {
      font-size: 12pt;
      font-weight: bold;
      margin: 12pt 0 6pt 0;
      color: #1F497D;
    }

    .a4-page p {
      margin: 0 0 12pt 0;
      text-align: justify;
      word-wrap: break-word;
    }

    .a4-page ul, .a4-page ol {
      margin: 0 0 12pt 0;
      padding-left: 36pt;
    }

    .a4-page li {
      margin-bottom: 6pt;
    }

    .a4-page table {
      border-collapse: collapse;
      width: 100%;
      margin: 12pt 0;
    }

    .a4-page table td, .a4-page table th {
      border: 1px solid #000;
      padding: 6pt;
    }

    .a4-page table th {
      background: #E7E6E6;
      font-weight: bold;
    }

    /* Image handling in Word documents */
    .a4-page img {
      max-width: 100%;
      height: auto;
      display: block;
      margin: 12pt auto;
    }

    /* Page break marker */
    .page-break-marker {
      display: block;
      width: 100%;
      height: 1px;
      background: #ccc;
      margin: 20px 0;
      position: relative;
    }

    .page-break-marker::after {
      content: 'Page Break';
      position: absolute;
      top: -10px;
      left: 50%;
      transform: translateX(-50%);
      background: white;
      padding: 0 10px;
      color: #999;
      font-size: 10pt;
    }

    /* Excel Viewer Styles */
    .excel-viewer {
      height: 100%;
      display: flex;
      flex-direction: column;
    }

    .excel-controls {
      background: white;
      padding: 12px 20px;
      border-bottom: 1px solid #e0e0e0;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 20px;
    }

    .table-wrapper {
      flex: 1;
      overflow: auto;
      background: white;
      padding: 20px;
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

    .table-wrapper tr:hover {
      background: #f8f9fa;
    }

    /* PPT Viewer Styles */
    .ppt-viewer {
      height: 100%;
      display: flex;
      flex-direction: column;
    }

    .ppt-controls {
      background: white;
      padding: 15px;
      border-bottom: 1px solid #e0e0e0;
      display: flex;
      align-items: center;
      gap: 20px;
      justify-content: center;
    }

    .ppt-controls button {
      padding: 8px 20px;
      background: #3498db;
      color: white;
      border: none;
      border-radius: 6px;
      cursor: pointer;
      font-size: 14px;
      transition: all 0.2s ease;
    }

    .ppt-controls button:hover:not(:disabled) {
      background: #2980b9;
      transform: translateY(-1px);
    }

    .ppt-controls button:disabled {
      background: #95a5a6;
      cursor: not-allowed;
      opacity: 0.6;
    }

    .slide-content {
      flex: 1;
      background: white;
      padding: 40px;
      overflow: auto;
      display: flex;
      align-items: center;
      justify-content: center;
    }

    /* Responsive */
    @media (max-width: 850px) {
      .a4-page {
        width: calc(100vw - 40px);
        height: calc((100vw - 40px) * 1.414);
        padding: 40px;
      }
    }

    @media (max-width: 768px) {
      .pdf-controls, .word-controls, .excel-controls {
        padding: 10px;
        gap: 10px;
      }

      .control-group {
        flex-wrap: wrap;
      }

      .pdf-controls button, .word-controls button {
        min-width: 32px;
        height: 32px;
      }

      .page-input {
        width: 40px;
      }

      .table-wrapper,
      .slide-content {
        padding: 20px;
      }
    }

    @media print {
      .word-controls, .pdf-controls, .excel-controls, .ppt-controls {
        display: none;
      }
      
      .word-document-container {
        padding: 0;
        background: white;
      }
      
      .a4-page {
        width: 100%;
        height: auto;
        box-shadow: none;
        padding: 0;
        page-break-after: always;
      }
    }
  `]
})
export class NgxUniversalFileViewerComponent implements OnInit, OnChanges {
  @ViewChild('pdfCanvas', { static: false }) pdfCanvas!: ElementRef<HTMLCanvasElement>;
  @ViewChild('canvasContainer', { static: false }) canvasContainer!: ElementRef<HTMLDivElement>;

  @Input() src!: string | ArrayBuffer | Blob;
  @Input() fileType: 'auto' | 'pdf' | 'word' | 'excel' | 'ppt' = 'auto';
  @Input() showToolbar: boolean = true;
  @Input() toolbarConfig: ToolbarConfig = {};
  @Input() enablePageView: boolean = true;
  @Input() linesPerPage: number = 45; // Average lines per A4 page
  
  @Output() onLoad = new EventEmitter<any>();
  @Output() onError = new EventEmitter<any>();
  @Output() pageChange = new EventEmitter<PageChangeEvent>();

  isLoading = false;
  loadingMessage = 'Loading file...';
  loadingProgress = 0;
  errorMessage = '';
  documentContent: SafeHtml = '';
  currentWordPageContent: SafeHtml = '';
  excelContent: SafeHtml = '';
  slideContent: SafeHtml = '';

  // PDF specific
  pdfDocument: any = null;
  currentPage = 1;
  totalPages = 0;
  scale = 1.0;
  rotation = 0;
  private isBrowser: boolean;
  private pdfLib: any = null;
  private originalFileData: any = null;

  // Enhanced Word specific
  wordContent: string = '';
  wordPages: string[] = [];
  currentWordPage = 1;
  totalWordPages = 1;
  wordZoom = 1;
  
  // Excel specific
  excelSheets: string[] = [];
  currentSheet = 0;
  workbook: any;

  // PPT specific
  currentSlide = 1;
  totalSlides = 1;
  slides: string[] = [];

  constructor(
    private sanitizer: DomSanitizer,
    @Inject(PLATFORM_ID) private platformId: Object
  ) {
    this.isBrowser = isPlatformBrowser(this.platformId);
  }

  ngOnInit() {
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
      if (window.pdfjsLib) {
        this.pdfLib = window.pdfjsLib;
        resolve();
        return;
      }

      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
      
      script.onload = () => {
        if (window.pdfjsLib) {
          this.pdfLib = window.pdfjsLib;
          this.pdfLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
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
      
      setTimeout(() => {
        this.renderPDFPage(1);
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

      let viewport = page.getViewport({ scale: this.scale, rotation: this.rotation });
      
      canvas.height = viewport.height;
      canvas.width = viewport.width;

      const renderContext = {
        canvasContext: context,
        viewport: viewport
      };

      await page.render(renderContext).promise;
      
      this.pageChange.emit({ 
        page: this.currentPage, 
        totalPages: this.totalPages,
        type: 'pdf' 
      });
    } catch (error) {
      console.error('Error rendering PDF page:', error);
      this.errorMessage = 'Failed to render PDF page';
    }
  }

 async loadWord() {
  try {
    this.loadingMessage = 'Processing Word document...';
    this.loadingProgress = 30;

    const arrayBuffer = await this.getArrayBuffer();
    this.loadingProgress = 60;
    
    // Convert Word document to HTML
    const result = await mammoth.convertToHtml({ 
      arrayBuffer
    });
    
    this.loadingProgress = 90;
    
    this.wordContent = result.value;
    
    // Enhanced page splitting that respects actual document pages
    this.splitWordIntoRealPages();
    this.renderWordPage(1);
    
    if (result.messages && result.messages.length > 0) {
      console.warn('Word conversion warnings:', result.messages);
    }
  } catch (error) {
    console.error('Word loading error:', error);
    throw new Error('Failed to load Word document');
  }
}

  splitWordIntoRealPages() {
    // Check for page break indicators in the HTML
    const pageBreakPatterns = [
      /<div[^>]*style="[^"]*page-break[^"]*"[^>]*>/gi,
      /<p[^>]*style="[^"]*page-break[^"]*"[^>]*>/gi,
      /<!-- pagebreak -->/gi,
      /<br\s*\/?>\s*<br\s*\/?>\s*<br\s*\/?>/gi // Multiple line breaks often indicate page breaks
    ];

    let pages: string[] = [];
    let content = this.wordContent;

    // First, try to find explicit page breaks
    let hasPageBreaks = false;
    for (const pattern of pageBreakPatterns) {
      if (pattern.test(content)) {
        hasPageBreaks = true;
        break;
      }
    }

    if (hasPageBreaks) {
      // Split by page breaks
      pages = this.splitByPageBreaks(content);
    } else {
      // Split by calculated content height
      pages = this.splitByCalculatedHeight(content);
    }

    // Ensure we have at least one page
    if (pages.length === 0) {
      pages = [content];
    }

    // Process and clean each page
    this.wordPages = pages.map((page, index) => {
      return this.processAndCleanPage(page, index + 1);
    });

    this.totalWordPages = this.wordPages.length;
    this.currentWordPage = 1;
  }

  splitByPageBreaks(content: string): string[] {
    // Split content by various page break patterns
    const pages: string[] = [];
    
    // Replace all types of page breaks with a unified marker
    let processedContent = content
      .replace(/<div[^>]*style="[^"]*page-break-before:\s*always[^"]*"[^>]*>/gi, '<!--PAGEBREAK-->')
      .replace(/<p[^>]*style="[^"]*page-break-before:\s*always[^"]*"[^>]*>/gi, '<!--PAGEBREAK-->')
      .replace(/<!-- pagebreak -->/gi, '<!--PAGEBREAK-->');

    // Split by the unified marker
    const rawPages = processedContent.split('<!--PAGEBREAK-->');
    
    for (const page of rawPages) {
      if (page.trim()) {
        pages.push(page);
      }
    }

    return pages;
  }

  splitByCalculatedHeight(content: string): string[] {
    const pages: string[] = [];
    
    if (!this.isBrowser) {
      // If not in browser, do simple splitting
      return this.simpleSplitByLines(content);
    }

    // Create a hidden measuring div
    const measureDiv = document.createElement('div');
    measureDiv.style.cssText = `
      position: absolute;
      visibility: hidden;
      width: 650px;
      padding: 72px;
      font-family: 'Calibri', 'Arial', sans-serif;
      font-size: 11pt;
      line-height: 1.6;
    `;
    document.body.appendChild(measureDiv);

    // Parse content into elements
    measureDiv.innerHTML = content;
    const elements = Array.from(measureDiv.children);

    let currentPageContent = '';
    let currentPageHeight = 0;
    const maxPageHeight = 900; // Approximate content height for A4 page

    for (const element of elements) {
      const elementHeight = (element as HTMLElement).offsetHeight;
      
      if (currentPageHeight + elementHeight > maxPageHeight && currentPageContent) {
        // Check if element is a heading - if so, move it to next page
        if (element.tagName.match(/^H[1-6]$/)) {
          pages.push(currentPageContent);
          currentPageContent = element.outerHTML;
          currentPageHeight = elementHeight;
        } else {
          // Add to current page if it fits
          currentPageContent += element.outerHTML;
          pages.push(currentPageContent);
          currentPageContent = '';
          currentPageHeight = 0;
        }
      } else {
        currentPageContent += element.outerHTML;
        currentPageHeight += elementHeight;
      }
    }

    // Add remaining content
    if (currentPageContent) {
      pages.push(currentPageContent);
    }

    document.body.removeChild(measureDiv);
    return pages;
  }

  simpleSplitByLines(content: string): string[] {
    const pages: string[] = [];
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = content;
    
    const elements = Array.from(tempDiv.children);
    let currentPage = '';
    let lineCount = 0;
    
    for (const element of elements) {
      const elementLines = this.estimateLines(element);
      
      if (lineCount + elementLines > this.linesPerPage && currentPage) {
        pages.push(currentPage);
        currentPage = element.outerHTML;
        lineCount = elementLines;
      } else {
        currentPage += element.outerHTML;
        lineCount += elementLines;
      }
    }
    
    if (currentPage) {
      pages.push(currentPage);
    }
    
    return pages;
  }

  estimateLines(element: Element): number {
    const text = element.textContent || '';
    const wordsPerLine = 12; // Average words per line
    const words = text.split(/\s+/).length;
    let lines = Math.ceil(words / wordsPerLine);
    
    // Add extra lines for headings and spacing
    if (element.tagName.match(/^H[1-6]$/)) {
      lines += 2;
    } else if (element.tagName === 'P') {
      lines += 1;
    }
    
    return lines;
  }

  processAndCleanPage(content: string, pageNumber: number): string {
    let processed = content;
    
    // Clean up empty tags
    processed = processed.replace(/<p[^>]*>\s*<\/p>/gi, '');
    processed = processed.replace(/<div[^>]*>\s*<\/div>/gi, '');
    
    // Ensure proper structure
    if (!processed.trim()) {
      processed = `<p style="color: #999; text-align: center;">Page ${pageNumber} - No content</p>`;
    }
    
    // Add classes for styling
    processed = processed.replace(/<table/gi, '<table class="word-table"');
    processed = processed.replace(/<img/gi, '<img class="word-image"');
    
    return processed;
  }

  renderWordPage(pageNum: number) {
    if (this.wordPages[pageNum - 1]) {
      const sanitizedHtml = this.sanitizer.sanitize(1, this.wordPages[pageNum - 1]);
      this.currentWordPageContent = sanitizedHtml || '';
      
      this.pageChange.emit({ 
        page: pageNum, 
        totalPages: this.totalWordPages,
        type: 'word' 
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
      '<div class="ppt-slide"><h1>Welcome to Presentation</h1><p>This is slide 1</p></div>',
      '<div class="ppt-slide"><h2>Key Points</h2><ul><li>Point 1</li><li>Point 2</li><li>Point 3</li></ul></div>',
      '<div class="ppt-slide"><h2>Data Analysis</h2><p>Charts and graphs would appear here</p></div>',
      '<div class="ppt-slide"><h2>Conclusion</h2><p>Summary of the presentation</p></div>',
      '<div class="ppt-slide"><h1>Thank You!</h1><p>Questions?</p></div>'
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
        type: 'ppt' 
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

  // PDF Navigation Controls
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

  // Word Navigation Controls
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

  goToWordPage() {
    if (this.currentWordPage < 1) {
      this.currentWordPage = 1;
    } else if (this.currentWordPage > this.totalWordPages) {
      this.currentWordPage = this.totalWordPages;
    }
    this.renderWordPage(this.currentWordPage);
  }

  // Word Zoom Controls
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

  changeWordZoom() {
    // Zoom is applied via CSS
  }

  // PDF Zoom Controls
  zoomIn() {
    if (this.scale < 3) {
      this.scale += 0.25;
      this.renderPDFPage(this.currentPage);
    }
  }

  zoomOut() {
    if (this.scale > 0.5) {
      this.scale -= 0.25;
      this.renderPDFPage(this.currentPage);
    }
  }

  changeZoom() {
    this.renderPDFPage(this.currentPage);
  }

  fitToWidth() {
    if (this.canvasContainer?.nativeElement) {
      const containerWidth = this.canvasContainer.nativeElement.clientWidth - 40;
      this.pdfDocument.getPage(this.currentPage).then((page: any) => {
        const viewport = page.getViewport({ scale: 1 });
        this.scale = containerWidth / viewport.width;
        this.renderPDFPage(this.currentPage);
      });
    }
  }

  // PDF Rotation
  rotate(degrees: number) {
    this.rotation = (this.rotation + degrees) % 360;
    this.renderPDFPage(this.currentPage);
  }

  // Download functions
  async downloadPDF() {
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
    console.log('Download PowerPoint presentation');
  }

  // Print functions
  printPDF() {
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
    const printWindow = window.open('', '_blank');
    if (printWindow) {
      printWindow.document.write(`
        <html>
          <head>
            <title>Print Document</title>
            <style>
              @media print {
                body { margin: 0; font-family: 'Calibri', 'Arial', sans-serif; }
                .page-break { page-break-after: always; }
              }
            </style>
          </head>
          <body>
      `);
      
      // Print all pages
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

  // Excel Controls
  onSheetChange() {
    this.renderExcelSheet(this.currentSheet);
  }

  // PPT Controls
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