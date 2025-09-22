
# NgxUniversalFileViewer

[![npm version](https://badge.fury.io/js/ngx-universal-file-viewer.svg)](https://www.npmjs.com/package/ngx-universal-file-viewer)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Angular](https://img.shields.io/badge/Angular-12%2B-red)](https://angular.io/)

A powerful and versatile Angular component for viewing multiple file formats including PDF, Word (DOC/DOCX), Excel (XLS/XLSX), and PowerPoint (PPT/PPTX) files with both continuous scroll and page-by-page view modes.

## ✨ Features

- 📄 **PDF Viewer** - Full-featured PDF viewing with zoom, rotation, navigation
- 📝 **Word Documents** - Display DOC and DOCX files with proper formatting
- 📊 **Excel Spreadsheets** - View XLS and XLSX files with sheet navigation
- 📽️ **PowerPoint Presentations** - View PPT and PPTX slides
- 🔄 **Dual View Modes** - Toggle between continuous scroll and page-by-page view
- 🎨 **Customizable Toolbar** - Configure which controls to display
- 📱 **Responsive Design** - Works seamlessly on desktop and mobile devices
- 🔍 **Auto File Type Detection** - Automatically detects file type from extension
- 🌐 **SSR Compatible** - Works with Angular Universal
- 💪 **TypeScript Support** - Fully typed for better development experience



## 📦 Installation

### Step 1: Install the package

```bash
npm install ngx-universal-file-viewer
```

#### Or Using Yarn:
```bash
yarn add ngx-universal-file-viewer
```

🚀 Getting Started
For Angular 14+ (Standalone Components)

```bash
import { Component } from '@angular/core';
import { NgxUniversalFileViewerComponent   } from 'ngx-universal-file-viewer';

@Component({
  selector: 'app-document-viewer',
  standalone: true,
  imports: [NgxUniversalFileViewerComponent],
  template: `
    <ngx-universal-file-viewer
      [src]="fileUrl"
      [fileType]="'auto'"
      [viewMode]="'continuous'">
    </ngx-universal-file-viewer>
  `
})
export class DocumentViewerComponent {
  fileUrl = 'assets/sample.pdf';
}
```
For Angular 12-13 (Module-based)
```bash 
import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { NgxUniversalFileViewerComponent  } from 'ngx-universal-file-viewer';

import { AppComponent } from './app.component';

@NgModule({
  declarations: [AppComponent],
  imports: [
    BrowserModule,
    NgxUniversalFileViewerModule
  ],
  providers: [],
  bootstrap: [AppComponent]
})
export class AppModule { }
```
📖 Basic Usage
Simple Implementation
```bash
<ngx-universal-file-viewer
  [src]="fileUrl"
      [fileType]="'auto'"
      [viewMode]="'continuous'"
      [showToolbar]="true"
      >
</ngx-universal-file-viewer>
```
TypeScript
```bash
export class AppComponent {
  fileUrl = 'https://example.com/document.pdf';
}
```
With All Options
```bash
<ngx-universal-file-viewer
  [src]="fileSource"
  [fileType]="fileType"
  [viewMode]="viewMode"
  [showToolbar]="showToolbar"
  [viewerConfig]="viewerConfig"
  [toolbarConfig]="toolbarConfig"
  (onLoad)="handleLoad($event)"
  (onError)="handleError($event)"
  (pageChange)="handlePageChange($event)"
  (viewModeChange)="handleViewModeChange($event)">
</ngx-universal-file-viewer>
```

### 📱 Mobile Support
#### The viewer is fully responsive and works on mobile devices:

##### Touch gestures for scrolling
##### Pinch to zoom (PDF)
##### Responsive toolbar
##### Optimized for small screens
##### 🔒 Security
##### Sanitizes HTML content for Word documents
##### Validates file types
##### Secure handling of file sources
##### No external dependencies for sensitive operations


### 🐛 Troubleshooting
##### Issue: PDF not loading
##### Solution: Ensure PDF.js is properly loaded:




#### 🙏 Acknowledgments
##### PDF.js - PDF rendering
##### Mammoth.js - Word document conversion
##### SheetJS - Excel file processing



### 📞 Support
##### For support, email abhishekrout128@gmail.com..

