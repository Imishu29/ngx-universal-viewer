import { Component } from '@angular/core';
import { 
  NgxUniversalFileViewerComponent,
  ToolbarConfig,
  ViewerConfig,
  PageChangeEvent 
} from 'ngx-universal-file-viewer';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [NgxUniversalFileViewerComponent],
  template: `
    <div class="app-container">
      <h1>File Viewer Demo</h1>

      <ngx-universal-file-viewer
        [src]="fileSource"
        fileType="auto"
       
       
        [showToolbar]="true"
        [toolbarConfig]="{ showPrint: true, showNavigation: true }"
        (pageChange)="onPageChanged($event)"
      >
      </ngx-universal-file-viewer>

      <div class="comments-section" *ngIf="currentPageInfo">
        <h3>
          Comments for {{ currentPageInfo?.type }} - Page
          {{ currentPageInfo?.page }}
        </h3>
        <!-- Your comment UI here -->
      </div>
    </div>
  `,
  styles: [
    `
      .app-container {
        padding: 20px;
        height: 100vh;
        display: flex;
        flex-direction: column;
      }

      ngx-universal-file-viewer {
        flex: 1;
        margin: 20px 0;
      }

      .comments-section {
        padding: 20px;
        background: #f5f5f5;
        border-radius: 8px;
        margin-top: 20px;
      }
    `,
  ],
})
export class AppComponent {
  fileSource = 'assets/hello.docx'; // Change to your file path
  currentPageInfo: PageChangeEvent | null = null;

  toolbarOptions: ToolbarConfig = {
    showDownload: true,
    showPrint: true,
    showZoom: true,
    showRotation: true,
    showNavigation: true,
    showPageInput: true,
    showFitToWidth: true,
  };

  onFileLoaded(event: any) {
    console.log('File loaded:', event);
  }

  onPageChanged(event: PageChangeEvent) {
    this.currentPageInfo = event;
    console.log(
      `Page ${event.page} of ${event.totalPages} in ${event.type} document`
    );
    // Load comments for this page
    this.loadCommentsForPage(event.type, event.page);
  }

  loadCommentsForPage(fileType: string, pageNumber: number) {
    // Implement your comment loading logic here
    console.log(`Loading comments for ${fileType} page ${pageNumber}`);
  }
}
