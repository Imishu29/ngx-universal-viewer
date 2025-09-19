import { NgModule } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { NgxUniversalFileViewerComponent } from './ngx-universal-file-viewer.component';

@NgModule({
  imports: [NgxUniversalFileViewerComponent, FormsModule],
  exports: [NgxUniversalFileViewerComponent]
})
export class NgxUniversalFileViewerModule { }