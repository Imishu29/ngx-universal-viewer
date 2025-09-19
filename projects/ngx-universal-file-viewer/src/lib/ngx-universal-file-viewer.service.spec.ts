import { TestBed } from '@angular/core/testing';

import { NgxUniversalFileViewerService } from './ngx-universal-file-viewer.service';

describe('NgxUniversalFileViewerService', () => {
  let service: NgxUniversalFileViewerService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(NgxUniversalFileViewerService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
