import { ComponentFixture, TestBed } from '@angular/core/testing';

import { NgxUniversalFileViewerComponent } from './ngx-universal-file-viewer.component';

describe('NgxUniversalFileViewerComponent', () => {
  let component: NgxUniversalFileViewerComponent;
  let fixture: ComponentFixture<NgxUniversalFileViewerComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [NgxUniversalFileViewerComponent]
    })
    .compileComponents();
    
    fixture = TestBed.createComponent(NgxUniversalFileViewerComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
