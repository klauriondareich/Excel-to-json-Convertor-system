import { TestBed } from '@angular/core/testing';

import { ExceltojsonService } from './exceltojson.service';

describe('ExceltojsonService', () => {
  let service: ExceltojsonService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(ExceltojsonService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
