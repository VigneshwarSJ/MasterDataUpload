import { Component, Input, Output, EventEmitter } from '@angular/core';

@Component({
  selector: 'app-data-modal',
  templateUrl: './data-modal.component.html',
  styleUrls: ['./data-modal.component.css']
})
export class DataModalComponent {
  @Input() isVisible: boolean = false;
  @Input() data: any[][] = [];
  @Input() title: string = 'Excel Data';
  @Input() validationResults: any[][] = [];
  @Input() showValidation: boolean = false;
  @Output() closed = new EventEmitter<void>();
  @Output() verify = new EventEmitter<void>();

  close(): void {
    this.isVisible = false;
    this.closed.emit();
  }

  onVerify(): void {
    this.verify.emit();
  }

  getCellClass(rowIdx: number, colIdx: number): string {
    if (!this.showValidation || !this.validationResults || !this.validationResults[rowIdx]) return '';
    const status = this.validationResults[rowIdx][colIdx]?.status;
    if (status === 'valid') return 'cell-valid';
    if (status === 'length') return 'cell-length';
    if (status === 'type') return 'cell-type';
    if (status === 'duplicate') return 'cell-duplicate';
    if (status === 'namecolumn') return 'cell-namecolumn';
    return '';
  }

  getCellMessage(rowIdx: number, colIdx: number): string {
    if (!this.showValidation || !this.validationResults || !this.validationResults[rowIdx]) return '';
    return this.validationResults[rowIdx][colIdx]?.message || '';
  }
}
