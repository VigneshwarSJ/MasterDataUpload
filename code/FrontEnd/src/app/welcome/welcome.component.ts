import { Component, OnInit } from '@angular/core';
import { Router } from '@angular/router';
import * as XLSX from 'xlsx';
import { HttpClient } from '@angular/common/http';

@Component({
  selector: 'app-welcome',
  templateUrl: './welcome.component.html',
  styleUrls: ['./welcome.component.css']
})
export class WelcomeComponent implements OnInit {
  username: string = 'admin'; // Default value
  selectedFile: File | null = null;
  sheetNames: string[] = [];
  selectedSheet: string = '';
  workbook: XLSX.WorkBook | null = null;
  displayData: any[][] = []; // Array to store Excel data for table display
  showDataModal: boolean = false; // Control visibility of the data modal
  validationResults: any[][] = [];
  showValidation = false;
  notification: string = '';
  notificationVisible: boolean = false;
  notificationTimeout: any;
  notificationType: 'success' | 'error' = 'success';
  notificationHover: boolean = false;
  overrideColumn: boolean = false;
  keyColumnIndex: number = 0;

  constructor(private router: Router, private http: HttpClient) {}

  ngOnInit(): void {
    // Get the username from localStorage
    const storedUsername = localStorage.getItem('username');
    if (storedUsername) {
      this.username = storedUsername;
    }
  }

  onFileChange(event: Event): void {
    const input = event.target as HTMLInputElement;
    if (input.files && input.files.length > 0) {
      this.selectedFile = input.files[0];
      this.readExcelFile(this.selectedFile);
    } else {
      this.selectedFile = null;
      this.sheetNames = [];
      this.selectedSheet = '';
    }
  }

  readExcelFile(file: File): void {
    const reader = new FileReader();
    reader.onload = (e: any) => {
      const data = new Uint8Array(e.target.result);
      this.workbook = XLSX.read(data, { type: 'array' });
      this.sheetNames = this.workbook.SheetNames;
      if (this.sheetNames.length > 0) {
        this.selectedSheet = this.sheetNames[0];
      }
    };
    reader.readAsArrayBuffer(file);
  }

  view(): void {
    if (!this.workbook || !this.selectedSheet) {
      alert('Please select an Excel file and a sheet first.');
      return;
    }

    // Get the worksheet data from the selected sheet
    const worksheet = this.workbook.Sheets[this.selectedSheet];
    
    // Convert sheet data to JSON with headers in the first row
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
    
    // Basic verification - check if sheet has data
    if (sheetData.length === 0) {
      alert('The selected sheet is empty. Please select a sheet with data.');
      return;
    }

    // Populate the displayData array for the table
    this.displayData = sheetData as any[][];
    
    // Make sure the first row exists for headers
    if (this.displayData.length > 0 && (!this.displayData[0] || this.displayData[0].length === 0)) {
      // If there's no header row or it's empty, create default column headers
      const columnCount = Math.max(...this.displayData.map(row => row ? row.length : 0));
      const defaultHeaders = Array(columnCount).fill(0).map((_, i) => `Column ${i + 1}`);
      this.displayData.unshift(defaultHeaders);
    }
    
    this.showValidation = false;
    this.showDataModal = true;
    
    console.log('Sheet data for display:', this.displayData);
  }
  
  closeDataModal(): void {
    this.showDataModal = false;
  }

  insert(): void {
    if (!this.workbook || !this.selectedSheet) {
      alert('Please select an Excel file and a sheet first.');
      return;
    }

    const worksheet = this.workbook.Sheets[this.selectedSheet];
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

    if (sheetData.length < 2) {
      alert('The selected sheet must have at least a header row and one data row.');
      return;
    }

    const columns = sheetData[0];
    const rows = sheetData.slice(1);

    // If overrideColumn is not checked, always use 0
    const keyColumnIndex = this.overrideColumn ? this.keyColumnIndex : 0;

    const payload = {
      sheetName: this.selectedSheet,
      columns: columns,
      rows: rows,
      keyColumnIndex: keyColumnIndex
    };

    this.http.post('/api/insert', payload).subscribe({
      next: (response: any) => {
        // Show notification like image2
        this.showNotification(response.message || 'Data inserted successfully!', 'success');
      },
      error: (error) => {
        // Show error notification in red
        this.showNotification('Insert failed: ' + (error.error?.message || error.message), 'error');
      }
    });
  }

  showNotification(message: string, type: 'success' | 'error' = 'success') {
    this.notification = message;
    this.notificationType = type;
    this.notificationVisible = true;
    this.notificationHover = false;
    if (this.notificationTimeout) {
      clearTimeout(this.notificationTimeout);
    }
    this.setNotificationTimeout();
  }

  setNotificationTimeout() {
    this.notificationTimeout = setTimeout(() => {
      if (!this.notificationHover) {
        this.notificationVisible = false;
      }
    }, 3000);
  }

  onNotificationMouseEnter() {
    this.notificationHover = true;
    if (this.notificationTimeout) {
      clearTimeout(this.notificationTimeout);
    }
  }

  onNotificationMouseLeave() {
    this.notificationHover = false;
    this.setNotificationTimeout();
  }

  logout(): void {
    // Clear the stored username
    localStorage.removeItem('username');
    // Navigate back to login page
    this.router.navigate(['/login']);
  }

  verifySheet(): void {
    if (!this.workbook || !this.selectedSheet) {
      alert('Please select an Excel file and a sheet first.');
      return;
    }

    const worksheet = this.workbook.Sheets[this.selectedSheet];
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
    if (sheetData.length < 2) {
      alert('The selected sheet must have at least a header row and one data row.');
      return;
    }

    const columns = sheetData[0];
    const rows = sheetData.slice(1);

    const payload = {
      sheetName: this.selectedSheet,
      columns: columns,
      rows: rows
    };

    this.http.post<any>('/api/verify-sheet', payload).subscribe({
      next: (response) => {
        this.validationResults = response.results;
        this.displayData = sheetData;
        this.showValidation = true;
        this.showDataModal = true;

        // Duplicate detection for any column containing 'name'
        columns.forEach((col, colIdx) => {
          if (col && col.toLowerCase().includes('name')) {
            const valueCounts: Record<string, number> = {};
            rows.forEach(row => {
              const value = row[colIdx];
              if (value !== undefined && value !== null && value !== '') {
                valueCounts[value] = (valueCounts[value] || 0) + 1;
              }
            });
            rows.forEach((row, rowIdx) => {
              const value = row[colIdx];
              if (value !== undefined && value !== null && value !== '' && valueCounts[value] > 1) {
                if (this.validationResults[rowIdx] && this.validationResults[rowIdx][colIdx]) {
                  this.validationResults[rowIdx][colIdx] = {
                    status: 'duplicate',
                    message: 'Duplicate value in name column'
                  };
                }
              }
            });
          }
        });
      },
      error: (error) => {
        alert('Verification failed: ' + (error.error?.message || error.message));
      }
    });
  }
}
