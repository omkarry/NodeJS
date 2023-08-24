import { CellValue } from 'exceljs';
export interface ExcelRow {
    id: CellValue;
    name: CellValue;
    email: CellValue;
    mobileNumber: CellValue;
    joiningDate: CellValue;
}

export interface ValidationError {
    cell: string;
    errorMessage: string;
}
