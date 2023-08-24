import ExcelJS, { Worksheet, Row, Cell, CellValue } from "exceljs";
import { ExcelRow, ValidationError } from "../types/types";
import { count, error } from "console";
import { text } from "stream/consumers";

export class ExcelService {
    validateAndCopyData(inputSheet: Worksheet): ExcelJS.Workbook {
        const outputWorkbook = new ExcelJS.Workbook();
        const outputSheet = outputWorkbook.addWorksheet("Validated Data");

        const headerRow = outputSheet.addRow(inputSheet.getRow(1).values);
        this.applyHeaderStyles(headerRow);

        inputSheet.eachRow((inputRow: Row, rowNumber: number) => {
            if (rowNumber > 1) {
                const rowData: ExcelRow = {
                    id: inputRow.getCell(1).value,
                    name: inputRow.getCell(2).value,
                    email: inputRow.getCell(3).value,
                    mobileNumber: inputRow.getCell(4).value,
                    joiningDate: inputRow.getCell(5).value,
                };

                const validationErrors: ValidationError[] = this.validateRowData(rowData, rowNumber);
                const outputRow = outputSheet.addRow(inputRow.values);

                if (validationErrors.length > 0) {
                    this.applyErrorStyles(outputRow, validationErrors);
                }
            }
        });

        return outputWorkbook;
    }

    private applyHeaderStyles(row: Row): void {
        row.eachCell((cell, colNumber) => {
            if (cell.value) {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
            }
        });
    }

    private async applyErrorStyles(row: Row, errors: ValidationError[]) {
        const headerRow = row.worksheet.getRow(1); // Get the header row to determine the cell count

        for (let columnIndex = 1; columnIndex <= headerRow.cellCount; columnIndex++) {
            const cell = row.getCell(columnIndex);

            const cellAddress = cell.address;
            const errorMessages = errors
                .filter(error => error.cell === cellAddress)
                .map(error => error.errorMessage)
                .join("\n");

            if (errorMessages) {
                cell.fill = {
                    type: "pattern",
                    pattern: "solid",
                    fgColor: { argb: "FFFF0000" }, // Red background color
                };

                cell.note = {
                    texts: [{ text: errorMessages }]
                };
            }
        };
    }

    private validateRowData(rowData: ExcelRow, rowNumber: number): ValidationError[] {
        const validationErrors: ValidationError[] = [];

        const validateId = (id: any) => {
            if (!id) {
                validationErrors.push({ cell: `A${rowNumber}`, errorMessage: "ID is required" });
            }
            else if (typeof id !== 'number') {
                validationErrors.push({ cell: `A${rowNumber}`, errorMessage: "ID must be a number" });
            }
        };

        const validateName = (name: CellValue) => {
            if (!name) {
                validationErrors.push({ cell: `B${rowNumber}`, errorMessage: "Name is required" });
            }
            else if (!/^[a-zA-Z ]{1,50}$/.test(String(name))) {
                validationErrors.push({ cell: `B${rowNumber}`, errorMessage: "Name must contain only letters and have a maximum length of 50" });
            }
        };

        const validateEmail = (email: CellValue) => {
            if (!rowData.email) {
                validationErrors.push({ cell: `C${rowNumber}`, errorMessage: "Email is required" });
            }
            else if (!/^[\w\.-]+@[\w\.-]+\.\w+$/.test(String(email))) {
                validationErrors.push({ cell: `C${rowNumber}`, errorMessage: "Invalid email format" });
            }
        };

        const validateMobileNumber = (mobileNumber: CellValue) => {
            if (!mobileNumber) {
                validationErrors.push({ cell: `D${rowNumber}`, errorMessage: "Mobile number is required" });
            }
            else if (!/^\d{1,15}$/.test(String(mobileNumber))) {
                validationErrors.push({ cell: `D${rowNumber}`, errorMessage: "Mobile number must contain only digits and have a maximum length of 15" });
            }
        };

        const validateJoiningDate = (joiningDate: CellValue) => {
            if (!rowData.joiningDate) {
                validationErrors.push({ cell: `E${rowNumber}`, errorMessage: "Joining Date is required" });
            }
            else if (!/^(0[1-9]|1[0-2])\/(0[1-9]|[12]\d|3[01])\/\d{4}$/.test(
                new Date(String(joiningDate)).toLocaleDateString('en-US', {
                    day: '2-digit',
                    month: '2-digit',
                    year: 'numeric'
                }))) {
                validationErrors.push({ cell: `E${rowNumber}`, errorMessage: "Invalid date format. Use mm/dd/yyyy" });
            }
        };

        validateId(rowData.id);
        validateName(rowData.name);
        validateEmail(rowData.email);
        validateMobileNumber(rowData.mobileNumber);
        validateJoiningDate(rowData.joiningDate);
        return validationErrors;
    }
}
