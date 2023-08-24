import { Request, Response } from "express";
import path from "path";
import { ExcelService } from "../services/excelService";
import exceljs from "exceljs";

const excelService = new ExcelService();

export const uploadAndValidateExcel = async (req: Request, res: Response) => {
    try {
        if (!req.file) {
            return res.status(400).json({ message: "No file uploaded" });
        }

        // Load the file 
        const workbook = new exceljs.Workbook();
        await workbook.xlsx.load(req.file.buffer);

        const inputSheet = workbook.getWorksheet(1);

        // Create a separate sheet
        const outputWorkbook = excelService.validateAndCopyData(inputSheet);

        const outputFilePath = path.join(__dirname, "../", "output.xlsx");
        
        await outputWorkbook.xlsx.writeFile(outputFilePath);

        return res.status(200).download(outputFilePath, "output.xlsx");
    } catch (error) {
        console.error("Error:", error);
        return res.status(500).json({ message: "An error occurred" });
    }
};
