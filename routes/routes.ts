import express from "express";
import { uploadAndValidateExcel } from "../controllers/excelController";
import multer from "multer";

const router = express.Router();

// Configure multer for handling file uploads
const storage = multer.memoryStorage();
const upload = multer({ storage });


router.post("/upload", upload.single("file"), uploadAndValidateExcel);

export default router;
