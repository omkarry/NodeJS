import express from "express";
import excelRoutes from "./routes/routes";

const app = express();

app.use(express.json());

app.use("/api/excel", excelRoutes);

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
