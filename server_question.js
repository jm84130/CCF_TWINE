const express = require('express');
const ExcelJS = require('exceljs');
const cors = require('cors');
const axios = require('axios'); // Pour télécharger le fichier Excel depuis GitHub

const app = express();
app.use(cors());
app.use(express.json());

// URL du fichier Excel sur GitHub
const excelFileURL = 'https://raw.githubusercontent.com/jm84130/CCF-twine/main/fichier_eleves/J1.xlsx';

// Mapping des questions et des cellules correspondantes
const cellMappings = {
    question1: ['E16', 'F16', 'G16', 'H16'],
    question2: ['E25', 'F25', 'G25', 'H25'],
    question3: ['E40', 'F40', 'G40', 'H40'],
    question4: ['E29', 'H29'],
    question5: ['E19', 'F19', 'G19', 'H19', 'E21', 'H21'],
    question6: ['E54', 'F54', 'G54', 'H54', 'E55', 'H55'],
    question7: ['E26', 'F26', 'G26', 'H26'],
    question8: ['E44', 'F44', 'G44', 'H44', 'E45', 'F45', 'G45', 'H45'],
    question9: ['E65', 'H65', 'E66', 'F66', 'G66', 'H66'],
    question10: ['E51', 'F51', 'G51', 'H51'],
    question11: ['E23', 'H23'],
};

// Fonction pour télécharger et charger le fichier Excel
async function fetchExcelWorkbook() {
    const response = await axios.get(excelFileURL, { responseType: 'arraybuffer' });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(response.data);
    return workbook;
}

// Fonction pour mettre à jour les cellules spécifiques
async function updateNotation(questionKey, cellules) {
    try {
        const workbook = await fetchExcelWorkbook();

        const notationSheet = workbook.getWorksheet('Notation');
        const validCells = cellMappings[questionKey];

        if (!validCells) {
            throw new Error(`Question "${questionKey}" non reconnue.`);
        }

        // Vérifie que toutes les cellules sont valides
        const invalidCells = cellules.filter(cell => !validCells.includes(cell));
        if (invalidCells.length > 0) {
            throw new Error(`Les cellules suivantes ne sont pas valides pour la question "${questionKey}": ${invalidCells.join(", ")}`);
        }

        // Efface les cellules de la question dans Excel
        validCells.forEach(col => {
            notationSheet.getCell(col).value = '';
        });

        // Met à jour les cellules spécifiées
        cellules.forEach(cell => {
            notationSheet.getCell(cell).value = 'X';
        });

        await workbook.xlsx.writeFile('fichier_eleves/J1.xlsx'); // Sauvegarde du fichier
        console.log(`Cellules mises à jour avec succès pour ${questionKey}: ${cellules.join(", ")}`);
        return { success: true, message: `Cellules ${cellules.join(", ")} mises à jour.` };
    } catch (error) {
        console.error("Erreur dans updateNotation :", error);
        return { success: false, message: "Erreur lors de la mise à jour de la feuille Notation." };
    }
}

// Route POST pour mettre à jour la notation
app.post('/update-notation', async (req, res) => {
    const { questionKey, cellules } = req.body;

    if (!questionKey || !cellules || !Array.isArray(cellules)) {
        res.status(400).json({ success: false, message: "Question ou cellules non spécifiées ou invalides." });
        return;
    }

    const result = await updateNotation(questionKey, cellules);
    if (result.success) {
        res.json(result);
    } else {
        res.status(500).json(result);
    }
});

// Utilisation du port attribué par Heroku ou 3003 en local
const PORT = process.env.PORT || 3003;
app.listen(PORT, () => {
    console.log(`Serveur de notation en écoute sur http://localhost:${PORT}`);
});
