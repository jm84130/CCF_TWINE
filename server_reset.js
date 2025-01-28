const express = require('express');
const ExcelJS = require('exceljs');
const axios = require('axios'); // Pour télécharger le fichier Excel depuis GitHub
const cors = require('cors');

const app = express();
app.use(cors()); // Activation de CORS pour toutes les requêtes

// Port dynamique ou 3002 par défaut
const port = process.env.PORT || 3002;

// URL du fichier Excel sur GitHub
const excelFileURL = 'https://raw.githubusercontent.com/jm84130/CCF-twine/main/fichier_eleves/J1.xlsx';

// Fonction pour télécharger et charger le fichier Excel depuis GitHub
async function fetchExcelWorkbook() {
    try {
        const response = await axios.get(excelFileURL, { responseType: 'arraybuffer' });
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(response.data); // Chargement du fichier Excel
        return workbook;
    } catch (error) {
        console.error("Erreur lors du téléchargement ou du chargement du fichier Excel :", error.message);
        throw new Error("Impossible de charger le fichier Excel depuis GitHub.");
    }
}

// Fonction pour réinitialiser les cellules de la feuille Notation
async function resetNotationSheet() {
    try {
        // Charger le fichier Excel depuis GitHub
        const workbook = await fetchExcelWorkbook();

        const notationSheet = workbook.getWorksheet('Notation');
        if (!notationSheet) {
            throw new Error("Feuille 'Notation' introuvable dans le fichier Excel.");
        }

        // Liste des cellules des colonnes E, F, G et H à réinitialiser
        const rowsToClear = [16, 19, 21, 22, 23, 25, 26, 27, 29, 38, 40, 44, 45, 51, 54, 55, 65, 66, 70];
        rowsToClear.forEach(row => {
            notationSheet.getCell(`E${row}`).value = '';
            notationSheet.getCell(`F${row}`).value = '';
            notationSheet.getCell(`G${row}`).value = '';
            notationSheet.getCell(`H${row}`).value = '';
        });

        console.log("Réinitialisation de la feuille Notation réussie.");
    } catch (error) {
        console.error("Erreur dans resetNotationSheet :", error.message);
        throw error;
    }
}

// Route POST pour déclencher la réinitialisation de la feuille Notation
app.post('/reset-sheets', async (req, res) => {
    try {
        await resetNotationSheet();
        res.json({ success: true, message: "Réinitialisation réussie !" });
    } catch (error) {
        console.error("Erreur lors de la réinitialisation :", error.message);
        res.status(500).json({ success: false, message: "Erreur lors de la réinitialisation de la feuille Notation." });
    }
});

// Démarrage du serveur
app.listen(port, () => {
    console.log(`Serveur de réinitialisation en écoute sur le port ${port}`);
});
