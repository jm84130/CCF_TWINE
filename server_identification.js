const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const ExcelJS = require('exceljs');
const axios = require('axios'); // Pour télécharger le fichier Excel depuis GitHub

const app = express();
app.use(cors());
app.use(bodyParser.json());

// URL du fichier Excel sur GitHub
const excelFileURL = 'https://raw.githubusercontent.com/jm84130/CCF-twine/main/fichier_eleves/J1.xlsx';

// Fonction pour télécharger et charger le fichier Excel
async function fetchExcelWorkbook() {
    const response = await axios.get(excelFileURL, { responseType: 'arraybuffer' });
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(response.data);
    return workbook;
}

// Fonction pour enregistrer les informations d'inscription dans la feuille "Identification"
async function saveIdentification(data) {
    const workbook = await fetchExcelWorkbook();

    const identificationSheet = workbook.getWorksheet('Identification');
    if (!identificationSheet) {
        throw new Error("Feuille Identification introuvable dans le fichier Excel.");
    }

    identificationSheet.getCell('B5').value = data.session;
    identificationSheet.getCell('B6').value = data.nom;
    identificationSheet.getCell('B7').value = data.prenom;
    identificationSheet.getCell('B8').value = data.date;

    console.log("Données d'inscription mises à jour localement dans le fichier téléchargé.");
}

app.post('/update-identification', async (req, res) => {
    const data = req.body.identification;

    try {
        await saveIdentification(data);
        res.json({ success: true, message: "Données d'inscription mises à jour avec succès !" });
    } catch (error) {
        console.error("Erreur lors de l'enregistrement dans la feuille Identification :", error);
        res.status(500).json({ success: false, message: "Erreur lors de l'enregistrement des données d'inscription." });
    }
});

// Utilisation du port attribué par Heroku ou 3001 en local
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
    console.log(`Serveur Identification en cours d'exécution sur http://localhost:${PORT}`);
});
