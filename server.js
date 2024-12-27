require('dotenv').config();

const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const axios = require('axios');

// Configuration GitHub
const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
if (!GITHUB_TOKEN) {
    throw new Error("Le token GitHub (GITHUB_TOKEN) n'est pas défini dans .env");
}
const GITHUB_API_BASE_URL = "https://api.github.com"; // Base URL de l'API GitHub
const REPO_OWNER = "jm84130"; // Nom d'utilisateur GitHub
const REPO_NAME = "CCF_TWINE"; // Nom du dépôt GitHub
const FILE_PATH = "fichier_eleves/J1.xlsx"; // Chemin du fichier dans le dépôt
const BRANCH = "main"; // Branche cible

// Chemin local du fichier Excel
const excelFilePath = path.join(__dirname, 'fichier_eleves', 'J1.xlsx');

// Initialisation de l'application Express
const app = express();
app.use(cors());
app.use(bodyParser.json());

// Route pour réinitialiser les feuilles Excel sur GitHub
app.post('/reset-sheets', async (req, res) => {
    try {
        // Construire l'URL pour récupérer les informations du fichier
        const url = `${GITHUB_API_BASE_URL}/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;

        // Étape 1 : Récupérer le SHA actuel du fichier
        const getResponse = await axios.get(url, {
            headers: { Authorization: `Bearer ${GITHUB_TOKEN}` }
        });

        const fileSHA = getResponse.data.sha;

        // Étape 2 : Contenu réinitialisé (fichier vide, exemple basique)
        const content = Buffer.from("Nom,Prenom,Score\n", 'utf-8').toString('base64');

        // Étape 3 : Préparer les données pour mettre à jour le fichier
        const updateData = {
            message: "Réinitialisation des données",
            content: content,
            sha: fileSHA,
            branch: BRANCH
        };

        // Étape 4 : Mettre à jour le fichier sur GitHub
        const updateResponse = await axios.put(url, updateData, {
            headers: { Authorization: `Bearer ${GITHUB_TOKEN}` }
        });

        console.log('Réinitialisation réussie :', updateResponse.data);
        res.json({ success: true, message: 'Feuille réinitialisée avec succès.' });
    } catch (error) {
        console.error('Erreur lors de la réinitialisation :', error.response?.data || error.message);
        res.status(500).json({ success: false, message: 'Erreur lors de la réinitialisation des feuilles.' });
    }
});

// Fonction pour enregistrer les données d'inscription dans le fichier Excel
async function saveIdentification(data) {
    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();

    // Charger le fichier Excel
    await workbook.xlsx.readFile(excelFilePath);

    // Sélectionner la feuille "Identification"
    const identificationSheet = workbook.getWorksheet('Identification');
    if (!identificationSheet) {
        throw new Error("Feuille 'Identification' introuvable dans le fichier Excel.");
    }

    // Remplir les cellules avec les données reçues
    identificationSheet.getCell('B5').value = data.session;
    identificationSheet.getCell('B6').value = data.nom;
    identificationSheet.getCell('B7').value = data.prenom;
    identificationSheet.getCell('B8').value = data.date;

    // Sauvegarder les modifications localement
    await workbook.xlsx.writeFile(excelFilePath);
    console.log("Données d'inscription mises à jour localement.");
}

// Route pour recevoir les informations d'inscription
app.post('/update-identification', async (req, res) => {
    const data = req.body.identification;
    console.log("Requête reçue avec les données d'inscription :", data);

    try {
        // Étape 1 : Mise à jour locale du fichier Excel
        await saveIdentification(data);

        // Étape 2 : Mise à jour du fichier sur GitHub
        const url = `${GITHUB_API_BASE_URL}/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;
        const getResponse = await axios.get(url, {
            headers: { Authorization: `Bearer ${GITHUB_TOKEN}` }
        });

        const fileSHA = getResponse.data.sha;
        const fileContent = fs.readFileSync(excelFilePath);
        const base64Content = fileContent.toString("base64");

        const updateData = {
            message: "Mise à jour des données d'inscription",
            content: base64Content,
            sha: fileSHA,
            branch: BRANCH
        };

        const updateResponse = await axios.put(url, updateData, {
            headers: { Authorization: `Bearer ${GITHUB_TOKEN}` }
        });

        console.log("Données d'inscription mises à jour sur GitHub :", updateResponse.data);
        res.json({ success: true, message: "Données d'inscription mises à jour avec succès !" });
    } catch (error) {
        console.error("Erreur lors de la mise à jour des données d'inscription :", error.response?.data || error.message);
        res.status(500).json({ success: false, message: "Erreur lors de la mise à jour des données d'inscription." });
    }
});

// Fonction pour mettre à jour les réponses dans le fichier Excel
async function updateNotation(questionKey, cellules) {
    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();

    // Charger le fichier Excel
    await workbook.xlsx.readFile(excelFilePath);

    // Sélectionner la feuille "Notation"
    const notationSheet = workbook.getWorksheet('Notation');
    if (!notationSheet) {
        throw new Error("Feuille 'Notation' introuvable dans le fichier Excel.");
    }

    // Effacer les cellules avant de les mettre à jour
    const validCells = {
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
    }[questionKey];

    if (!validCells) {
        throw new Error(`Question "${questionKey}" non reconnue.`);
    }

    // Réinitialiser les cellules
    validCells.forEach(cell => {
        notationSheet.getCell(cell).value = '';
    });

    // Mettre à jour les cellules avec "X"
    cellules.forEach(cell => {
        notationSheet.getCell(cell).value = 'X';
    });

    // Sauvegarder les modifications localement
    await workbook.xlsx.writeFile(excelFilePath);
    console.log(`Cellules mises à jour pour ${questionKey}.`);
}

// Route pour mettre à jour les réponses aux questions
app.post('/update-notation', async (req, res) => {
    const { questionKey, cellules } = req.body;
    console.log("Requête reçue pour la question :", questionKey, "avec cellules :", cellules);

    try {
        // Étape 1 : Mise à jour locale du fichier Excel
        await updateNotation(questionKey, cellules);

        // Étape 2 : Mise à jour du fichier sur GitHub
        const url = `${GITHUB_API_BASE_URL}/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;
        const getResponse = await axios.get(url, {
            headers: { Authorization: `Bearer ${GITHUB_TOKEN}` }
        });

        const fileSHA = getResponse.data.sha;
        const fileContent = fs.readFileSync(excelFilePath);
        const base64Content = fileContent.toString("base64");

        const updateData = {
            message: `Mise à jour des réponses pour ${questionKey}`,
            content: base64Content,
            sha: fileSHA,
            branch: BRANCH
        };

        const updateResponse = await axios.put(url, updateData, {
            headers: { Authorization: `Bearer ${GITHUB_TOKEN}` }
        });

        console.log(`Réponses mises à jour pour ${questionKey} sur GitHub :`, updateResponse.data);
        res.json({ success: true, message: `Cellules ${cellules.join(', ')} mises à jour.` });
    } catch (error) {
        console.error("Erreur lors de la mise à jour des réponses :", error.response?.data || error.message);
        res.status(500).json({ success: false, message: "Erreur lors de la mise à jour des réponses." });
    }
});

app.get('/', (req, res) => {
    res.send('Bienvenue sur le serveur de l\'application Twine !');
});

// Gestion globale des erreurs non capturées
process.on('uncaughtException', (err) => {
    console.error('Une erreur non capturée est survenue :', err.message);
    console.error(err.stack);
});

process.on('unhandledRejection', (reason, promise) => {
    console.error('Rejet non géré détecté dans une promesse :', promise);
    console.error('Raison :', reason);
});

// Port d'écoute du serveur
const PORT = process.env.PORT || 3000;

// Lancement du serveur
app.listen(PORT, () => {
    console.log(`Serveur en cours d'exécution sur http://localhost:${PORT}`);
});
