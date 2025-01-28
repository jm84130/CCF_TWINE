<<<<<<< HEAD
// Import des modules nécessaires
require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const axios = require('axios');
const puppeteer = require('puppeteer'); // Pour la capture d'écran
const { Document: WordDocument, Packer, Paragraph, Media } = require('docx'); // Pour manipuler Word
const ExcelJS = require('exceljs'); // Pour manipuler Excel

// Initialisation de l'application Express
const app = express();

// Configuration GitHub
const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
if (!GITHUB_TOKEN) {
    throw new Error("Le token GitHub (GITHUB_TOKEN) n'est pas défini dans .env");
}

const GITHUB_API_BASE_URL = "https://api.github.com";
const REPO_OWNER = "jm84130";
const REPO_NAME = "CCF_TWINE";
const FILE_PATH = "fichier_eleves/J1.xlsx";
const BRANCH = "main"; 

app.get('/', (req, res) => {
  res.send('Bienvenue sur l\'application!');  // Réponse simple
});

// Route pour réinitialiser la feuille Excel
app.post('/reset-sheets', async (req, res) => {
    try {
        console.log("Réinitialisation de la feuille Excel en cours...");

        // Télécharger le fichier Excel à partir de GitHub
        const downloadUrl = `https://raw.githubusercontent.com/${REPO_OWNER}/${REPO_NAME}/${BRANCH}/${FILE_PATH}`;
        const response = await axios.get(downloadUrl, { responseType: 'arraybuffer' });

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(response.data);

        // Ici, vous pouvez réinitialiser la feuille ou les données nécessaires
        const worksheet = workbook.getWorksheet('Feuille1');
        if (!worksheet) {
            return res.status(404).send("La feuille spécifiée n'existe pas.");
        }

        // Exemple de réinitialisation : vider les données de la feuille
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell(cell => {
                cell.value = null;
            });
        });

        // Sauvegarder les changements dans le fichier Excel
        await workbook.xlsx.writeFile(path.join(__dirname, 'fichier_eleves/J1.xlsx'));

        console.log("Réinitialisation terminée.");

        return res.status(200).send("La feuille Excel a été réinitialisée avec succès.");
    } catch (error) {
        console.error("Erreur lors de la réinitialisation de la feuille Excel :", error);
        return res.status(500).send("Une erreur est survenue lors de la réinitialisation.");
    }
});

// Route pour l'enregistrement des informations d'identification
app.post('/update-identification', async (req, res) => {
    try {
        const { nom, prenom, date, session } = req.body;

        if (!nom || !prenom || !date || !session) {
            return res.status(400).send("Tous les champs sont requis.");
        }

        console.log(`Réception des données d'identification : ${nom} ${prenom}, ${date}, ${session}`);

        // Télécharger le fichier Excel depuis GitHub
        const downloadUrl = `https://raw.githubusercontent.com/${REPO_OWNER}/${REPO_NAME}/${BRANCH}/${FILE_PATH}`;
        const response = await axios.get(downloadUrl, { responseType: 'arraybuffer' });

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(response.data);

        // Sélectionner la feuille "Identification"
        const worksheet = workbook.getWorksheet('Identification');
        if (!worksheet) {
            return res.status(404).send("La feuille 'Identification' n'a pas été trouvée.");
        }

        // Mettre à jour les données dans les cellules spécifiques
        worksheet.getCell('B5').value = session;   // Session dans B5
        worksheet.getCell('B6').value = nom;       // Nom dans B6
        worksheet.getCell('B7').value = prenom;    // Prénom dans B7
        worksheet.getCell('B8').value = date;      // Date dans B8

        // Sauvegarder les modifications dans le fichier Excel
        await workbook.xlsx.writeFile(path.join(__dirname, 'fichier_eleves/J1.xlsx'));

        console.log("Identification enregistrée avec succès.");

        return res.status(200).send("Les données d'identification ont été enregistrées avec succès.");
    } catch (error) {
        console.error("Erreur lors de l'enregistrement de l'identification :", error);
        return res.status(500).send("Une erreur est survenue lors de l'enregistrement.");
    }
});

// Route pour la mise à jour des notations
app.post('/update-notation', async (req, res) => {
    try {
        const { questionKey, cellules } = req.body;

        if (!questionKey || !cellules || !Array.isArray(cellules)) {
            return res.status(400).send("Les paramètres 'questionKey' et 'cellules' sont requis.");
        }

        console.log(`Réception des données pour la question : ${questionKey}`);
        console.log(`Cellules à mettre à jour : ${cellules.join(", ")}`);

        // Télécharger le fichier Excel depuis GitHub
        const downloadUrl = `https://raw.githubusercontent.com/${REPO_OWNER}/${REPO_NAME}/${BRANCH}/${FILE_PATH}`;
        const response = await axios.get(downloadUrl, { responseType: 'arraybuffer' });

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(response.data);

        // Sélectionner la feuille "Notation"
        const worksheet = workbook.getWorksheet('Notation');
        if (!worksheet) {
            return res.status(404).send("La feuille 'Notation' n'a pas été trouvée.");
        }

        // Dictionnaire des cellules par question
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
        };

        const validCellsForQuestion = validCells[questionKey];
        if (!validCellsForQuestion) {
            return res.status(400).send(`La question ${questionKey} n'est pas reconnue.`);
        }

        // Réinitialiser les cellules avant de les remplir
        validCellsForQuestion.forEach(cell => {
            worksheet.getCell(cell).value = '';
        });

        // Mettre à jour les cellules avec "X"
        cellules.forEach(cell => {
            if (validCellsForQuestion.includes(cell)) {
                worksheet.getCell(cell).value = 'X';
            }
        });

        // Sauvegarder les modifications dans le fichier Excel
        await workbook.xlsx.writeFile(path.join(__dirname, 'fichier_eleves/J1.xlsx'));

        console.log(`Notation mise à jour pour ${questionKey}.`);

        return res.status(200).send("La notation a été mise à jour avec succès.");
    } catch (error) {
        console.error("Erreur lors de la mise à jour de la notation :", error);
        return res.status(500).send("Une erreur est survenue lors de la mise à jour de la notation.");
    }
});

// Route pour capturer une page web
app.post('/capture-screen', async (req, res) => {
    try {
        const { url, filename } = req.body;

        if (!url || !filename) {
            return res.status(400).send("Les paramètres 'url' et 'filename' sont requis.");
        }

        console.log(`Capture de l'écran pour l'URL : ${url} et sauvegarde sous le nom : ${filename}`);

        // Lancer Puppeteer
        const browser = await puppeteer.launch();
        const page = await browser.newPage();

        // Aller à l'URL fournie
        await page.goto(url);

        // Capturer une capture d'écran et la sauvegarder
        const filePath = path.join(__dirname, 'captures', filename);
        await page.screenshot({ path: filePath });

        console.log(`Capture d'écran sauvegardée sous ${filePath}`);

        // Fermer Puppeteer
        await browser.close();

        // Retourner la réponse avec le chemin du fichier
        return res.status(200).send({ message: "Capture d'écran réalisée avec succès", filePath: filePath });
    } catch (error) {
        console.error("Erreur lors de la capture de l'écran :", error);
        return res.status(500).send("Une erreur est survenue lors de la capture de l'écran.");
    }
});

// Route pour ajouter des captures d'écran dans un fichier Word
app.post('/compile-screenshots', async (req, res) => {
    try {
        const { screenshots } = req.body; // Une liste de chemins de captures d'écran
        if (!screenshots || !Array.isArray(screenshots)) {
            return res.status(400).send("Le paramètre 'screenshots' est requis et doit être un tableau.");
        }

        console.log(`Compilation des captures d'écran dans le fichier Word...`);

        // Créer un document Word
        const doc = new Document();

        // Ajouter les captures d'écran au document
        screenshots.forEach((screenshotPath) => {
            doc.addSection({
                children: [
                    new Paragraph({
                        children: [
                            new Media({
                                path: screenshotPath,
                                width: 600, // Largeur de l'image
                                height: 400, // Hauteur de l'image
                            }),
                        ],
                    }),
                ],
            });
        });

        // Sauvegarder le fichier Word
        const filePath = path.join(__dirname, 'fichiers', 'J1.docx');
        await Packer.toBuffer(doc).then((buffer) => {
            fs.writeFileSync(filePath, buffer);
        });

        console.log(`Le fichier Word a été généré : ${filePath}`);

        // Retourner la réponse avec le chemin du fichier généré
        return res.status(200).send({ message: "Compilation des captures d'écran réussie.", filePath: filePath });
    } catch (error) {
        console.error("Erreur lors de la compilation des captures d'écran :", error);
        return res.status(500).send("Une erreur est survenue lors de la compilation des captures d'écran.");
    }
});

// Démarrage du serveur
const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
    console.log(`Serveur en cours d'exécution sur le port ${PORT}`);
});

=======
const express = require('express');
const bodyParser = require('body-parser');
const cors = require('cors');

const app = express();

// Middleware
app.use(cors());
app.use(bodyParser.json());

// Route par défaut
app.get('/', (req, res) => {
    res.send('Serveur Node.js opérationnel !');
});

// Démarrage du serveur
const PORT = process.env.PORT || 3000; // Utilisation du port Heroku ou 3000 en local
app.listen(PORT, () => {
    console.log(`Serveur en cours d'exécution sur http://localhost:${PORT}`);
});
>>>>>>> 74a3c44 (Initial commit pour déploiement Heroku)
