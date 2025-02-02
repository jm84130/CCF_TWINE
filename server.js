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

// Configuration GitHub
const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
const REPO_OWNER = "jm84130"; // Nom d'utilisateur GitHub
const REPO_NAME = "CCF_TWINE"; // Nom du dépôt GitHub
const EXCEL_PATH = "fichier_eleves/J1.xlsx"; // Chemin vers le fichier Excel
const DOCX_PATH = "fichier_eleves/J1.docx; // Chemin vers le fichier Word
const BRANCH = "main"; // Branche cible

axios.get('https://api.github.com/repos/jm84130/CCF_TWINE/contents/fichier_eleves/J1.xlsx', {
    headers: {
        Authorization: `Bearer ${githubToken}`,
    },
})
.then(response => {
    console.log('Fichier récupéré avec succès :', response.data);
})
.catch(error => {
    console.error('Erreur lors de la récupération du fichier :', error.response?.data || error.message);
});

// Chemin local du fichier Excel
const excelFilePath = path.join(__dirname, 'fichier_eleves', 'J1.xlsx');
// Chemin local du fichier Word
const excelFilePath = path.join(__dirname, 'fichier_eleves', 'J1.docx');

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

/// Route pour réinitialiser la feuille Excel
app.post('/reset-sheets', async (req, res) => {
    try {
        console.log("Réinitialisation de la feuille Excel en cours...");

        // Récupérer le SHA actuel du fichier
        const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;
        const getResponse = await axios.get(url, {
            headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` }
        });

        // Vérifiez si la réponse contient le champ `sha`
        if (!getResponse.data || !getResponse.data.sha) {
            throw new Error('Le fichier GitHub n\'a pas pu être récupéré ou le SHA est manquant.');
        }

        const fileSHA = getResponse.data.sha;

        // Contenu réinitialisé (fichier vide, exemple basique)
        const content = Buffer.from("Nom,Prenom,Score\n", 'utf-8').toString('base64');

        // Préparer les données pour mettre à jour le fichier
        const updateData = {
            message: "Réinitialisation des données",
            content: content,
            sha: fileSHA,
            branch: BRANCH
        };

        // Mettre à jour le fichier sur GitHub
        const updateResponse = await axios.put(url, updateData, {
            headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` }
        });

        // Vérifiez si la mise à jour a été réussie
        if (!updateResponse.data || updateResponse.data.message !== "Updated file") {
            throw new Error('La mise à jour du fichier sur GitHub a échoué.');
        }

        console.log('Réinitialisation réussie :', updateResponse.data);
        res.json({ success: true, message: 'Feuille réinitialisée avec succès.' });

    } catch (error) {
        console.error('Erreur lors de la réinitialisation :', error.response?.data || error.message);
        res.status(500).json({ success: false, message: `Erreur lors de la réinitialisation des feuilles. ${error.message}` });
    }
});


// Fonction pour enregistrer les données d'inscription dans le fichier Excel
async function saveIdentification(data) {
    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);

    const identificationSheet = workbook.getWorksheet('Identification');
    if (!identificationSheet) {
        throw new Error("Feuille 'Identification' introuvable dans le fichier Excel.");
    }

    // Vérifier les données d'entrée
    if (!data.session || !data.nom || !data.prenom || !data.date) {
        throw new Error("Toutes les informations d'inscription doivent être fournies.");
    }

    identificationSheet.getCell('B5').value = data.session;
    identificationSheet.getCell('B6').value = data.nom;
    identificationSheet.getCell('B7').value = data.prenom;
    identificationSheet.getCell('B8').value = data.date;

    await workbook.xlsx.writeFile(excelFilePath);
    console.log("Données d'inscription mises à jour localement.");
}

// Route pour recevoir les informations d'inscription
app.post('/update-identification', async (req, res) => {
    const data = req.body.identification;
    console.log("Requête reçue avec les données d'inscription :", data);

    try {
        await saveIdentification(data);

        // Mise à jour sur GitHub
        const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;
        const getResponse = await axios.get(url, {
            headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` }
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
            headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` }
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
    await workbook.xlsx.readFile(excelFilePath);

    const notationSheet = workbook.getWorksheet('Notation');
    if (!notationSheet) {
        throw new Error("Feuille 'Notation' introuvable dans le fichier Excel.");
    }

    // Dictionnaire des cellules valides par question
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

    // Vérification de la validité de la question
    if (!validCells) {
        throw new Error(`Question "${questionKey}" non reconnue.`);
    }

    // Réinitialisation des cellules
    validCells.forEach(cell => {
        notationSheet.getCell(cell).value = '';
    });

    // Mise à jour des cellules avec "X"
    if (Array.isArray(cellules) && cellules.length > 0) {
        cellules.forEach(cell => {
            if (validCells.includes(cell)) {
                notationSheet.getCell(cell).value = 'X';
            }
        });
    } else {
        throw new Error("Les cellules doivent être un tableau valide.");
    }

    // Sauvegarde dans le fichier Excel local
    await workbook.xlsx.writeFile(excelFilePath);
    console.log(`Cellules mises à jour pour ${questionKey}.`);
}

// Route pour mettre à jour les réponses aux questions
app.post('/update-notation', async (req, res) => {
    const { questionKey, cellules } = req.body;
    console.log("Requête reçue pour la question :", questionKey, "avec cellules :", cellules);

    try {
        await updateNotation(questionKey, cellules);

        // Mise à jour sur GitHub
        const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;
        const getResponse = await axios.get(url, {
            headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` }
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
            headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` }
        });

        console.log(`Réponses mises à jour pour ${questionKey} sur GitHub :`, updateResponse.data);
        res.json({ success: true, message: `Cellules ${cellules.join(', ')} mises à jour.` });
    } catch (error) {
        console.error("Erreur lors de la mise à jour des réponses :", error.response?.data || error.message);
        res.status(500).json({ success: false, message: "Erreur lors de la mise à jour des réponses." });
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

        // Récupérer le SHA actuel du fichier sur GitHub
        const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;
        const getResponse = await axios.get(url, {
            headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` }
        });

        const fileSHA = getResponse.data.sha;
        const fileContent = Buffer.from(getResponse.data.content, 'base64'); // Contenu base64 récupéré depuis GitHub
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(fileContent);

        // Sélectionner la feuille "Identification"
        const worksheet = workbook.getWorksheet('Identification');
        if (!worksheet) {
            return res.status(404).send("La feuille 'Identification' n'a pas été trouvée.");
        }

        // Mettre à jour les données dans les cellules spécifiques
        worksheet.getCell('B5').value = session;
        worksheet.getCell('B6').value = nom;
        worksheet.getCell('B7').value = prenom;
        worksheet.getCell('B8').value = date;

        // Convertir le fichier modifié en base64
        const updatedContent = await workbook.xlsx.writeBuffer();
        const base64UpdatedContent = updatedContent.toString('base64');

        // Mettre à jour le fichier Excel sur GitHub
        const updateData = {
            message: "Mise à jour des données d'identification",
            content: base64UpdatedContent,
            sha: fileSHA,
            branch: BRANCH
        };

        await axios.put(url, updateData, {
            headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` }
        });

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

        // Récupérer le SHA actuel du fichier Excel sur GitHub
        const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/${FILE_PATH}`;
        const getResponse = await axios.get(url, {
            headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` }
        });

        const fileSHA = getResponse.data.sha;
        const fileContent = Buffer.from(getResponse.data.content, 'base64'); // Contenu base64 récupéré depuis GitHub
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(fileContent);

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

        // Convertir le fichier modifié en base64
        const updatedContent = await workbook.xlsx.writeBuffer();
        const base64UpdatedContent = updatedContent.toString('base64');

        // Mettre à jour le fichier Excel sur GitHub
        const updateData = {
            message: `Mise à jour des réponses pour ${questionKey}`,
            content: base64UpdatedContent,
            sha: fileSHA,
            branch: BRANCH
        };

        await axios.put(url, updateData, {
            headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` }
        });

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

// Route pour ajouter des captures d'écran dans le fichier Word sur GitHub

app.post('/compile-screenshots', async (req, res) => {
    try {
        const { screenshots } = req.body; // Une liste de chemins de captures d'écran
        if (!screenshots || !Array.isArray(screenshots)) {
            return res.status(400).send("Le paramètre 'screenshots' est requis et doit être un tableau.");
        }

        console.log(`Compilation des captures d'écran dans le fichier Word sur GitHub...`);

        // Télécharger le fichier Word depuis GitHub
        const url = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/fichier_eleves/J1.docx`;
        const getResponse = await axios.get(url, {
            headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` }
        });

        const fileSHA = getResponse.data.sha;
        const fileContent = Buffer.from(getResponse.data.content, 'base64');

        // Charger le fichier Word
        const doc = await Packer.fromBuffer(fileContent);

        // Ajouter les captures d'écran au document Word
        screenshots.forEach((screenshotPath) => {
            doc.addSection({
                children: [
                    new Paragraph({
                        children: [
                            new Media({
                                path: screenshotPath, // Le chemin de la capture d'écran
                                width: 600, // Largeur de l'image
                                height: 400, // Hauteur de l'image
                            }),
                        ],
                    }),
                ],
            });
        });

        // Convertir le document Word en buffer et mettre à jour GitHub
        const updatedFileBuffer = await Packer.toBuffer(doc);
        const updatedBase64Content = updatedFileBuffer.toString('base64');

        // Préparer les données pour mettre à jour le fichier sur GitHub
        const updateData = {
            message: "Mise à jour du fichier Word avec des captures d'écran",
            content: updatedBase64Content,
            sha: fileSHA,
            branch: BRANCH
        };

        // Mettre à jour le fichier Word sur GitHub
        const updateUrl = `https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/contents/fichier_eleves/J1.docx`;
        const updateResponse = await axios.put(updateUrl, updateData, {
            headers: { Authorization: `Bearer ${process.env.GITHUB_TOKEN}` }
        });

        console.log("Le fichier Word a été mis à jour avec les captures d'écran sur GitHub.");
        return res.status(200).send({ message: "Compilation des captures d'écran réussie et fichier Word mis à jour sur GitHub." });
    } catch (error) {
        console.error("Erreur lors de la compilation des captures d'écran ou de la mise à jour du fichier Word :", error);
        return res.status(500).send("Une erreur est survenue lors de la compilation des captures d'écran ou de la mise à jour du fichier Word.");
    }
});

// Démarrage du serveur
const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
    console.log(`Serveur en cours d'exécution sur le port ${PORT}`);
});

