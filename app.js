// Variables globales
let currentWorkbook = null;
let currentFileName = '';
let personInfo = {
    nom: '',
    prenom: '',
    dateNaissance: '',
    moniteurs: ''
};
let allSheetNames = [];
let currentSheetData = null;
let templateQuestions = [];

// Questionnaire de base (sauvegardé en dur)
const baseQuestionnaire = [
    {
        question: "Situation de famille",
        answers: ["Célibataire", "Marié(e)/Couple", "Autre", "Enfant(s) (Préciser nombre)"],
        observation: ""
    },
    {
        question: "Hébergement (Préciser adresse)",
        answers: ["Domicile", "Famille", "Famille d'acceuil", "Foyer"],
        observation: ""
    },
    {
        question: "Mesure de protection (Préciser organisme)",
        answers: ["Tutelle", "Curatelle simple", "Curatelle renforcée", "Sauvegarde de justice", "Habilitation familiale", "Aucune mesure de protection juridique", "Autre (Préciser)"],
        observation: ""
    },
    {
        question: "Objet du Bilan/PPI (Projet Personnalisé Individuel)",
        answers: ["Bilan/PPI Annuel", "Bilan de Stage (Préciser date)"],
        observation: ""
    },
    {
        question: "Temps de travail (ETP)",
        answers: ["Temps plein", "4/5 jours", "3/5 jours", "Mi-temps", "2/5 jours", "1/5 jours", "Autre (Précier)"],
        observation: ""
    },
    {
        question: "Poste(s) Occupé(s) (Depuis le dernier Bilan/Projet Personnalisé Individuel)",
        answers: [],
        observation: ""
    },
    {
        question: "Tâches réalisées (Depuis le dernier Bilan/Projet Personnalisé Individuel)",
        answers: [],
        observation: ""
    },
    {
        question: "Formation(s) suivie(s) (Depuis le dernier Bilan/Projet Personnalisé Individuel)",
        answers: [],
        observation: ""
    },
    {
        question: "Activité de la personne à l'ESAT (Soutien professionnel, extra-profesionnel à l'ESAT)",
        answers: [],
        observation: ""
    },
    {
        question: "Activité de la personne à l'exterieur (activité(s) et loisir(s) culturel(s) et sportif(s))",
        answers: [],
        observation: ""
    },
    {
        question: "Accompagnement psychologique régulier",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Accompagnement psychologique sur demande",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Accompagnement régulier par l'infirmière de l'ESAT",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Accompagnement à l'exterieur de l'ESAT (Si oui, préciser)",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Évolution professionnelle",
        answers: ["En progression", "stabilité", "En régression", "Irrégulière"],
        observation: ""
    },
    {
        question: "Évolution au niveau social",
        answers: ["En progression", "stabilité", "En régression", "Irrégulière"],
        observation: ""
    },
    {
        question: "Sait lire",
        answers: ["Oui", "Oui, avec difficulté", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Sait écrire ",
        answers: ["Oui", "Oui, avec difficulté", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Sait compter",
        answers: ["Oui, plus de 50", "Jusqu'à 20", "Jusqu'à 10", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Sait calculer",
        answers: ["Addition", "Soustraction", "Multiplication", "Division", "Non évalué"],
        observation: ""
    },
    {
        question: "S’orienter dans le temps",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "S’orienter dans l’espace",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Comprendre une consigne, (une phrase simple)",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Comprendre plusieurs consignes, (une phrase complexe)",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Mémoriser ses acquis",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Capable de corriger ses erreurs",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Ponctuel",
        answers: ["Oui", "Non"],
        observation: ""
    },
    {
        question: "Assidue au travail",
        answers: ["Oui", "Non"],
        observation: ""
    },
    {
        question: "Reste a son poste de travail",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Motivé pour accomplir son travail",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Exécuter des opérations variées",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "S'adapter à des postes ou des tâches variées",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Gerer son travail de façon autonome",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Attentif à la sécurité",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Attentif à la qualité du travail réalisé",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Attentif à la quantité",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Concentré lors de la réalisation d'une tâche",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Capable d'initiative(s) adaptée(s) dans une situation connue",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Capable d'initiative(s) adaptée(s) dans une situation nouvelle",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Sait controler le résultat de son travail et sait reprendre un travail mal exécuté",
        answers: ["Oui", "Non", "Oui, reprend un travail mal exécuté", "Non, ne reprend pas un travail mal exécuté", "Non évalué"],
        observation: ""
    },
    {
        question: "Capable de gestes précis et coordonés",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Rythme de travail",
        answers: ["Trèp rapide", "Rapide", "Adapté", "Lent", "Très lent", "irrégulier", "Inadapté", "Non évalué"],
        observation: ""
    },
    {
        question: "Capable de suivre une cadence imposée (Voire temporairement soutenue)",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Accèpte de changer de poste",
        answers: ["Facilement", "Difficilement", "Refuse", "Non évalué"],
        observation: ""
    },
    {
        question: "Sait organiser efficacement son poste de travail",
        answers: ["En autonomie", "Avec de l'aide", "Ne sais pas", "Non évalué"],
        observation: ""
    },
    {
        question: "Entretient l'ordre et la propreté de son poste de travail et de ses outils",
        answers: ["En autonomie", "Avec de l'aide", "Ne sais pas", "Non évalué"],
        observation: ""
    },
    {
        question: "Fait preuve d'une bonne résistance physique",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "A le sens des responsabilité (agir, décider par/pour sois même et le groupe, accepte la responsabilité de ses erreurs)",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Sait communiquer",
        answers: [" Par la parole", " Communication non verbale", "Pas de communication", "Aide technique (Préciser)", "Non évalué"],
        observation: ""
    },
    {
        question: "Utiliser le téléphone",
        answers: ["En autonomie (appel/message/mail/gps/application divers)", "En usage limité (appel/message/mail)", "Avec de l'aide", "Ne sait pas", "Non évalué"],
        observation: ""
    },
    {
        question: "Capacité à être dans un groupe restreint",
        answers: ["Oui", "Oui, avec difficulté", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Capacité à être dans un grand groupe",
        answers: ["Oui", "Oui, avec difficulté", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Relations avec les professionnels",
        answers: ["Communique de lui même facilement", "Communique, quand il est solicité", "Ne communique pas, même quand solicité", "Communication conflictuel", "Non évalué"],
        observation: ""
    },
    {
        question: "Relations avec ses pairs",
        answers: ["Communique de lui même facilement", "Communique, quand il est solicité", "Ne communique pas, même quand solicité", "Communication conflictuel", "Non évalué"],
        observation: ""
    },
    {
        question: "Relations avec l’entourage familial ou amical",
        answers: ["Communique de lui même facilement", "Communique, quand il est solicité", "Ne communique pas, même quand solicité", "Communication conflictuel", "Non évalué"],
        observation: ""
    },
    {
        question: "Respectueux des règles (règlements, sécurité...)",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Se met, par son comportement en danger pour gérer sa sécurité",
        answers: ["Souvent", "Parfois", "Jamais", "Non évalué"],
        observation: ""
    },
    {
        question: "Respectueux des autres (civisme, politesse, amabilité...)",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Maîtrise son comportement dans ses relations avec autrui",
        answers: ["Souvent", "Parfois", "Jamais", "Non évalué"],
        observation: ""
    },
    {
        question: "Comportement adapté (aux situations rencontrées)",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Tenue et Hygiène adaptées au travail",
        answers: ["Tenue et hygiène correctes, adaptées au travail", "Tenue correcte, mais non adaptée au travail", "Tenue incorrecte", "Hygiène correcte", "Hygiène incorrecte", "Non évalué"],
        observation: ""
    },
    {
        question: "Aptitude pour gérer ses rendez-vous extérieur (médicaux, activités, …)",
        answers: ["En autonomie", "Avec de l'aide", "Ne sais pas", "Non évalué"],
        observation: ""
    },
    {
        question: "Capable de violence (physique ou verbale)",
        answers: ["Oui", "Non", "Non évalué"],
        observation: ""
    },
    {
        question: "Autonomie dans les déplacements sur site (à l'interieur du site)",
        answers: ["En autonomie", "Avec de l'aide", "Ne sais pas", "Non évalué"],
        observation: ""
    },
    {
        question: "Autonomie dans les déplacements hors site (à l'exterieur du site)",
        answers: ["En autonomie", "Avec de l'aide", "Ne sais pas", "Non évalué"],
        observation: ""
    },
    {
        question: "Utiliser les transports en communs",
        answers: ["En autonomie", "Avec de l'aide", "Ne sais pas", "Non évalué"],
        observation: ""
    },
    {
        question: "Conduire un véhicule",
        answers: ["En autonomie", "Avec de l'aide", "Ne sais pas", "Non évalué"],
        observation: ""
    },
    {
        question: "Tâche(s) préférée(s) à l'ESAT (Préciser)",
        answers: [],
        observation: ""
    },
    {
        question: "Tâche(s) moins appréciée(s) à l'ESAT (Préciser)",
        answers: [],
        observation: ""
    },
    {
        question: "Evenement marquant ou faits significatifs (Positif ou négatif)",
        answers: [],
        observation: ""
    },
    {
        question: "Ressenti et note sur 10, avec commentaire",
        answers: ["Positif", "Neutre", "Négatif"],
        observation: ""
    },
    {
        question: "Souhait de formation",
        answers: [],
        observation: ""
    },
    {
        question: "Besoin en formation pour l'activité (Repérées par le MA)",
        answers: [],
        observation: ""
    },
    {
        question: "Souhaite Présenter une R.A.E (Si oui, préciser référentiel métier)",
        answers: ["Oui", "Non", "En réflexion"],
        observation: ""
    },
    {
        question: "Comptences à acquérir en lien avec le tableau des compétences",
        answers: [],
        observation: ""
    },
{
        question: "Métier rêvé/préféré (Préciser)",
        answers: [],
        observation: ""
    },
    {
        question: "Envie de ... (Préciser)",
        answers: ["DuoDay", "Stage dans un autre atelier", "Stage dans un autre ESAT", "Stage en milieu ordinaire", "quitter l'ESAT"],
        observation: ""
    },
    {
        question: "Autres besoins repérés ... (Préciser)",
        answers: ["Temps partiel", "Retraite", "Réorientation", "Soins", "Accompagnement", "Autre"],
        observation: ""
    },
    {
        question: "Le Projet Personnalisé précedent à t'il été réalisé (Préciser)",
        answers: ["Dans son ensemble", "Partiellement", "Non"],
        observation: ""
    },
    {
        question: "Nouveau Projet Personnalisé",
        answers: [],
        observation: ""
    },
    {
        question: "Point de vue de l'usager",
        answers: [],
        observation: ""
    },
    {
        question: "Point de vue des professionnels",
        answers: [],
        observation: ""
    }
];

// Initialisation
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

function initializeApp() {
    // Boutons page d'accueil
    document.getElementById('searchPersonBtn').addEventListener('click', () => {
        document.getElementById('fileInput').click();
    });
    
    document.getElementById('fileInput').addEventListener('change', handleFileSelect);
    
    document.getElementById('guideBtn').addEventListener('click', openGuideModal);
    
    // Boutons page principale
    document.getElementById('consultBtn').addEventListener('click', showConsultPage);
    document.getElementById('newEvalBtn').addEventListener('click', showNewEvalPage);
    document.getElementById('backToHomeBtn').addEventListener('click', () => showPage('homePage'));
    
    // Boutons page consultation
    document.getElementById('backFromConsultBtn').addEventListener('click', () => showPage('mainPage'));
    
    // Boutons page nouvelle évaluation
    document.getElementById('backFromNewEvalBtn').addEventListener('click', () => showPage('mainPage'));
    document.getElementById('saveXlsxBtn').addEventListener('click', () => saveEvaluation('xlsx'));
    document.getElementById('saveOdsBtn').addEventListener('click', () => saveEvaluation('ods'));
    
    // Modal
    document.querySelector('.close').addEventListener('click', closeGuideModal);
    document.getElementById('exportBaseBtn').addEventListener('click', () => exportBaseQuestionnaire('xlsx'));
    document.getElementById('exportBaseOdsBtn').addEventListener('click', () => exportBaseQuestionnaire('ods'));
    
    // Fermer modal en cliquant à l'extérieur
    window.addEventListener('click', (e) => {
        const modal = document.getElementById('guideModal');
        if (e.target === modal) {
            closeGuideModal();
        }
    });
    
    // Événements pour les champs d'informations personnelles
    setupPersonInfoFields();
}

// Événements pour les champs d'informations personnelles
setupPersonInfoFields();

function setupPersonInfoFields() {
    // Champs sur la page principale
    const nomField = document.getElementById('personNom');
    const prenomField = document.getElementById('personPrenom');
    const dateField = document.getElementById('personDateNaissance');
    const moniteursField = document.getElementById('personMoniteurs');
    
    if (nomField) {
        nomField.addEventListener('input', (e) => {
            personInfo.nom = e.target.value.trim().toUpperCase();
            updatePersonInfoDisplay();
        });
    }
    
    if (prenomField) {
        prenomField.addEventListener('input', (e) => {
            // Capitaliser chaque partie séparée par un tiret
            const parts = e.target.value.trim().split('-');
            personInfo.prenom = parts.map(part => 
                part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()
            ).join('-');
            updatePersonInfoDisplay();
        });
    }
    
    if (dateField) {
    dateField.addEventListener('input', function(e) {
        let value = e.target.value.replace(/[^0-9]/g, ''); // Garder uniquement les chiffres
        
        // Limiter à 8 chiffres
        if (value.length > 8) {
            value = value.substring(0, 8);
        }
        
        // Ajouter automatiquement les "/"
        let formatted = '';
        if (value.length > 0) {
            formatted = value.substring(0, 2);
            if (value.length >= 3) {
                formatted += '/' + value.substring(2, 4);
            }
            if (value.length >= 5) {
                formatted += '/' + value.substring(4, 8);
            }
        }
        
        // Mettre à jour le champ
        e.target.value = formatted;
        
        // Sauvegarder si la date est complète (10 caractères = JJ/MM/AAAA)
        if (formatted.length === 10) {
            personInfo.dateNaissance = formatted;
            updatePersonInfoDisplay();
        }
    });
}
    
    if (moniteursField) {
        moniteursField.addEventListener('input', (e) => {
            personInfo.moniteurs = e.target.value;
            updatePersonInfoDisplay();
        });
    }
    
    // Champs sur la page nouvelle évaluation
    const newNomField = document.getElementById('newEvalNom');
    const newPrenomField = document.getElementById('newEvalPrenom');
    const newDateField = document.getElementById('newEvalDateNaissance');
    const newMoniteursField = document.getElementById('newEvalMoniteurs');
    
    if (newNomField) {
        newNomField.addEventListener('input', (e) => {
            personInfo.nom = e.target.value.trim().toUpperCase();
            updatePersonInfoDisplay();
        });
    }
    
    if (newPrenomField) {
        newPrenomField.addEventListener('input', (e) => {
            const parts = e.target.value.trim().split('-');
            personInfo.prenom = parts.map(part => 
                part.charAt(0).toUpperCase() + part.slice(1).toLowerCase()
            ).join('-');
            updatePersonInfoDisplay();
        });
    }
    
    if (newDateField) {
    newDateField.addEventListener('input', function(e) {
        let value = e.target.value.replace(/[^0-9]/g, ''); // Garder uniquement les chiffres
        
        // Limiter à 8 chiffres
        if (value.length > 8) {
            value = value.substring(0, 8);
        }
        
        // Ajouter automatiquement les "/"
        let formatted = '';
        if (value.length > 0) {
            formatted = value.substring(0, 2);
            if (value.length >= 3) {
                formatted += '/' + value.substring(2, 4);
            }
            if (value.length >= 5) {
                formatted += '/' + value.substring(4, 8);
            }
        }
        
        // Mettre à jour le champ
        e.target.value = formatted;
        
        // Sauvegarder si la date est complète (10 caractères = JJ/MM/AAAA)
        if (formatted.length === 10) {
            personInfo.dateNaissance = formatted;
            updatePersonInfoDisplay();
        }
    });
}
    
    if (newMoniteursField) {
        newMoniteursField.addEventListener('input', (e) => {
            personInfo.moniteurs = e.target.value;
            updatePersonInfoDisplay();
        });
    }
}

function updatePersonInfoDisplay() {
    // Synchroniser tous les champs
    const fields = [
        {id: 'personNom', value: personInfo.nom},
        {id: 'personPrenom', value: personInfo.prenom},
        {id: 'personDateNaissance', value: personInfo.dateNaissance},
        {id: 'personMoniteurs', value: personInfo.moniteurs},
        {id: 'newEvalNom', value: personInfo.nom},
        {id: 'newEvalPrenom', value: personInfo.prenom},
        {id: 'newEvalDateNaissance', value: personInfo.dateNaissance},
        {id: 'newEvalMoniteurs', value: personInfo.moniteurs}
    ];
    
    fields.forEach(field => {
        const element = document.getElementById(field.id);
        if (element && element.value !== field.value) {
            element.value = field.value;
        }
    });
}

function showPage(pageId) {
    document.querySelectorAll('.page').forEach(page => {
        page.classList.remove('active');
    });
    document.getElementById(pageId).classList.add('active');
}

// Gestion des fichiers
function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    currentFileName = file.name;
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        currentWorkbook = XLSX.read(data, {type: 'array'});
        
        // Extraire les informations de la personne depuis le nom du fichier
        extractPersonInfo(currentFileName);
        
        // Si le fichier ne contient pas les infos dans le nom, essayer de les extraire de la première feuille
        if (!personInfo.nom && currentWorkbook.SheetNames.length > 0) {
            extractPersonInfoFromSheet(currentWorkbook.SheetNames[0]);
        }
        
        // Récupérer toutes les feuilles (dates)
        allSheetNames = currentWorkbook.SheetNames;
        
        // Afficher la page principale
        displayPersonInfo();
        showPage('mainPage');
    };
    
    reader.readAsArrayBuffer(file);
}

function extractPersonInfo(fileName) {
    // Format attendu: NOM Prénom - JJMMAAAA.xlsx (avec support des tirets dans noms/prénoms)
    const nameWithoutExt = fileName.replace(/\.(xlsx|ods)$/i, '');
    const parts = nameWithoutExt.split(' - ');
    
    if (parts.length >= 2) {
        const namePart = parts[0].trim();
        // Trouver le dernier mot en majuscules (qui sera le début du prénom)
        const words = namePart.split(' ');
        let nomEndIndex = 0;
        
        // Identifier où se termine le NOM (dernière partie en majuscules)
        for (let i = 0; i < words.length; i++) {
            if (words[i] === words[i].toUpperCase() && words[i].length > 0) {
                nomEndIndex = i;
            } else {
                break;
            }
        }
        
        personInfo.nom = words.slice(0, nomEndIndex + 1).join(' ');
        personInfo.prenom = words.slice(nomEndIndex + 1).join(' ');
        
        const dateStr = parts[1].trim();
        if (dateStr.length === 8) {
            const day = dateStr.substring(0, 2);
            const month = dateStr.substring(2, 4);
            const year = dateStr.substring(4, 8);
            personInfo.dateNaissance = `${day}/${month}/${year}`;
        }
    } else {
        personInfo.nom = '';
        personInfo.prenom = '';
        personInfo.dateNaissance = '';
    }
    
    personInfo.moniteurs = '';
}

// Fonction pour valider le format de date
function isValidDate(dateStr) {
    if (!dateStr) return false;
    
    // Accepter les formats JJ/MM/AAAA ou JJMMAAAA
    const regexSlash = /^(\d{2})\/(\d{2})\/(\d{4})$/;
    const regexNoSlash = /^(\d{8})$/;
    
    let day, month, year;
    
    if (regexSlash.test(dateStr)) {
        const match = dateStr.match(regexSlash);
        day = parseInt(match[1], 10);
        month = parseInt(match[2], 10);
        year = parseInt(match[3], 10);
    } else if (regexNoSlash.test(dateStr)) {
        day = parseInt(dateStr.substring(0, 2), 10);
        month = parseInt(dateStr.substring(2, 4), 10);
        year = parseInt(dateStr.substring(4, 8), 10);
    } else {
        return false;
    }
    
    // Vérifier que la date est valide
    if (month < 1 || month > 12) return false;
    if (day < 1 || day > 31) return false;
    if (year < 1900 || year > 2100) return false;
    
    // Vérifier les jours selon le mois
    const daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    
    // Année bissextile
    if ((year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0)) {
        daysInMonth[1] = 29;
    }
    
    if (day > daysInMonth[month - 1]) return false;
    
    return true;
}

// Fonction pour formater la date au format JJ/MM/AAAA
function formatDate(dateStr) {
    if (!dateStr) return '';
    
    // Si déjà au format JJ/MM/AAAA
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) {
        return dateStr;
    }
    
    // Si au format JJMMAAAA
    if (/^\d{8}$/.test(dateStr)) {
        return `${dateStr.substring(0, 2)}/${dateStr.substring(2, 4)}/${dateStr.substring(4, 8)}`;
    }
    
    return dateStr;
}

function extractPersonInfoFromSheet(sheetName) {
    const worksheet = currentWorkbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ''});
    
    // Ligne 1 contient : NOM | Prénom | Date de naissance | (vide) | (vide) | Moniteurs
    if (data.length > 0) {
        personInfo.nom = data[0][0] || '';
        personInfo.prenom = data[0][1] || '';
        personInfo.dateNaissance = data[0][2] || '';
        personInfo.moniteurs = data[0][5] || ''; // Colonne F (index 5)
    }
}

function displayPersonInfo() {
    // Afficher dans les champs éditables
    document.getElementById('personNom').value = personInfo.nom;
    document.getElementById('personPrenom').value = personInfo.prenom;
    document.getElementById('personDateNaissance').value = personInfo.dateNaissance;
    document.getElementById('personMoniteurs').value = personInfo.moniteurs;
}

// Page consultation
function showConsultPage() {
    showPage('consultPage');
    displaySheetTabs();
    
    // Afficher la dernière feuille par défaut
    if (allSheetNames.length > 0) {
        displaySheet(allSheetNames[allSheetNames.length - 1]);
    }
}

function displaySheetTabs() {
    const tabsContainer = document.getElementById('sheetTabs');
    tabsContainer.innerHTML = '';

    // Trier les feuilles par date (MMAAAA)
    const sortedSheets = [...allSheetNames].sort((a, b) => {
        const dateA = parseSheetDate(a);
        const dateB = parseSheetDate(b);
        return dateA - dateB;
    });

    sortedSheets.forEach(sheetName => {
        const tab = document.createElement('div');
        tab.className = 'sheet-tab';
        tab.textContent = formatSheetName(sheetName);
        tab.addEventListener('click', () => {
            document.querySelectorAll('.sheet-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            displaySheet(sheetName);
        });
        tabsContainer.appendChild(tab);
    });

    // Activer le dernier onglet
    if (sortedSheets.length > 0) {
        tabsContainer.lastChild.classList.add('active');
    }

    // Bouton export toutes évaluations
    const exportAllContainer = document.getElementById('exportAllBtnContainer');
    
    // Vider le conteneur au cas où
    exportAllContainer.innerHTML = '';
    
    // Créer le bouton
    const exportAllBtn = document.createElement('button');
    exportAllBtn.className = 'btn-export';
    exportAllBtn.textContent = '📚 Exporter toutes les évaluations (DOCX)';
    exportAllBtn.onclick = exportAllSheetsToWord;
    
    // Insérer dans le conteneur dédié
    exportAllContainer.appendChild(exportAllBtn);
}

function parseSheetDate(sheetName) {
    // Format MMAAAA
    if (sheetName.length === 6) {
        const month = parseInt(sheetName.substring(0, 2));
        const year = parseInt(sheetName.substring(2, 6));
        return new Date(year, month - 1);
    }
    return new Date(0);
}

function formatSheetName(sheetName) {
    // Convertir MMAAAA en MM/AAAA
    if (sheetName.length === 6) {
        return `${sheetName.substring(0, 2)}/${sheetName.substring(2, 6)}`;
    }
    return sheetName;
}

function displaySheet(sheetName) {
    const worksheet = currentWorkbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ''});
    
    // Extraire les informations de la personne de cette évaluation
    const sheetPersonInfo = {
        nom: data[0][0] || '',
        prenom: data[0][1] || '',
        dateNaissance: data[0][2] || '',
        moniteurs: data[0][5] || ''
    };
    
    currentSheetData = parseSheetData(data);
        // Bouton export cette évaluation
    const exportBtn = document.createElement('button');
    exportBtn.className = 'btn-export';
    exportBtn.textContent = '📄 Exporter cette évaluation (DOCX)';
    exportBtn.onclick = () => exportSheetToWord(sheetName);
    
    const consultContent = document.getElementById('consultContent');
    consultContent.innerHTML = ''; // Vider le contenu précédent
    consultContent.appendChild(exportBtn);
    renderEvaluation('consultContent', currentSheetData, true, sheetPersonInfo);
}

// Parser les données de la feuille
function parseSheetData(data) {
    const questions = [];
    
    // Ligne 1: informations personne (déjà extraites)
    // Ligne 2: vide
    // À partir de ligne 3: questions
    
    let i = 2; // Index 2 = ligne 3 (0-indexed)
    while (i < data.length) {
        if (data[i] && data[i][0]) { // Si la cellule A contient quelque chose
            const question = {
                title: data[i][0],
                answers: [],
                selectedAnswers: [],
                observation: ''
            };
            
            // Ligne suivante: réponses possibles
            if (i + 1 < data.length && data[i + 1]) {
                for (let j = 0; j < 10; j++) {
                    if (data[i + 1][j]) {
                        question.answers.push(data[i + 1][j]);
                    }
                }
            }
            
            // Ligne suivante: choix (0 ou 1)
            if (i + 2 < data.length && data[i + 2]) {
                for (let j = 0; j < question.answers.length; j++) {
                    question.selectedAnswers.push(data[i + 2][j] == 1);
                }
            }
            
            // Ligne suivante: observation
            if (i + 3 < data.length && data[i + 3] && data[i + 3][0]) {
                question.observation = data[i + 3][0];
            }
            
            questions.push(question);
            i += 5; // Question + réponses + choix + observation + ligne vide
        } else {
            i++;
        }
    }
    
    return questions;
}

// Afficher une évaluation
function renderEvaluation(containerId, questionsData, readonly = false, sheetPersonInfo = null) {
    const container = document.getElementById(containerId);
    container.innerHTML = '';
    
    // Si en mode consultation et qu'on a les infos de la personne, les afficher
    if (readonly && sheetPersonInfo) {
        const personInfoBlock = document.createElement('div');
        personInfoBlock.className = 'person-info-display';
        personInfoBlock.innerHTML = `
            <div class="info-row">
                <strong>NOM :</strong> ${sheetPersonInfo.nom}
            </div>
            <div class="info-row">
                <strong>Prénom :</strong> ${sheetPersonInfo.prenom}
            </div>
            <div class="info-row">
                <strong>Date de naissance :</strong> ${sheetPersonInfo.dateNaissance}
            </div>
            <div class="info-row moniteurs-row">
                <strong>Moniteur(s) présent lors de l'évaluation :</strong> ${sheetPersonInfo.moniteurs || 'Non renseigné'}
            </div>
        `;
        container.appendChild(personInfoBlock);
        
        const separator = document.createElement('hr');
        separator.style.margin = '20px 0';
        separator.style.border = 'none';
        separator.style.borderTop = '2px solid #dee2e6';
        container.appendChild(separator);
    }
    
    questionsData.forEach((question, qIndex) => {
        const questionBlock = document.createElement('div');
        questionBlock.className = 'question-block';
        
        // Titre de la question
        const questionTitle = document.createElement('div');
        questionTitle.className = 'question-title';
        questionTitle.textContent = question.title;
        questionBlock.appendChild(questionTitle);
        
        // Réponses
        const answersContainer = document.createElement('div');
        answersContainer.className = 'answers-container';
        
        question.answers.forEach((answer, aIndex) => {
            const answerOption = document.createElement('div');
            answerOption.className = 'answer-option';
            
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox';
            checkbox.id = `${containerId}_q${qIndex}_a${aIndex}`;
            checkbox.checked = question.selectedAnswers[aIndex] || false;
            checkbox.disabled = readonly;
            
            if (!readonly) {
                checkbox.addEventListener('change', function() {
                    question.selectedAnswers[aIndex] = this.checked;
                });
            }
            
            const label = document.createElement('label');
            label.htmlFor = checkbox.id;
            label.textContent = answer;
            
            answerOption.appendChild(checkbox);
            answerOption.appendChild(label);
            answersContainer.appendChild(answerOption);
        });
        
        questionBlock.appendChild(answersContainer);
        
        // Observation
        const observationContainer = document.createElement('div');
        observationContainer.className = 'observation-container';
        
        const observationLabel = document.createElement('label');
        observationLabel.className = 'observation-label';
        observationLabel.textContent = 'Observation :';
        observationContainer.appendChild(observationLabel);
        
        if (readonly) {
            const observationText = document.createElement('div');
            observationText.className = 'observation-readonly';
            observationText.textContent = question.observation || 'Aucune observation';
            observationContainer.appendChild(observationText);
        } else {
            const observationTextarea = document.createElement('textarea');
            observationTextarea.className = 'observation-text';
            observationTextarea.maxLength = 3000;
            observationTextarea.value = question.observation || '';
            observationTextarea.placeholder = 'Saisissez vos observations (max 3000 caractères)...';
            
            observationTextarea.addEventListener('input', function() {
                question.observation = this.value;
            });
            
            observationContainer.appendChild(observationTextarea);
        }
        
        questionBlock.appendChild(observationContainer);
        container.appendChild(questionBlock);
    });
}

// Page nouvelle évaluation
function showNewEvalPage() {
    showPage('newEvalPage');
    
    // Mettre à jour les champs d'informations personnelles
    document.getElementById('newEvalNom').value = personInfo.nom;
    document.getElementById('newEvalPrenom').value = personInfo.prenom;
    document.getElementById('newEvalDateNaissance').value = personInfo.dateNaissance;
    document.getElementById('newEvalMoniteurs').value = personInfo.moniteurs;
    
    // Trouver la dernière évaluation
    let lastSheetName = '';
    if (allSheetNames.length > 0) {
        const sortedSheets = [...allSheetNames].sort((a, b) => {
            return parseSheetDate(b) - parseSheetDate(a);
        });
        lastSheetName = sortedSheets[0];
    }
    
    // Afficher la date de la dernière évaluation
    document.getElementById('lastEvalDate').textContent = lastSheetName ? `(${formatSheetName(lastSheetName)})` : '(Aucune)';
    
    // Afficher la date actuelle pour la nouvelle évaluation
    const now = new Date();
    const currentDateStr = `${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`;
    document.getElementById('currentDate').textContent = `(${currentDateStr})`;
    
    // Charger et afficher l'ancienne évaluation
    if (lastSheetName) {
        const worksheet = currentWorkbook.Sheets[lastSheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ''});
        
        // Extraire les infos de la personne de l'ancienne évaluation
        const oldPersonInfo = {
            nom: data[0][0] || '',
            prenom: data[0][1] || '',
            dateNaissance: data[0][2] || '',
            moniteurs: data[0][5] || ''
        };
        
        const oldEvalData = parseSheetData(data);
        renderEvaluation('oldEvalContent', oldEvalData, true, oldPersonInfo);
        
        // Créer une nouvelle évaluation basée sur la structure de l'ancienne
        templateQuestions = oldEvalData.map(q => ({
            title: q.title,
            answers: [...q.answers],
            selectedAnswers: new Array(q.answers.length).fill(false),
            observation: ''
        }));
    } else {
        // Utiliser le questionnaire de base si aucune évaluation n'existe
        document.getElementById('oldEvalContent').innerHTML = '<p style="text-align: center; color: #999;">Aucune évaluation précédente</p>';
        templateQuestions = baseQuestionnaire.map(q => ({
            title: q.question,
            answers: [...q.answers],
            selectedAnswers: new Array(q.answers.length).fill(false),
            observation: ''
        }));
    }
    
    // Afficher la nouvelle évaluation vierge
    renderEvaluation('newEvalContent', templateQuestions, false);
}

// Sauvegarder l'évaluation
async function saveEvaluation(format) {
    // Valider les informations personnelles
if (!personInfo.nom.trim() || !personInfo.prenom.trim() || !personInfo.dateNaissance.trim()) {
    alert('Veuillez renseigner le NOM, le Prénom et la Date de naissance avant de sauvegarder.');
    return;
}

// Valider et formater la date
if (!isValidDate(personInfo.dateNaissance)) {
    alert('Le format de la date de naissance doit être JJ/MM/AAAA (ex: 26/01/1990)');
    return;
}

// S'assurer que la date est au bon format
personInfo.dateNaissance = formatDate(personInfo.dateNaissance);
    
    // Si pas de workbook existant, en créer un nouveau
    if (!currentWorkbook) {
        currentWorkbook = XLSX.utils.book_new();
        allSheetNames = [];
    }
    
    // Créer le nom de la nouvelle feuille (MMAAAA)
    const now = new Date();
    const newSheetName = `${String(now.getMonth() + 1).padStart(2, '0')}${now.getFullYear()}`;
    
    // Vérifier si une feuille avec ce nom existe déjà
    if (currentWorkbook.Sheets[newSheetName]) {
        if (!confirm(`Une évaluation pour ${formatSheetName(newSheetName)} existe déjà. Voulez-vous la remplacer ?`)) {
            return;
        }
    }
    
    // Créer les données de la feuille
    const sheetData = [];
    
    // Ligne 1: Informations de la personne
    // A: NOM | B: Prénom | C: Date de naissance | D: vide | E: vide | F: Moniteurs
    sheetData.push([
        personInfo.nom, 
        personInfo.prenom, 
        personInfo.dateNaissance, 
        '', 
        '', 
        personInfo.moniteurs
    ]);
    
    // Ligne 2: vide
    sheetData.push([]);
    
    // Questions
    templateQuestions.forEach(question => {
        // Ligne question
        sheetData.push([question.title]);
        
        // Ligne réponses
        const answersRow = [];
        question.answers.forEach(answer => {
            answersRow.push(answer);
        });
        sheetData.push(answersRow);
        
        // Ligne choix (0 ou 1)
        const choicesRow = [];
        question.selectedAnswers.forEach(selected => {
            choicesRow.push(selected ? 1 : 0);
        });
        sheetData.push(choicesRow);
        
        // Ligne observation
        sheetData.push([question.observation || '']);
        
        // Ligne vide
        sheetData.push([]);
    });
    
    // Créer la nouvelle feuille
    const newWorksheet = XLSX.utils.aoa_to_sheet(sheetData);
    
    // Ajouter ou remplacer la feuille dans le workbook
    currentWorkbook.Sheets[newSheetName] = newWorksheet;
    
    // Ajouter le nom de la feuille si elle n'existe pas
    if (!currentWorkbook.SheetNames.includes(newSheetName)) {
        currentWorkbook.SheetNames.push(newSheetName);
    }
    
    // Générer le nom du fichier à partir des infos personnelles
    const dateForFileName = personInfo.dateNaissance.replace(/\//g, '');
    const fileName = `${personInfo.nom} ${personInfo.prenom} - ${dateForFileName}.${format}`;
    
    // Sauvegarder le fichier
    const wbout = XLSX.write(currentWorkbook, {
        bookType: format === 'ods' ? 'ods' : (format === 'xls' ? 'xls' : 'xlsx'),
        type: 'array'
    });
    
    const blob = new Blob([wbout], {type: 'application/octet-stream'});
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(url);
    
    alert('Évaluation sauvegardée avec succès !');
    
    // ✅ EXPORT AUTOMATIQUE EN DOCX
    try {
        await exportSheetToWord(newSheetName);
    } catch (error) {
        console.error('Erreur export DOCX:', error);
        alert('Le fichier Excel a été sauvegardé, mais l\'export Word a échoué.');
    }
    
    // Mettre à jour la liste des feuilles
    allSheetNames = currentWorkbook.SheetNames;
}

// Export d'une évaluation en DOCX
async function exportSheetToWord(sheetName) {
    const { Document, Paragraph, TextRun, AlignmentType } = docx;
    
    const worksheet = currentWorkbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ''});
    
    // Extraire infos personne
    const nom = data[0][0] || '';
    const prenom = data[0][1] || '';
    const dateNaissance = data[0][2] || '';
    const moniteurs = data[0][5] || '';
    
    // Parser les questions
    const questions = parseSheetData(data);
    
    // Créer le document
    const docParagraphs = [];
    
    // En-tête
    docParagraphs.push(
        new Paragraph({
            children: [new TextRun({ text: `${prenom} ${nom}`, bold: true, size: 32 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 }
        }),
        new Paragraph({
            children: [new TextRun({ text: `Date de naissance : ${dateNaissance}`, size: 24 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 }
        }),
        new Paragraph({
            children: [new TextRun({ text: `Date d'évaluation : ${formatSheetName(sheetName)}`, size: 24 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 }
        }),
        new Paragraph({
            children: [new TextRun({ text: `Moniteur(s) : ${moniteurs}`, size: 24 })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 }
        })
    );
    
    // Questions
    questions.forEach((q, index) => {
        // Titre question
        docParagraphs.push(
            new Paragraph({
                children: [new TextRun({ text: `N°${index + 1} - ${q.title}`, bold: true, size: 24 })],
                spacing: { before: 300, after: 100 }
            })
        );
        
        // Réponses avec emojis
        const answersText = q.answers.map((answer, i) => 
            `${q.selectedAnswers[i] ? '☑️' : '🔲'} ${answer}`
        ).join('  ');
        
        docParagraphs.push(
            new Paragraph({
                children: [new TextRun({ text: answersText, size: 22 })],
                spacing: { after: 100 }
            })
        );
        
        // Observation
        docParagraphs.push(
            new Paragraph({
                children: [
                    new TextRun({ text: 'Observation : ', bold: true, size: 22 }),
                    new TextRun({ text: q.observation || '', size: 22 })
                ],
                spacing: { after: 200 }
            })
        );
    });
    
    const doc = new Document({ sections: [{ children: docParagraphs }] });
    
    // Générer et télécharger
    const blob = await docx.Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    
    // Nom fichier : NOM Prénom - JJMMAAAA - ÉVALUATION.docx
    const dateStr = sheetName.replace(/\//g, '');
    link.download = `${nom} ${prenom} - ${dateStr} - ÉVALUATION.docx`;
    link.click();
    URL.revokeObjectURL(url);
}

// Export de toutes les évaluations en DOCX
async function exportAllSheetsToWord() {
    const { Document, Paragraph, TextRun, AlignmentType, PageBreak } = docx;
    
    const allParagraphs = [];
    const nom = personInfo.nom;
    const prenom = personInfo.prenom;
    const dateNaissance = personInfo.dateNaissance;
    
    // Pour chaque feuille
    for (let sheetIndex = 0; sheetIndex < allSheetNames.length; sheetIndex++) {
        const sheetName = allSheetNames[sheetIndex];
        const worksheet = currentWorkbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ''});
        
        const moniteurs = data[0][5] || '';
        const questions = parseSheetData(data);
        
        // En-tête de l'évaluation
        allParagraphs.push(
            new Paragraph({
                children: [new TextRun({ text: `${prenom} ${nom}`, bold: true, size: 32 })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 }
            }),
            new Paragraph({
                children: [new TextRun({ text: `Date de naissance : ${dateNaissance}`, size: 24 })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 }
            }),
            new Paragraph({
                children: [new TextRun({ text: `Date d'évaluation : ${formatSheetName(sheetName)}`, size: 24 })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 }
            }),
            new Paragraph({
                children: [new TextRun({ text: `Moniteur(s) : ${moniteurs}`, size: 24 })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 }
            })
        );
        
        // Questions
        questions.forEach((q, index) => {
            allParagraphs.push(
                new Paragraph({
                    children: [new TextRun({ text: `N°${index + 1} - ${q.title}`, bold: true, size: 24 })],
                    spacing: { before: 300, after: 100 }
                })
            );
            
            const answersText = q.answers.map((answer, i) => 
                `${q.selectedAnswers[i] ? '☑️' : '🔲'} ${answer}`
            ).join('  ');
            
            allParagraphs.push(
                new Paragraph({
                    children: [new TextRun({ text: answersText, size: 22 })],
                    spacing: { after: 100 }
                })
            );
            
            allParagraphs.push(
                new Paragraph({
                    children: [
                        new TextRun({ text: 'Observation : ', bold: true, size: 22 }),
                        new TextRun({ text: q.observation || '', size: 22 })
                    ],
                    spacing: { after: 200 }
                })
            );
        });
        
        // Saut de page sauf pour la dernière
        if (sheetIndex < allSheetNames.length - 1) {
            allParagraphs.push(new Paragraph({ children: [new PageBreak()] }));
        }
    }
    
    const doc = new Document({ sections: [{ children: allParagraphs }] });
    
    const blob = await docx.Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    
    // Format date JJMMAAAA du jour
    const today = new Date();
    const dateStr = String(today.getDate()).padStart(2, '0') + 
                   String(today.getMonth() + 1).padStart(2, '0') + 
                   today.getFullYear();
    
    link.download = `${nom} ${prenom} - ${dateStr} - HISTORIQUE.docx`;
    link.click();
    URL.revokeObjectURL(url);
}

// Export de TOUTES les évaluations en un seul DOCX
async function exportAllSheetsToWord() {
    const { Document, Paragraph, TextRun, AlignmentType, PageBreak } = docx;

    if (!currentWorkbook) {
        alert('Aucun fichier chargé');
        return;
    }

    // Filtrer uniquement les feuilles au format MMAAAA (6 chiffres)
    const validSheets = allSheetNames.filter(sheetName => /^\d{6}$/.test(sheetName));

    if (validSheets.length === 0) {
        alert('Aucune évaluation valide trouvée (format MMAAAA requis)');
        return;
    }

    // Trier les feuilles par date (plus ancienne en premier)
    const sortedSheets = validSheets.sort((a, b) => {
        return parseSheetDate(a) - parseSheetDate(b);
    });

    const allParagraphs = [];
    
    // Pour chaque évaluation
    for (let i = 0; i < sortedSheets.length; i++) {
        const sheetName = sortedSheets[i];
        const worksheet = currentWorkbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, {header: 1, defval: ''});
        
        // Extraire infos personne
        const nom = data[0][0] || '';
        const prenom = data[0][1] || '';
        const dateNaissance = data[0][2] || '';
        const moniteurs = data[0][5] || '';
        
        // Parser les questions
        const questions = parseSheetData(data);
        
        // En-tête de cette évaluation
        allParagraphs.push(
            new Paragraph({
                children: [new TextRun({ text: `${prenom} ${nom}`, bold: true, size: 32 })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 }
            }),
            new Paragraph({
                children: [new TextRun({ text: `Date de naissance : ${dateNaissance}`, size: 24 })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 }
            }),
            new Paragraph({
                children: [new TextRun({ text: `Date d'évaluation : ${formatSheetName(sheetName)}`, size: 24 })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 }
            }),
            new Paragraph({
                children: [new TextRun({ text: `Moniteur(s) : ${moniteurs}`, size: 24 })],
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 }
            })
        );
        
        // Questions
        questions.forEach((q, index) => {
            // Titre question
            allParagraphs.push(
                new Paragraph({
                    children: [new TextRun({ text: `N°${index + 1} - ${q.title}`, bold: true, size: 24 })],
                    spacing: { before: 300, after: 100 }
                })
            );
            
            // Réponses avec emojis
            const answersText = q.answers.map((answer, i) => 
                `${q.selectedAnswers[i] ? '☑️' : '🔲'} ${answer}`
            ).join('  ');
            
            allParagraphs.push(
                new Paragraph({
                    children: [new TextRun({ text: answersText, size: 22 })],
                    spacing: { after: 100 }
                })
            );
            
            // Observation
            allParagraphs.push(
                new Paragraph({
                    children: [new TextRun({ text: `Observation : ${q.observation || ''}`, size: 22, italics: true })],
                    spacing: { after: 200 }
                })
            );
        });
        
        // Saut de page entre les évaluations (sauf pour la dernière)
        if (i < sortedSheets.length - 1) {
            allParagraphs.push(
                new Paragraph({
                    children: [new PageBreak()],
                    spacing: { after: 400 }
                })
            );
        }
    }
    
    // Créer le document
    const doc = new Document({
        sections: [{
            properties: {},
            children: allParagraphs
        }]
    });
    
    // Générer et télécharger
    const blob = await docx.Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${personInfo.nom}_${personInfo.prenom}_TOUTES_EVALUATIONS.docx`;
    link.click();
    URL.revokeObjectURL(url);
    
    alert(`✅ ${sortedSheets.length} évaluation(s) exportée(s) avec succès !`);
}

// Modal guide
function openGuideModal() {
    document.getElementById('guideModal').classList.add('active');
}

function closeGuideModal() {
    document.getElementById('guideModal').classList.remove('active');
}

// Exporter le questionnaire de base
function exportBaseQuestionnaire(format) {
    try {
        console.log('Début exportation, format:', format);
        console.log('Nombre de questions:', baseQuestionnaire.length);
        
        const sheetData = [];

        // Ligne 1: En-têtes
        sheetData.push([
            'NOM', 
            'Prénom', 
            'Date de naissance (JJ/MM/AAAA)', 
            '', 
            '', 
            'Moniteur(s) présent lors de l\'évaluation'
        ]);

        // Ligne 2: vide pour la saisie
        sheetData.push(['', '', '', '', '', '']);

        // Ligne 3: vide de séparation
        sheetData.push([]);

        // Questions de base
        baseQuestionnaire.forEach((question, index) => {
            console.log(`Question ${index + 1}:`, question.question);
            
            // Ligne question
            sheetData.push([question.question]);

            // Ligne réponses
            const answersRow = [];
            question.answers.forEach(answer => {
                answersRow.push(answer);
            });
            // Compléter jusqu'à 10 colonnes
            while (answersRow.length < 10) {
                answersRow.push('');
            }
            sheetData.push(answersRow);

            // Ligne choix (vides pour cocher)
            sheetData.push(new Array(10).fill(''));

            // Ligne observation
            const obsRow = new Array(10).fill('');
            obsRow[5] = ''; // Cellule F pour l'observation
            sheetData.push(obsRow);

            // Ligne vide de séparation
            sheetData.push([]);
        });

        console.log('Données préparées, nombre de lignes:', sheetData.length);

        // Créer le workbook et la feuille
        const wb = XLSX.utils.book_new();
        console.log('Workbook créé');
        
        const ws = XLSX.utils.aoa_to_sheet(sheetData);
        console.log('Feuille créée');
        
        // Définir les largeurs de colonnes
        ws['!cols'] = [
            { wch: 50 },  // A: Question
            { wch: 15 },  // B: Réponse 1
            { wch: 15 },  // C: Réponse 2
            { wch: 15 },  // D: Réponse 3
            { wch: 15 },  // E: Réponse 4
            { wch: 30 },  // F: Observation
            { wch: 15 },  // G: Réponse 5
            { wch: 15 },  // H: Réponse 6
            { wch: 15 },  // I: Réponse 7
            { wch: 15 }   // J: Réponse 8
        ];
        
        XLSX.utils.book_append_sheet(wb, ws, 'Questionnaire');
        console.log('Feuille ajoutée au workbook');

        // Déterminer le type de fichier
        const bookType = format === 'ods' ? 'ods' : 'xlsx';
        console.log('Type de fichier:', bookType);

        // Sauvegarder
        const fileName = `Questionnaire_Base.${format}`;
        console.log('Nom du fichier:', fileName);
        
        const wbout = XLSX.write(wb, {
            bookType: bookType,
            type: 'array'
        });
        console.log('Fichier généré');

        const blob = new Blob([wbout], {type: 'application/octet-stream'});
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        
        console.log('Téléchargement lancé');
        alert('Questionnaire de base exporté avec succès !');
        
    } catch (error) {
        console.error('ERREUR lors de l\'exportation:', error);
        console.error('Stack:', error.stack);
        alert('Erreur lors de l\'exportation du questionnaire.\nDétails: ' + error.message + '\n\nOuvrez la console (F12) pour plus d\'informations.');
    }
}