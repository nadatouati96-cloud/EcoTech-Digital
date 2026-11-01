const pptxgen = require('pptxgenjs');
const fs = require('fs');
const path = require('path');

// Initialisation de la présentation
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9'; // 10 x 5.625 pouces
pres.author = 'EcoTech Digital Team';
pres.company = 'EcoTech Digital';
pres.subject = 'Soutenance de Projet - Techniques de recherche d’emploi et Marketing de soi';
pres.title = 'EcoTech Digital — Le numérique au service de la planète';

// Palette de couleurs de la marque
const C = {
  FOREST_GREEN: '103B2B',   // Vert Forêt Profond (marque principale)
  EMERALD: '2D6A4F',        // Vert Éco Dynamique
  MINT: '52B788',           // Vert Menthe / Accent
  LIGHT_MINT: 'D8F3DC',     // Vert clair très doux
  DARK_TEXT: '1E293B',      // Gris ardoise foncé
  MUTED_TEXT: '64748B',     // Gris intermédiaire
  LIGHT_BG: 'F8FAF8',       // Fond écru / blanc cassé écologique
  CARD_BG: 'FFFFFF',        // Blanc pur pour les cartes
  CARD_BORDER: 'E2E8F0',    // Bordure subtile
  BEIGE: 'F4F1DE',          // Beige doux
  WHITE: 'FFFFFF',
  ACCENT_GOLD: 'D4A373'
};

// Polices
const FONT_TITLE = 'Arial'; // Compatible universellement
const FONT_BODY = 'Arial';

// Fonction d'aide pour ajouter l'en-tête standard
function addSlideHeader(slide, category, title, subtitle) {
  // Fond général
  slide.background = { color: C.LIGHT_BG };

  // Barre supérieure d'accentuation
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.08,
    fill: { color: C.FOREST_GREEN }, line: { color: C.FOREST_GREEN }
  });

  // Badge de catégorie
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 0.35, w: 2.8, h: 0.28, rectRadius: 0.08,
    fill: { color: C.LIGHT_MINT }, line: { color: C.MINT, width: 1 }
  });
  slide.addText(category.toUpperCase(), {
    x: 0.6, y: 0.35, w: 2.8, h: 0.28,
    fontSize: 9, bold: true, color: C.FOREST_GREEN,
    fontFace: FONT_TITLE, align: 'center', valign: 'middle'
  });

  // Titre principal
  slide.addText(title, {
    x: 0.6, y: 0.68, w: 8.8, h: 0.45,
    fontSize: 20, bold: true, color: C.FOREST_GREEN,
    fontFace: FONT_TITLE, valign: 'middle'
  });

  // Sous-titre explicatif
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.6, y: 1.15, w: 8.8, h: 0.28,
      fontSize: 11, italic: false, color: C.MUTED_TEXT,
      fontFace: FONT_BODY, valign: 'top'
    });
  }

  // Pied de page discret
  slide.addText('EcoTech Digital — « Le numérique au service de la planète » | EcoConnect', {
    x: 0.6, y: 5.3, w: 7.0, h: 0.25,
    fontSize: 8.5, color: C.MUTED_TEXT, fontFace: FONT_BODY
  });
  
  slide.addText('Projet Académique — TRE & Marketing de soi', {
    x: 7.6, y: 5.3, w: 1.8, h: 0.25,
    fontSize: 8.5, color: C.MUTED_TEXT, fontFace: FONT_BODY, align: 'right'
  });
}

// ==========================================
// SLIDE 1 : COUVERTURE (Page de Garde)
// ==========================================
{
  const s1 = pres.addSlide();
  s1.background = { color: C.FOREST_GREEN };

  // Motifs géométriques décoratifs modernes
  s1.addShape(pres.shapes.OVAL, {
    x: 7.0, y: -1.0, w: 4.5, h: 4.5,
    fill: { color: C.EMERALD, transparency: 75 },
    line: { color: C.MINT, width: 1, transparency: 60 }
  });
  s1.addShape(pres.shapes.OVAL, {
    x: 8.2, y: 2.5, w: 3.5, h: 3.5,
    fill: { color: C.MINT, transparency: 85 },
    line: { color: C.WHITE, width: 1, transparency: 80 }
  });

  // Badge Contexte Académique
  s1.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 0.7, w: 4.2, h: 0.35, rectRadius: 0.1,
    fill: { color: '1B4D3E' }, line: { color: C.MINT, width: 1 }
  });
  s1.addText('PROJET ACADÉMIQUE • MARKETING DE SOI & TRE', {
    x: 0.8, y: 0.7, w: 4.2, h: 0.35,
    fontSize: 9, bold: true, color: C.MINT,
    fontFace: FONT_TITLE, align: 'center', valign: 'middle'
  });

  // Titre Entreprise
  s1.addText('EcoTech Digital', {
    x: 0.8, y: 1.25, w: 6.5, h: 0.8,
    fontSize: 38, bold: true, color: C.WHITE,
    fontFace: FONT_TITLE
  });

  // Slogan Officiel
  s1.addText('« Le numérique au service de la planète. »', {
    x: 0.8, y: 2.05, w: 6.5, h: 0.4,
    fontSize: 16, italic: true, color: C.LIGHT_MINT,
    fontFace: FONT_BODY
  });

  // Tagline Descriptive
  s1.addText('GreenTech tunisienne combinant ingénierie logicielle éco-responsable,\ncommunication digitale éthique et engagement environnemental civique.', {
    x: 0.8, y: 2.55, w: 6.0, h: 0.65,
    fontSize: 11, color: 'E2E8F0', fontFace: FONT_BODY, lineSpacingMultiple: 1.2
  });

  // Encart Produit Phare
  s1.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 3.35, w: 5.8, h: 0.55, rectRadius: 0.08,
    fill: { color: '164E3A' }, line: { color: C.MINT, width: 1 }
  });
  s1.addText('PRODUIT PHARE : EcoConnect', {
    x: 1.0, y: 3.42, w: 2.2, h: 0.4,
    fontSize: 9.5, bold: true, color: C.MINT, fontFace: FONT_TITLE
  });
  s1.addText('|  Plateforme citoyenne de tri et de recyclage en Tunisie', {
    x: 3.2, y: 3.42, w: 3.3, h: 0.4,
    fontSize: 9.5, color: C.WHITE, fontFace: FONT_BODY
  });

  // CARTE INFOS ÉTUDIANT & UNIVERSITÉ (Champs modifiables)
  s1.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 7.0, y: 1.1, w: 2.4, h: 3.8, rectRadius: 0.12,
    fill: { color: C.WHITE }, line: { color: C.CARD_BORDER, width: 1 }
  });

  s1.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 7.15, y: 1.25, w: 2.1, h: 0.32, rectRadius: 0.06,
    fill: { color: C.LIGHT_MINT }
  });
  s1.addText('INFORMATIONS DE SOUTENANCE', {
    x: 7.15, y: 1.25, w: 2.1, h: 0.32,
    fontSize: 7.5, bold: true, color: C.FOREST_GREEN,
    align: 'center', valign: 'middle'
  });

  const studentFields = [
    { label: 'Étudiant(e) / Porteur :', val: '[Prénom NOM]' },
    { label: 'Spécialité / Filière :', val: '[Licence / Mastère Info-Marketing]' },
    { label: 'Établissement :', val: '[Nom de l’Université / Institut]' },
    { label: 'Enseignant(e) :', val: '[Nom du Professeur]' },
    { label: 'Module :', val: 'TRE & Marketing de soi' },
    { label: 'Année universitaire :', val: '2025 – 2026' }
  ];

  let curY = 1.68;
  studentFields.forEach(f => {
    s1.addText(f.label, {
      x: 7.15, y: curY, w: 2.1, h: 0.22,
      fontSize: 7.5, bold: true, color: C.MUTED_TEXT, fontFace: FONT_TITLE
    });
    s1.addText(f.val, {
      x: 7.15, y: curY + 0.18, w: 2.1, h: 0.28,
      fontSize: 8.5, bold: true, color: C.FOREST_GREEN, fontFace: FONT_BODY
    });
    curY += 0.46;
  });

  s1.addNotes(
    "Slide 1 — Page de garde.\n" +
    "Monsieur/Madame le Professeur, chers membres du jury, j'ai l'honneur de vous présenter aujourd'hui le projet EcoTech Digital, développé dans le cadre du module Techniques de recherche d’emploi et Marketing de soi. " +
    "Notre mission résumée par notre slogan : 'Le numérique au service de la planète'. " +
    "Ce projet me permet de valoriser une double compétence : l'ingénierie web et le marketing digital éthique, appliqués à un défi écologique urgent en Tunisie."
  );
}

// ==========================================
// SLIDE 2 : PRÉSENTATION DE L'ENTREPRISE
// ==========================================
{
  const s2 = pres.addSlide();
  addSlideHeader(s2, '01 • Identité & Vision', 'Présentation d’EcoTech Digital', 'Une startup tunisienne pionnière articulant technologies de l’information et transition écologique');

  // Carte Mission (Gauche)
  s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.55, w: 4.25, h: 1.65, rectRadius: 0.1,
    fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
  });
  s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 1.7, w: 1.2, h: 0.28, rectRadius: 0.05,
    fill: { color: C.LIGHT_MINT }
  });
  s2.addText('NOTRE MISSION', {
    x: 0.8, y: 1.7, w: 1.2, h: 0.28,
    fontSize: 8, bold: true, color: C.FOREST_GREEN, align: 'center', valign: 'middle'
  });
  s2.addText('Concevoir des solutions technologiques responsables et accessibles, permettant aux citoyens, entreprises et collectivités de réduire leur empreinte écologique grâce au levier du numérique.', {
    x: 0.8, y: 2.1, w: 3.85, h: 0.95,
    fontSize: 10.5, color: C.DARK_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.2
  });

  // Carte Vision (Droite)
  s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.15, y: 1.55, w: 4.25, h: 1.65, rectRadius: 0.1,
    fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
  });
  s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.35, y: 1.7, w: 1.1, h: 0.28, rectRadius: 0.05,
    fill: { color: 'FEF3C7' }
  });
  s2.addText('NOTRE VISION', {
    x: 5.35, y: 1.7, w: 1.1, h: 0.28,
    fontSize: 8, bold: true, color: '92400E', align: 'center', valign: 'middle'
  });
  s2.addText('Devenir la référence tunisienne de la GreenTech et de l’éco-conception logicielle, en prouvant que l’innovation numérique peut être un accélérateur concret du développement durable en Tunisie.', {
    x: 5.35, y: 2.1, w: 3.85, h: 0.95,
    fontSize: 10.5, color: C.DARK_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.2
  });

  // 4 Piliers Fondateurs (Bas)
  const pillars = [
    { title: 'Éco-Responsabilité', desc: 'Sobriété logicielle, réduction de l’empreinte carbone numérique et hébergement vert.' },
    { title: 'Rigueur Scientifique', desc: 'Informations vérifiées sur les matières recyclables et les filières de revalorisation.' },
    { title: 'Inclusion Citoyenne', desc: 'Interfaces intuitives, multilingues et pensées pour tous les profils en Tunisie.' },
    { title: 'Transparence Totale', desc: 'Honnêteté sur les données environnementales et rejet catégorique du greenwashing.' }
  ];

  pillars.forEach((p, i) => {
    const xPos = 0.6 + i * 2.25;
    s2.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: xPos, y: 3.35, w: 2.15, h: 1.75, rectRadius: 0.08,
      fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
    });
    s2.addText(`0${i+1}`, {
      x: xPos + 0.15, y: 3.45, w: 0.8, h: 0.3,
      fontSize: 14, bold: true, color: C.MINT, fontFace: FONT_TITLE
    });
    s2.addText(p.title, {
      x: xPos + 0.15, y: 3.8, w: 1.85, h: 0.35,
      fontSize: 9.5, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
    });
    s2.addText(p.desc, {
      x: xPos + 0.15, y: 4.18, w: 1.85, h: 0.85,
      fontSize: 8.5, color: C.MUTED_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.15
    });
  });

  s2.addNotes(
    "Slide 2 — Présentation de l'entreprise.\n" +
    "EcoTech Digital est un projet de startup GreenTech ancré dans le contexte tunisien. Notre vocation est double : accompagner les acteurs économiques dans leur transition digitale responsable (via des sites web éco-conçus et du marketing éthique), tout en développant des outils citoyens gratuits comme EcoConnect. " +
    "Nous reposons sur 4 piliers inébranlables : la sobriété logicielle, la rigueur scientifique des données, l'accessibilité citoyenne et la transparence éthique, loin de tout greenwashing."
  );
}

// ==========================================
// SLIDE 3 : PROBLÉMATIQUE
// ==========================================
{
  const s3 = pres.addSlide();
  addSlideHeader(s3, '02 • Contexte & Défi', 'La Problématique Environnementale en Tunisie', 'Le décalage critique entre la volonté citoyenne d’agir et l’absence de repères numériques fiables');

  const problems = [
    {
      num: '01',
      tag: 'OPACITÉ TERRITORIALE',
      title: 'Accès limité aux points de collecte',
      bullets: [
        'Absence de cartographie numérique centralisée et mise à jour.',
        'Les citoyens ignorent où déposer le plastique, le verre ou les piles.',
        'Découragement rapide des initiatives citoyennes spontanées.'
      ]
    },
    {
      num: '02',
      tag: 'CONFUSION TECHNIQUE',
      title: 'Difficulté du tri et méconnaissance',
      bullets: [
        'Consignes de tri floues ou contradictoires pour le grand public.',
        'Confusion majeure entre plastiques valorisables (PET/PEHD) et déchets résiduels.',
        'Contamination des bennes de recyclage par des erreurs évitables.'
      ]
    },
    {
      num: '03',
      tag: 'DÉFICIT DIGITAL',
      title: 'Manque d’outils numériques modernes',
      bullets: [
        'Les associations et acteurs du recyclage manquent de visibilité digitale.',
        'Campagnes de sensibilisation traditionnelles peu engageantes pour les jeunes.',
        'Absence d’outils interactifs mesurant l’impact des éco-gestes du quotidien.'
      ]
    }
  ];

  problems.forEach((p, i) => {
    const xPos = 0.6 + i * 2.98;
    s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: xPos, y: 1.55, w: 2.85, h: 2.7, rectRadius: 0.1,
      fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
    });
    
    // Numéro & Tag
    s3.addText(p.num, {
      x: xPos + 0.2, y: 1.7, w: 0.8, h: 0.35,
      fontSize: 16, bold: true, color: 'DC2626', fontFace: FONT_TITLE
    });
    s3.addText(p.tag, {
      x: xPos + 0.8, y: 1.75, w: 1.85, h: 0.25,
      fontSize: 7.5, bold: true, color: C.MUTED_TEXT, fontFace: FONT_TITLE, align: 'right'
    });

    s3.addText(p.title, {
      x: xPos + 0.2, y: 2.12, w: 2.45, h: 0.45,
      fontSize: 11, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
    });

    let bY = 2.65;
    p.bullets.forEach(b => {
      s3.addText('• ' + b, {
        x: xPos + 0.2, y: bY, w: 2.45, h: 0.48,
        fontSize: 8.5, color: C.DARK_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.15
      });
      bY += 0.5;
    });
  });

  // Bannière Synthèse Défi
  s3.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 4.4, w: 8.8, h: 0.75, rectRadius: 0.08,
    fill: { color: C.FOREST_GREEN }
  });
  s3.addText('LE DÉFI FONDAMENTAL :', {
    x: 0.9, y: 4.48, w: 2.2, h: 0.6,
    fontSize: 9.5, bold: true, color: C.MINT, fontFace: FONT_TITLE
  });
  s3.addText('Comment transformer la sensibilité écologique grandissante de la jeunesse et des organisations tunisiennes en réflexes quotidiens mesurables, grâce à une plateforme numérique simple, accessible et scientifique ?', {
    x: 3.1, y: 4.48, w: 6.1, h: 0.6,
    fontSize: 9, color: C.WHITE, fontFace: FONT_BODY, lineSpacingMultiple: 1.15
  });

  s3.addNotes(
    "Slide 3 — Problématique.\n" +
    "En Tunisie, la volonté de trier et de protéger l'environnement est bien réelle, en particulier chez les étudiants et les jeunes actifs. Cependant, elle se heurte à 3 obstacles majeurs : " +
    "1) L'opacité territoriale : on ne sait tout simplement pas où déposer ses déchets recyclables. " +
    "2) La confusion technique : beaucoup jettent des emballages non recyclables par méconnaissance des résines plastiques. " +
    "3) Le déficit d'outils numériques : les informations sont éparpillées sur des groupes sociaux éphémères sans cartographie claire."
  );
}

// ==========================================
// SLIDE 4 : NOTRE SOLUTION
// ==========================================
{
  const s4 = pres.addSlide();
  addSlideHeader(s4, '03 • Proposition de Valeur', 'Notre Solution : L’Écosystème EcoTech Digital', 'La synergie vertueuse entre prestations numériques professionnelles et impact citoyen');

  // Boîte Gauche : EcoTech Digital
  s4.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.55, w: 3.6, h: 3.5, rectRadius: 0.1,
    fill: { color: C.CARD_BG }, line: { color: C.EMERALD, width: 1.5 }
  });
  s4.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 1.75, w: 2.0, h: 0.3, rectRadius: 0.06,
    fill: { color: C.LIGHT_MINT }
  });
  s4.addText('PÔLE SERVICES • AGENCE', {
    x: 0.8, y: 1.75, w: 2.0, h: 0.3,
    fontSize: 7.5, bold: true, color: C.FOREST_GREEN, align: 'center', valign: 'middle'
  });
  s4.addText('EcoTech Digital', {
    x: 0.8, y: 2.15, w: 3.2, h: 0.4,
    fontSize: 16, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
  });
  s4.addText('Prestations B2B & Conseil Éco-Numérique', {
    x: 0.8, y: 2.55, w: 3.2, h: 0.25,
    fontSize: 9, bold: true, color: C.MUTED_TEXT, fontFace: FONT_BODY
  });

  const agencyServices = [
    { t: 'EcoWeb', d: 'Conception de sites vitrines et applications web éco-conçus (code optimisé, faible consommation).' },
    { t: 'EcoMarketing', d: 'Stratégie de communication digitale éthique et community management vert.' },
    { t: 'EcoAwareness', d: 'Création de campagnes et ateliers de sensibilisation pour écoles et entreprises.' }
  ];
  let agY = 2.85;
  agencyServices.forEach(srv => {
    s4.addText(`✔ ${srv.t} : `, {
      x: 0.8, y: agY, w: 1.1, h: 0.3,
      fontSize: 8.5, bold: true, color: C.EMERALD, fontFace: FONT_TITLE
    });
    s4.addText(srv.d, {
      x: 1.9, y: agY, w: 2.1, h: 0.45,
      fontSize: 8, color: C.DARK_TEXT, fontFace: FONT_BODY
    });
    agY += 0.52;
  });

  // Connecteur Central (Synergie)
  s4.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 4.35, y: 2.2, w: 1.3, h: 2.2, rectRadius: 0.08,
    fill: { color: C.BEIGE }, line: { color: C.ACCENT_GOLD, width: 1 }
  });
  s4.addText('SYNERGIE\nBIDIRECTIONNELLE', {
    x: 4.4, y: 2.3, w: 1.2, h: 0.5,
    fontSize: 7.5, bold: true, color: '92400E', align: 'center'
  });
  s4.addText('➡ L’agence développe et finance la plateforme', {
    x: 4.4, y: 2.85, w: 1.2, h: 0.65,
    fontSize: 7, color: C.DARK_TEXT, align: 'center'
  });
  s4.addText('⬅ La plateforme apporte visibilité et crédibilité RSE', {
    x: 4.4, y: 3.55, w: 1.2, h: 0.65,
    fontSize: 7, color: C.DARK_TEXT, align: 'center'
  });

  // Boîte Droite : EcoConnect
  s4.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.8, y: 1.55, w: 3.6, h: 3.5, rectRadius: 0.1,
    fill: { color: C.CARD_BG }, line: { color: C.MINT, width: 1.5 }
  });
  s4.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 6.0, y: 1.75, w: 2.0, h: 0.3, rectRadius: 0.06,
    fill: { color: 'E0F2FE' }
  });
  s4.addText('PRODUIT PHARE • CITOYEN', {
    x: 6.0, y: 1.75, w: 2.0, h: 0.3,
    fontSize: 7.5, bold: true, color: '0369A1', align: 'center', valign: 'middle'
  });
  s4.addText('EcoConnect', {
    x: 6.0, y: 2.15, w: 3.2, h: 0.4,
    fontSize: 16, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
  });
  s4.addText('Plateforme Civique Gratuite de Tri & Recyclage', {
    x: 6.0, y: 2.55, w: 3.2, h: 0.25,
    fontSize: 9, bold: true, color: C.MUTED_TEXT, fontFace: FONT_BODY
  });

  const productFeatures = [
    { t: 'Localisation', d: 'Carte interactive des points d’apport volontaire et partenaires de collecte.' },
    { t: 'Guide du Tri', d: 'Répertoire détaillé des matières recyclables et consignes pratiques.' },
    { t: 'Éco-Gestes', d: 'Conseils quotidiens pour réduire les déchets à la source.' }
  ];
  let prY = 2.85;
  productFeatures.forEach(ft => {
    s4.addText(`★ ${ft.t} : `, {
      x: 6.0, y: prY, w: 1.1, h: 0.3,
      fontSize: 8.5, bold: true, color: C.MINT, fontFace: FONT_TITLE
    });
    s4.addText(ft.d, {
      x: 7.1, y: prY, w: 2.1, h: 0.45,
      fontSize: 8, color: C.DARK_TEXT, fontFace: FONT_BODY
    });
    prY += 0.52;
  });

  s4.addNotes(
    "Slide 4 — Notre solution.\n" +
    "Pour répondre à cette problématique, nous avons imaginé une structure hybride novatrice. D'un côté, EcoTech Digital génère des revenus grâce à des prestations de développement web éco-conçu et de marketing responsable pour les entreprises tunisiennes. " +
    "De l'autre, elle réinvestit son savoir-faire technique dans EcoConnect, un portail citoyen gratuit qui guide concrètement les Tunisiens vers le recyclage. " +
    "C'est une synergie parfaite : le produit valide l'expertise de l'agence, et l'agence pérennise le produit sans dépendre uniquement de subventions."
  );
}

// ==========================================
// SLIDE 5 : PRÉSENTATION D'ECOCONNECT
// ==========================================
{
  const s5 = pres.addSlide();
  addSlideHeader(s5, '04 • Innovation Produit', 'Présentation de la Plateforme EcoConnect', 'Le compagnon digital interactif pour guider le tri et géolocaliser les points de recyclage');

  // Mockup Mobile Simulé (Gauche)
  s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.55, w: 2.8, h: 3.5, rectRadius: 0.2,
    fill: { color: '0F172A' }, line: { color: '334155', width: 2 }
  });
  // Écran du Smartphone
  s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.72, y: 1.7, w: 2.56, h: 3.2, rectRadius: 0.12,
    fill: { color: C.WHITE }
  });
  // Header App
  s5.addShape(pres.shapes.RECTANGLE, {
    x: 0.72, y: 1.7, w: 2.56, h: 0.45,
    fill: { color: C.FOREST_GREEN }
  });
  s5.addText('EcoConnect • Tunis', {
    x: 0.8, y: 1.75, w: 2.0, h: 0.35,
    fontSize: 9, bold: true, color: C.WHITE, fontFace: FONT_TITLE
  });
  // Barre de recherche dans l'app
  s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.82, y: 2.25, w: 2.36, h: 0.3, rectRadius: 0.05,
    fill: { color: 'F1F5F9' }, line: { color: 'CBD5E1', width: 0.8 }
  });
  s5.addText('🔍 Rechercher un déchet ou point...', {
    x: 0.9, y: 2.27, w: 2.2, h: 0.26,
    fontSize: 7.5, color: C.MUTED_TEXT, fontFace: FONT_BODY
  });
  // Badges Matières
  s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.82, y: 2.65, w: 0.7, h: 0.24, rectRadius: 0.04,
    fill: { color: C.LIGHT_MINT }
  });
  s5.addText('Plastique', { x: 0.82, y: 2.65, w: 0.7, h: 0.24, fontSize: 7, bold: true, color: C.FOREST_GREEN, align: 'center', valign: 'middle' });
  
  s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 1.58, y: 2.65, w: 0.5, h: 0.24, rectRadius: 0.04,
    fill: { color: 'E0F2FE' }
  });
  s5.addText('Verre', { x: 1.58, y: 2.65, w: 0.5, h: 0.24, fontSize: 7, bold: true, color: '0369A1', align: 'center', valign: 'middle' });

  s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 2.14, y: 2.65, w: 0.55, h: 0.24, rectRadius: 0.04,
    fill: { color: 'FEF3C7' }
  });
  s5.addText('Carton', { x: 2.14, y: 2.65, w: 0.55, h: 0.24, fontSize: 7, bold: true, color: '92400E', align: 'center', valign: 'middle' });

  // Carte résultat simulée
  s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.82, y: 3.0, w: 2.36, h: 0.85, rectRadius: 0.06,
    fill: { color: 'F8FAFC' }, line: { color: 'E2E8F0', width: 0.8 }
  });
  s5.addText('📍 Point Collecte : Borj Cédria Eco', {
    x: 0.9, y: 3.05, w: 2.2, h: 0.2, fontSize: 7.5, bold: true, color: C.FOREST_GREEN
  });
  s5.addText('Bouteilles PET, papiers, bouchons\nOuvert 7j/7 • À 1.2 km de votre position', {
    x: 0.9, y: 3.28, w: 2.2, h: 0.5, fontSize: 6.8, color: C.MUTED_TEXT, lineSpacingMultiple: 1.15
  });

  // Bouton CTA dans le mock
  s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.82, y: 4.0, w: 2.36, h: 0.32, rectRadius: 0.06,
    fill: { color: C.EMERALD }
  });
  s5.addText('Itinéraire & Consignes de dépôt', {
    x: 0.82, y: 4.0, w: 2.36, h: 0.32, fontSize: 7.5, bold: true, color: C.WHITE, align: 'center', valign: 'middle'
  });

  // Parcours Utilisateur en 4 Étapes (Droite)
  s5.addText('PARCOURS UTILISATEUR OPTIMISÉ (USER JOURNEY)', {
    x: 3.65, y: 1.55, w: 5.75, h: 0.3,
    fontSize: 10, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
  });

  const steps = [
    { n: '1', title: 'Localiser', desc: 'Saisie de la municipalité ou détection GPS pour afficher les points d’apport volontaire.' },
    { n: '2', title: 'Identifier', desc: 'Sélection de la matière (plastique, verre, métal) avec filtre visuel des consignes d’acceptabilité.' },
    { n: '3', title: 'Préparer', desc: 'Fiche pédagogique immédiate : rincer les contenants, aplatir les cartons, trier les bouchons.' },
    { n: '4', title: 'Revaloriser', desc: 'Dépôt effectif au point de collecte et estimation de l’impact CO2 économisé.' }
  ];

  steps.forEach((st, idx) => {
    const yPos = 1.95 + idx * 0.75;
    s5.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 3.65, y: yPos, w: 5.75, h: 0.65, rectRadius: 0.08,
      fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
    });
    // Cercle numéro
    s5.addShape(pres.shapes.OVAL, {
      x: 3.8, y: yPos + 0.12, w: 0.4, h: 0.4,
      fill: { color: C.FOREST_GREEN }
    });
    s5.addText(st.n, {
      x: 3.8, y: yPos + 0.12, w: 0.4, h: 0.4,
      fontSize: 9, bold: true, color: C.WHITE, align: 'center', valign: 'middle'
    });
    // Titre et description
    s5.addText(st.title, {
      x: 4.35, y: yPos + 0.08, w: 4.9, h: 0.24,
      fontSize: 9.5, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
    });
    s5.addText(st.desc, {
      x: 4.35, y: yPos + 0.32, w: 4.9, h: 0.28,
      fontSize: 8.5, color: C.DARK_TEXT, fontFace: FONT_BODY
    });
  });

  s5.addNotes(
    "Slide 5 — Présentation d'EcoConnect.\n" +
    "EcoConnect est l'application au cœur de notre démarche citoyenne. Comme vous le voyez sur ce mockup, l'ergonomie est épurée, responsive et conçue selon l'approche Mobile First. " +
    "Le parcours utilisateur est réduit à 4 étapes clés : 1. Localiser son quartier, 2. Identifier le matériau, 3. Préparer son déchet avec nos fiches pédagogiques, et 4. Déposer au point de collecte. " +
    "Notre objectif : lever tous les freins cognitifs au tri en Tunisie en offrant une réponse en moins de 10 secondes."
  );
}

// ==========================================
// SLIDE 6 : NOS SERVICES
// ==========================================
{
  const s6 = pres.addSlide();
  addSlideHeader(s6, '05 • Offre Commerciale', 'Nos Services : Une Expertise Multidimensionnelle', 'Quatre pôles d’activité complémentaires alliant ingénierie logicielle et communication responsable');

  const services = [
    {
      code: '01',
      title: 'EcoConnect',
      tag: 'PORTAIL CITOYEN',
      desc: 'Plateforme numérique d’information et de localisation des filières de tri en libre accès.',
      deliverables: [
        'Répertoire interactif des matériaux recyclables',
        'Fiches de consignes de tri contextualisées',
        'Accès gratuit et universel sur mobile et PC'
      ],
      color: C.FOREST_GREEN
    },
    {
      code: '02',
      title: 'EcoWeb',
      tag: 'DÉVELOPPEMENT WEB',
      desc: 'Conception et refonte de sites internet éco-conçus à très faible empreinte carbone.',
      deliverables: [
        'Architecture Green Code & allègement des requêtes',
        'Accessibilité numérique conforme WCAG',
        'Audit de performance et optimisation SEO propre'
      ],
      color: C.EMERALD
    },
    {
      code: '03',
      title: 'EcoMarketing',
      tag: 'COMMUNICATION DIGITALE',
      desc: 'Gestion éthique des réseaux sociaux et stratégies de visibilité sans artifice.',
      deliverables: [
        'Community management responsable',
        'Création de carrousels & infographies pédagogiques',
        'Campagnes de storytelling d’impact prouvé'
      ],
      color: '0D9488'
    },
    {
      code: '04',
      title: 'EcoAwareness',
      tag: 'SENSIBILISATION RSE',
      desc: 'Programmes de sensibilisation environnementale clés en main pour organisations.',
      deliverables: [
        'Ateliers d’acculturation à la sobriété numérique',
        'Kits digitaux pour écoles et universités',
        'Accompagnement de projets RSE d’entreprises'
      ],
      color: 'B45309'
    }
  ];

  services.forEach((srv, i) => {
    const xPos = 0.6 + i * 2.25;
    s6.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: xPos, y: 1.55, w: 2.15, h: 3.55, rectRadius: 0.1,
      fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
    });

    // Bandeau couleur en haut de carte
    s6.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: xPos, y: 1.55, w: 2.15, h: 0.12, rectRadius: 0.05,
      fill: { color: srv.color }, line: { color: srv.color }
    });

    s6.addText(srv.tag, {
      x: xPos + 0.15, y: 1.75, w: 1.85, h: 0.22,
      fontSize: 7, bold: true, color: C.MUTED_TEXT, fontFace: FONT_TITLE
    });

    s6.addText(srv.title, {
      x: xPos + 0.15, y: 1.98, w: 1.85, h: 0.35,
      fontSize: 13, bold: true, color: srv.color, fontFace: FONT_TITLE
    });

    s6.addText(srv.desc, {
      x: xPos + 0.15, y: 2.38, w: 1.85, h: 0.7,
      fontSize: 8, color: C.DARK_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.15
    });

    s6.addShape(pres.shapes.LINE, {
      x: xPos + 0.15, y: 3.12, w: 1.85, h: 0,
      line: { color: C.CARD_BORDER, width: 0.8 }
    });

    s6.addText('LIVRABLES CLÉS :', {
      x: xPos + 0.15, y: 3.22, w: 1.85, h: 0.2,
      fontSize: 7, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
    });

    let lY = 3.45;
    srv.deliverables.forEach(d => {
      s6.addText('• ' + d, {
        x: xPos + 0.15, y: lY, w: 1.85, h: 0.48,
        fontSize: 7.5, color: C.MUTED_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.1
      });
      lY += 0.5;
    });
  });

  s6.addNotes(
    "Slide 6 — Nos services.\n" +
    "Pour pérenniser notre structure, nous articulons notre proposition autour de 4 offres complémentaires : " +
    "EcoConnect pour la dimension citoyenne ; EcoWeb pour concevoir des sites internet professionnels allégés et éco-responsables ; EcoMarketing pour animer les réseaux sociaux de manière éthique et pédagogique ; et EcoAwareness pour intervenir directement dans les campus et entreprises avec des ateliers RSE concrets."
  );
}

// ==========================================
// SLIDE 7 : PUBLIC CIBLE
// ==========================================
{
  const s7 = pres.addSlide();
  addSlideHeader(s7, '06 • Marché & Utilisateurs', 'Publics Cibles & Analyse des Besoins', 'Une segmentation réfléchie pour répondre aux attentes spécifiques de chaque acteur sociétal');

  const targets = [
    {
      segment: 'Citoyens & Étudiants',
      role: 'Utilisateurs finaux d’EcoConnect',
      needs: 'Recherche de simplicité, d’informations pratiques et rapides pour trier sans contrainte au quotidien.',
      solution: 'Application web responsive, accès libre, fiches de tri visuelles et astuces zéro déchet concrètes.',
      color: C.FOREST_GREEN
    },
    {
      segment: 'Écoles & Universités',
      role: 'Partenaires éducatifs & Campus',
      needs: 'Sensibilisation des jeunes générations et valorisation de leurs démarches éco-campus.',
      solution: 'Ateliers EcoAwareness, challenges inter-universités et interventions de sensibilisation numérique.',
      color: C.EMERALD
    },
    {
      segment: 'Associations Écologiques',
      role: 'Acteurs de terrain & Relais',
      needs: 'Manque cruel de visibilité numérique et d’outils web pour promouvoir leurs journées de collecte.',
      solution: 'Mise en avant prioritaire de leurs collectes sur EcoConnect et accompagnement en communication digitale.',
      color: '0369A1'
    },
    {
      segment: 'Entreprises & PME Locales',
      role: 'Clients des prestations de services',
      needs: 'Valorisation de leur responsabilité sociétale (RSE) et modernisation de leur vitrine web.',
      solution: 'Création de sites éco-conçus (EcoWeb) et community management vert (EcoMarketing).',
      color: 'B45309'
    }
  ];

  targets.forEach((t, i) => {
    const xPos = 0.6 + i * 2.25;
    s7.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: xPos, y: 1.55, w: 2.15, h: 3.55, rectRadius: 0.1,
      fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
    });

    s7.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: xPos, y: 1.55, w: 2.15, h: 0.35, rectRadius: 0.05,
      fill: { color: t.color }
    });
    s7.addText(t.segment, {
      x: xPos + 0.1, y: 1.55, w: 1.95, h: 0.35,
      fontSize: 8.5, bold: true, color: C.WHITE, fontFace: FONT_TITLE, align: 'center', valign: 'middle'
    });

    s7.addText(t.role, {
      x: xPos + 0.15, y: 2.0, w: 1.85, h: 0.28,
      fontSize: 7.5, bold: true, color: C.MUTED_TEXT, fontFace: FONT_TITLE
    });

    s7.addText('BESOINS IDENTIFIÉS :', {
      x: xPos + 0.15, y: 2.35, w: 1.85, h: 0.2,
      fontSize: 7, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
    });
    s7.addText(t.needs, {
      x: xPos + 0.15, y: 2.58, w: 1.85, h: 0.95,
      fontSize: 8, color: C.DARK_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.15
    });

    s7.addShape(pres.shapes.LINE, {
      x: xPos + 0.15, y: 3.58, w: 1.85, h: 0,
      line: { color: C.CARD_BORDER, width: 0.8 }
    });

    s7.addText('RÉPONSE ÉCOTECH :', {
      x: xPos + 0.15, y: 3.7, w: 1.85, h: 0.2,
      fontSize: 7, bold: true, color: C.MINT, fontFace: FONT_TITLE
    });
    s7.addText(t.solution, {
      x: xPos + 0.15, y: 3.92, w: 1.85, h: 1.05,
      fontSize: 8, color: C.MUTED_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.15
    });
  });

  s7.addNotes(
    "Slide 7 — Public cible.\n" +
    "Une bonne démarche entrepreneuriale repose sur une segmentation précise. " +
    "Nous adressons d'abord les citoyens et étudiants qui sont les utilisateurs finaux en demande de simplicité. Ensuite, les campus et associations avec qui nous formons des synergies de terrain. " +
    "Enfin, les entreprises et PME tunisiennes qui constituent notre clientèle B2B payante pour financer l'écosystème à travers nos prestations de développement web et de marketing d'impact."
  );
}

// ==========================================
// SLIDE 8 : MODÈLE ÉCONOMIQUE
// ==========================================
{
  const s8 = pres.addSlide();
  addSlideHeader(s8, '07 • Viabilité Financière', 'Modèle Économique Réaliste & Progressif', 'Une monétisation B2B responsable pour pérenniser une plateforme citoyenne d’utilité publique');

  // Encart Rigueur Académique
  s8.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.55, w: 8.8, h: 0.45, rectRadius: 0.06,
    fill: { color: C.BEIGE }, line: { color: C.ACCENT_GOLD, width: 0.8 }
  });
  s8.addText('TRANSPARENCE ACADÉMIQUE : Projet en phase de conception universitaire. Les sources de revenus présentées ci-dessous constituent des flux modélisés pour l’évaluation de la viabilité économique, sans déclaration de chiffre d’affaires antérieur.', {
    x: 0.8, y: 1.58, w: 8.4, h: 0.38,
    fontSize: 7.5, bold: true, color: '78350F', fontFace: FONT_BODY, align: 'center', valign: 'middle'
  });

  // 3 Blocs de Flux de Revenus B2B
  const streams = [
    {
      title: 'Prestations Web (EcoWeb)',
      type: 'REVENUS B2B PONCTUELS',
      desc: 'Création et refonte de sites web vitrines éco-conçus pour entreprises et professionnels soucieux de leur empreinte carbone.',
      kpis: ['Paniers moyens selon l’envergure du site', 'Audit carbone web & optimisation SEO', 'Contrats de maintenance annuelle']
    },
    {
      title: 'Community Management (EcoMarketing)',
      type: 'REVENUS B2B RÉCURRENTS',
      desc: 'Abonnements mensuels de gestion de présence digitale responsable pour PME et marques engagées en Tunisie.',
      kpis: ['Forfaits mensuels modulables', 'Production de contenus & infographies', 'Storytelling RSE & gestion de communauté']
    },
    {
      title: 'Ateliers & Partenariats (EcoAwareness)',
      type: 'REVENUS ÉVÉNEMENTIELS & B2G',
      desc: 'Animation de séminaires de sensibilisation environnementale pour campus, écoles privées et comités RSE d’entreprises.',
      kpis: ['Facturation à la session ou au module', 'Kits pédagogiques sous licence', 'Perspectives d’appels à projets publics']
    }
  ];

  streams.forEach((st, i) => {
    const xPos = 0.6 + i * 2.98;
    s8.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: xPos, y: 2.15, w: 2.85, h: 2.9, rectRadius: 0.1,
      fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
    });

    s8.addText(st.type, {
      x: xPos + 0.15, y: 2.3, w: 2.55, h: 0.2,
      fontSize: 7, bold: true, color: C.MINT, fontFace: FONT_TITLE
    });

    s8.addText(st.title, {
      x: xPos + 0.15, y: 2.52, w: 2.55, h: 0.45,
      fontSize: 11, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
    });

    s8.addText(st.desc, {
      x: xPos + 0.15, y: 3.0, w: 2.55, h: 0.75,
      fontSize: 8, color: C.DARK_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.15
    });

    s8.addShape(pres.shapes.LINE, {
      x: xPos + 0.15, y: 3.8, w: 2.55, h: 0,
      line: { color: C.CARD_BORDER, width: 0.8 }
    });

    s8.addText('MODALITÉS ENVISAGÉES :', {
      x: xPos + 0.15, y: 3.92, w: 2.55, h: 0.2,
      fontSize: 7, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
    });

    let kY = 4.15;
    st.kpis.forEach(k => {
      s8.addText('• ' + k, {
        x: xPos + 0.15, y: kY, w: 2.55, h: 0.28,
        fontSize: 7.5, color: C.MUTED_TEXT, fontFace: FONT_BODY
      });
      kY += 0.28;
    });
  });

  s8.addNotes(
    "Slide 8 — Modèle économique.\n" +
    "Pour garantir la crédibilité de ce projet devant un jury universitaire, nous adoptons une posture de transparence rigoureuse : il s'agit d'un modèle économique projeté. " +
    "Notre structure de rentabilité ne fait pas payer les citoyens pour recycler. Elle repose sur des prestations de services B2B réelles : facturation de sites web éco-conçus (EcoWeb), abonnements mensuels de community management (EcoMarketing), et interventions d'ateliers RSE en entreprises (EcoAwareness). " +
    "Ce modèle permet d'assurer l'équilibre financier et l'hébergement gratuit d'EcoConnect."
  );
}

// ==========================================
// SLIDE 9 : STRATÉGIE MARKETING DIGITAL
// ==========================================
{
  const s9 = pres.addSlide();
  addSlideHeader(s9, '08 • Communication & Croissance', 'Stratégie de Marketing Digital & Réseaux Sociaux', 'Mobiliser les communautés tunisiennes sur Facebook et Instagram par la pédagogie positive');

  // Colonne Gauche : Piliers et Exemples Concrets
  s9.addText('PILIERS DE CONTENU ÉTHIQUE (INSTAGRAM & FACEBOOK)', {
    x: 0.6, y: 1.55, w: 4.8, h: 0.25,
    fontSize: 9, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
  });

  const contentPillars = [
    {
      format: 'Carrousels Pédagogiques (Instagram)',
      theme: '« 5 erreurs fréquentes lors du tri du plastique en Tunisie »',
      impact: 'Démystifier les types de plastique (PET vs PVC) avec visuels clairs.'
    },
    {
      format: 'Reels & Vidéos Courtes (30s)',
      theme: '« Comment trouver un bac de collecte près de chez soi en 3 clics »',
      impact: 'Démonstrations concrètes sur smartphone pour susciter l’adoption.'
    },
    {
      format: 'Infographies & Astuces Zéro Déchet',
      theme: '« Les alternatives durables aux sacs plastiques à usage unique »',
      impact: 'Valorisation des gestes simples et des savoir-faire artisanaux locaux.'
    },
    {
      format: 'Sondages & Quiz Interactifs en Story',
      theme: '« Vrai ou Faux : les tickets de caisse sont-ils recyclables ? »',
      impact: 'Engagement communautaire actif et apprentissage par le jeu.'
    }
  ];

  let pY = 1.85;
  contentPillars.forEach(cp => {
    s9.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.6, y: pY, w: 4.8, h: 0.72, rectRadius: 0.06,
      fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
    });
    s9.addText(cp.format, {
      x: 0.75, y: pY + 0.06, w: 4.5, h: 0.2,
      fontSize: 8, bold: true, color: C.EMERALD, fontFace: FONT_TITLE
    });
    s9.addText(cp.theme, {
      x: 0.75, y: pY + 0.26, w: 4.5, h: 0.22,
      fontSize: 8.5, bold: true, color: C.FOREST_GREEN, fontFace: FONT_BODY
    });
    s9.addText(cp.impact, {
      x: 0.75, y: pY + 0.48, w: 4.5, h: 0.2,
      fontSize: 7.5, color: C.MUTED_TEXT, fontFace: FONT_BODY
    });
    pY += 0.8;
  });

  // Colonne Droite : Calendrier Éditorial Type
  s9.addText('CALENDRIER ÉDITORIAL HEBDOMADAIRE TYPE', {
    x: 5.65, y: 1.55, w: 3.75, h: 0.25,
    fontSize: 9, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
  });

  const calendar = [
    { day: 'Lundi', type: 'Le Chiffre Vert', topic: 'Statistique ou réalité écologique vulgarisée sur la Tunisie.' },
    { day: 'Mercredi', type: 'Focus Matière', topic: 'Guide pratique d’un matériau recyclable sur EcoConnect.' },
    { day: 'Vendredi', type: 'Éco-Geste Local', topic: 'Conseil pratique zéro déchet pour le week-end.' },
    { day: 'Dimanche', type: 'Quiz & Défi', topic: 'Sondage interactif en story et réponse commentée.' }
  ];

  calendar.forEach((cal, i) => {
    const cY = 1.85 + i * 0.58;
    s9.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.65, y: cY, w: 3.75, h: 0.52, rectRadius: 0.06,
      fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
    });
    s9.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.75, y: cY + 0.1, w: 0.9, h: 0.32, rectRadius: 0.04,
      fill: { color: C.LIGHT_MINT }
    });
    s9.addText(cal.day, {
      x: 5.75, y: cY + 0.1, w: 0.9, h: 0.32,
      fontSize: 7.5, bold: true, color: C.FOREST_GREEN, align: 'center', valign: 'middle'
    });
    s9.addText(cal.type, {
      x: 6.75, y: cY + 0.08, w: 2.55, h: 0.18,
      fontSize: 8, bold: true, color: C.DARK_TEXT, fontFace: FONT_TITLE
    });
    s9.addText(cal.topic, {
      x: 6.75, y: cY + 0.28, w: 2.55, h: 0.22,
      fontSize: 7.5, color: C.MUTED_TEXT, fontFace: FONT_BODY
    });
  });

  // Hashtags recommandés
  s9.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.65, y: 4.3, w: 3.75, h: 0.75, rectRadius: 0.06,
    fill: { color: C.BEIGE }, line: { color: C.CARD_BORDER, width: 0.8 }
  });
  s9.addText('HASHTAGS STRATÉGIQUES ÉCO-RESPONSABLES :', {
    x: 5.8, y: 4.36, w: 3.45, h: 0.2,
    fontSize: 7, bold: true, color: '92400E', fontFace: FONT_TITLE
  });
  s9.addText('#EcoTechTunisie  #EcoConnect  #RecyclageTunisie\n#NumeriqueResponsable  #ZeroDechetTunisie  #GreenTech', {
    x: 5.8, y: 4.58, w: 3.45, h: 0.42,
    fontSize: 7.5, color: C.FOREST_GREEN, fontFace: FONT_BODY, lineSpacingMultiple: 1.15
  });

  s9.addNotes(
    "Slide 9 — Stratégie marketing digital.\n" +
    "Pour sensibiliser efficacement en Tunisie, nous ciblons en priorité Instagram et Facebook avec une approche éditoriale bienveillante, axée sur la valorisation et le jeu, sans discours anxiogène. " +
    "Notre ligne éditoriale s'articule autour d'un calendrier régulier à 4 rendez-vous par semaine : le chiffre clé le lundi, le focus technique le mercredi, l'éco-geste le vendredi, et le quiz communautaire le dimanche. " +
    "Cette régularité permet de fidéliser une communauté engagée autour de la marque."
  );
}

// ==========================================
// SLIDE 10 : IDENTITÉ VISUELLE & COMMUNICATION
// ==========================================
{
  const s10 = pres.addSlide();
  addSlideHeader(s10, '09 • Charte Graphique', 'Identité Visuelle & Communication de Marque', 'Une esthétique moderne, épurée et technologique, incarnant l’harmonie entre nature et digital');

  // Colonne Gauche : Concepts de Logos
  s10.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.55, w: 4.25, h: 1.65, rectRadius: 0.1,
    fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
  });
  s10.addText('CONCEPT LOGO : EcoTech Digital', {
    x: 0.8, y: 1.7, w: 3.85, h: 0.25,
    fontSize: 9, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
  });
  s10.addText('L’emblème fusionne les circuits imprimés d’une puce informatique avec les nervures d’une feuille végétale. Il matérialise visuellement la symbiose entre la puissance de calcul technologique et la régénération biologique.', {
    x: 0.8, y: 2.0, w: 3.85, h: 0.6,
    fontSize: 8, color: C.DARK_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.15
  });
  s10.addText('Slogan associé : « Le numérique au service de la planète. »', {
    x: 0.8, y: 2.7, w: 3.85, h: 0.35,
    fontSize: 8, italic: true, bold: true, color: C.EMERALD, fontFace: FONT_BODY
  });

  s10.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 3.35, w: 4.25, h: 1.7, rectRadius: 0.1,
    fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
  });
  s10.addText('CONCEPT LOGO : EcoConnect (Produit)', {
    x: 0.8, y: 3.5, w: 3.85, h: 0.25,
    fontSize: 9, bold: true, color: C.MINT, fontFace: FONT_TITLE
  });
  s10.addText('Réinterprétation contemporaine du ruban de Möbius (symbole universel du recyclage) sous forme de nœuds de connexion réseau et de point de géolocalisation. Évoque la boucle infinie de valorisation des matières.', {
    x: 0.8, y: 3.8, w: 3.85, h: 0.6,
    fontSize: 8, color: C.DARK_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.15
  });
  s10.addText('Identité produit distincte mais harmonisée avec la maison mère.', {
    x: 0.8, y: 4.5, w: 3.85, h: 0.35,
    fontSize: 8, bold: true, color: C.MUTED_TEXT, fontFace: FONT_BODY
  });

  // Colonne Droite : Palette Chromatique & Typographie
  s10.addText('PALETTE CHROMATIQUE OFFICIELLE', {
    x: 5.15, y: 1.55, w: 4.25, h: 0.25,
    fontSize: 9, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
  });

  const swatches = [
    { name: 'Vert Forêt Profond', hex: '#103B2B', role: 'Solidité, crédibilité', c: '103B2B', text: 'FFFFFF' },
    { name: 'Vert Éco Dynamique', hex: '#2D6A4F', role: 'Action, innovation', c: '2D6A4F', text: 'FFFFFF' },
    { name: 'Vert Menthe Accent', hex: '#52B788', role: 'Fraîcheur, repères', c: '52B788', text: '1E293B' },
    { name: 'Sable Méditerranéen', hex: '#F4F1DE', role: 'Chaleur naturelle', c: 'F4F1DE', text: '1E293B' }
  ];

  swatches.forEach((sw, i) => {
    const sY = 1.85 + i * 0.48;
    s10.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.15, y: sY, w: 1.2, h: 0.4, rectRadius: 0.05,
      fill: { color: sw.c }
    });
    s10.addText(sw.hex, {
      x: 5.15, y: sY, w: 1.2, h: 0.4,
      fontSize: 8, bold: true, color: sw.text, align: 'center', valign: 'middle'
    });
    s10.addText(`${sw.name} — ${sw.role}`, {
      x: 6.45, y: sY, w: 2.95, h: 0.4,
      fontSize: 8.5, color: C.DARK_TEXT, fontFace: FONT_BODY, valign: 'middle'
    });
  });

  // Bloc Typographie & Ton
  s10.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.15, y: 3.9, w: 4.25, h: 1.15, rectRadius: 0.08,
    fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
  });
  s10.addText('TYPOGRAPHIE & TON DE VOIX', {
    x: 5.3, y: 4.0, w: 3.95, h: 0.22,
    fontSize: 8, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
  });
  s10.addText('• Typographie Sans-Serif géométrique : modernité, haute lisibilité écran.\n• Ton de voix : scientifique mais chaleureux, constructif, pédagogique et valorisant.', {
    x: 5.3, y: 4.25, w: 3.95, h: 0.65,
    fontSize: 8, color: C.MUTED_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.2
  });

  s10.addNotes(
    "Slide 10 — Identité visuelle et communication.\n" +
    "L'identité graphique d'EcoTech Digital a été pensée pour exprimer une convergence parfaite entre nature et technologie. " +
    "Le vert forêt profond apporte le sérieux institutionnel nécessaire, tandis que le vert menthe et le sable apportent dynamisme et ancrage méditerranéen. " +
    "Le logo de la marque mère stylise les nervures végétales sous forme de circuits imprimés, tandis qu'EcoConnect adopte le ruban de Möbius du recyclage revisité en nœuds de réseau. " +
    "Cette cohérence visuelle renforce la crédibilité professionnelle de notre projet."
  );
}

// ==========================================
// SLIDE 11 : DIMENSION INFORMATIQUE & COMPÉTENCES
// ==========================================
{
  const s11 = pres.addSlide();
  addSlideHeader(s11, '10 • Rigueur Technique', 'Dimension Informatique & Compétences Mobilisées', 'Une architecture logicielle progressive, rigoureusement documentée et éco-conçue');

  // Colonne Gauche : Prototype Actuel vs Roadmap Future
  s11.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 1.55, w: 4.25, h: 1.65, rectRadius: 0.1,
    fill: { color: C.CARD_BG }, line: { color: C.EMERALD, width: 1.2 }
  });
  s11.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 1.7, w: 1.8, h: 0.26, rectRadius: 0.05,
    fill: { color: C.LIGHT_MINT }
  });
  s11.addText('SOCLE ACTUEL DU PROTOTYPE', {
    x: 0.8, y: 1.7, w: 1.8, h: 0.26,
    fontSize: 7.5, bold: true, color: C.FOREST_GREEN, align: 'center', valign: 'middle'
  });
  s11.addText('• Front-End : HTML5 sémantique, CSS3 moderne (Flexbox, CSS Grid).\n• Interactivité : JavaScript Vanilla modulaire sans frameworks lourds.\n• Ergonomie : Interface 100% responsive et respect strict des contrastes WCAG.\n• Données : Structures JSON légères pour simuler les points de collecte.', {
    x: 0.8, y: 2.05, w: 3.85, h: 1.0,
    fontSize: 8, color: C.DARK_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.2
  });

  s11.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.6, y: 3.35, w: 4.25, h: 1.7, rectRadius: 0.1,
    fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
  });
  s11.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 3.5, w: 2.1, h: 0.26, rectRadius: 0.05,
    fill: { color: 'E0F2FE' }
  });
  s11.addText('FEUILLE DE ROUTE TECHNIQUE (FUTURE)', {
    x: 0.8, y: 3.5, w: 2.1, h: 0.26,
    fontSize: 7.5, bold: true, color: '0369A1', align: 'center', valign: 'middle'
  });
  s11.addText('• Back-End prévu : Architecture REST avec Node.js & Express.\n• Base de Données relationnelle : MySQL pour le référencement des points réels.\n• Cartographie : Intégration OpenStreetMap & Leaflet.js sans traceurs intrusifs.\n• Sécurité & Confidentialité : Conformité stricte INPDP (Tunisie) et RGPD.', {
    x: 0.8, y: 3.85, w: 3.85, h: 1.05,
    fontSize: 8, color: C.DARK_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.2
  });

  // Colonne Droite : Marketing de Soi & Compétences Développées
  s11.addText('COMPÉTENCES DÉMONTRÉES (MARKETING DE SOI)', {
    x: 5.15, y: 1.55, w: 4.25, h: 0.25,
    fontSize: 9, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
  });

  const skills = [
    { cat: 'Développement Web & Éco-Conception', det: 'Maîtrise de l’optimisation front-end, du Green Code et de la réduction de la charge réseau.' },
    { cat: 'UI / UX Design Centré Utilisateur', det: 'Conception de parcours fluides, maquettage responsive et ergonomie d’utilité publique.' },
    { cat: 'Gestion de Projet & Stratégie', det: 'Définition d’une feuille de route réaliste, distinction prototype vs déploiement industriel.' },
    { cat: 'Communication Digitale & RSE', det: 'Valorisation d’un engagement environnemental vérifiable, éloigné de toute imposture commerciale.' }
  ];

  skills.forEach((sk, i) => {
    const skY = 1.85 + i * 0.8;
    s11.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.15, y: skY, w: 4.25, h: 0.72, rectRadius: 0.08,
      fill: { color: C.CARD_BG }, line: { color: C.CARD_BORDER, width: 1 }
    });
    s11.addText(`★ ${sk.cat}`, {
      x: 5.3, y: skY + 0.08, w: 3.95, h: 0.22,
      fontSize: 8.5, bold: true, color: C.FOREST_GREEN, fontFace: FONT_TITLE
    });
    s11.addText(sk.det, {
      x: 5.3, y: skY + 0.32, w: 3.95, h: 0.35,
      fontSize: 7.8, color: C.MUTED_TEXT, fontFace: FONT_BODY, lineSpacingMultiple: 1.15
    });
  });

  s11.addNotes(
    "Slide 11 — Dimension informatique et compétences.\n" +
    "Sur le plan technique et scientifique, nous appliquons une déontologie stricte : nous distinguons clairement ce qui est déjà développé dans notre prototype fonctionnel (un front-end complet, responsive et accessible en HTML/CSS/JavaScript), de notre feuille de route future (l'implémentation d'une API Node.js/Express, d'une base de données MySQL et d'une cartographie OpenStreetMap). " +
    "Pour le cours de Marketing de soi, ce projet démontre ma capacité à concevoir une solution logicielle sobre, à respecter les normes d'accessibilité WCAG et la législation sur la protection des données personnelles (INPDP)."
  );
}

// ==========================================
// SLIDE 12 : CONCLUSION & PERSPECTIVES
// ==========================================
{
  const s12 = pres.addSlide();
  s12.background = { color: C.FOREST_GREEN };

  // Motifs d'accentuation
  s12.addShape(pres.shapes.OVAL, {
    x: -1.0, y: 3.5, w: 4.0, h: 4.0,
    fill: { color: C.EMERALD, transparency: 80 }
  });
  s12.addShape(pres.shapes.OVAL, {
    x: 8.0, y: -1.0, w: 4.0, h: 4.0,
    fill: { color: C.MINT, transparency: 85 }
  });

  // Badge Clôture
  s12.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 0.6, w: 3.2, h: 0.3, rectRadius: 0.06,
    fill: { color: '1B4D3E' }, line: { color: C.MINT, width: 1 }
  });
  s12.addText('SYNTHÈSE & PERSPECTIVES', {
    x: 0.8, y: 0.6, w: 3.2, h: 0.3,
    fontSize: 8.5, bold: true, color: C.MINT, align: 'center', valign: 'middle'
  });

  s12.addText('Conclusion & Perspectives d’Avenir', {
    x: 0.8, y: 1.05, w: 8.4, h: 0.5,
    fontSize: 26, bold: true, color: C.WHITE, fontFace: FONT_TITLE
  });

  // 3 Cartes de Perspectives (Horizons)
  const horizons = [
    {
      h: 'HORIZON 1 • COURT TERME',
      title: 'Validation & Tests Campus',
      desc: 'Déploiement du prototype d’EcoConnect auprès d’un panel test d’étudiants et partenariats pilotes avec des clubs écologiques universitaires.'
    },
    {
      h: 'HORIZON 2 • MOYEN TERME',
      title: 'Cartographie & Données Réelles',
      desc: 'Enrichissement d’une base de données vérifiée sur le Grand Tunis avec des associations de valorisation des déchets et des municipalités.'
    },
    {
      h: 'HORIZON 3 • LONG TERME',
      title: 'Évolution PWA & Extension',
      desc: 'Transformation en Progressive Web App (PWA) téléchargeable hors-ligne et élargissement progressif aux principaux gouvernorats de Tunisie.'
    }
  ];

  horizons.forEach((hz, i) => {
    const xPos = 0.8 + i * 2.85;
    s12.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: xPos, y: 1.7, w: 2.7, h: 2.0, rectRadius: 0.08,
      fill: { color: '164E3A' }, line: { color: C.MINT, width: 0.8 }
    });
    s12.addText(hz.h, {
      x: xPos + 0.15, y: 1.85, w: 2.4, h: 0.22,
      fontSize: 7.5, bold: true, color: C.MINT, fontFace: FONT_TITLE
    });
    s12.addText(hz.title, {
      x: xPos + 0.15, y: 2.12, w: 2.4, h: 0.35,
      fontSize: 11, bold: true, color: C.WHITE, fontFace: FONT_TITLE
    });
    s12.addText(hz.desc, {
      x: xPos + 0.15, y: 2.52, w: 2.4, h: 1.0,
      fontSize: 8.5, color: 'E2E8F0', fontFace: FONT_BODY, lineSpacingMultiple: 1.15
    });
  });

  // Slogan Final en Majesté
  s12.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.8, y: 3.95, w: 8.4, h: 0.95, rectRadius: 0.1,
    fill: { color: C.WHITE }
  });
  s12.addText('SLOGAN DE CLÔTURE :', {
    x: 1.0, y: 4.05, w: 8.0, h: 0.22,
    fontSize: 8, bold: true, color: C.MUTED_TEXT, align: 'center', fontFace: FONT_TITLE
  });
  s12.addText('« EcoTech Digital : innovons aujourd’hui pour une planète plus durable. »', {
    x: 1.0, y: 4.3, w: 8.0, h: 0.45,
    fontSize: 14, bold: true, color: C.FOREST_GREEN, align: 'center', fontFace: FONT_TITLE
  });

  s12.addText('Merci de votre attention et de vos questions.', {
    x: 0.8, y: 5.1, w: 8.4, h: 0.3,
    fontSize: 9.5, italic: true, color: C.LIGHT_MINT, align: 'center'
  });

  s12.addNotes(
    "Slide 12 — Conclusion et perspectives.\n" +
    "Pour conclure, EcoTech Digital démontre qu'il est possible de concilier utilité sociale, rigueur scientifique et viabilité professionnelle. " +
    "Ce projet académique en 'Techniques de recherche d'emploi et Marketing de soi' m'a permis de construire un projet entrepreneurial complet, structuré et responsable. " +
    "Nos perspectives s'articulent sur trois horizons : la validation sur les campus universitaires à court terme, la cartographie des points réels avec les associations à moyen terme, et le passage à une Progressive Web App à long terme. " +
    "Je vous remercie sincèrement pour votre écoute et je suis maintenant à votre entière disposition pour répondre à toutes vos questions."
  );
}

// Enregistrement du fichier PPTX
const outputPath = path.join(__dirname, 'EcoTech_Digital_Presentation.pptx');
pres.writeFile({ fileName: outputPath })
  .then(fileName => {
    console.log(`Fichier PowerPoint généré avec succès : ${fileName}`);
  })
  .catch(err => {
    console.error('Erreur lors de la génération du fichier PPTX :', err);
    process.exit(1);
  });
