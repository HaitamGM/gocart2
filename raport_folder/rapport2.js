
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, NumberFormat, PageBreak, LevelFormat,
  TableOfContents, UnderlineType
} = require('docx');
const fs = require('fs');

// Colors
const BLUE_DARK = "1A3A5C";
const BLUE_MID  = "2563EB";
const BLUE_LIGHT = "DBEAFE";
const ACCENT = "3B82F6";
const GRAY_BG = "F1F5F9";
const GRAY_TEXT = "475569";
const WHITE = "FFFFFF";
const BLACK = "000000";
const GREEN = "166534";
const GREEN_BG = "DCFCE7";

const border = { style: BorderStyle.SINGLE, size: 1, color: "CBD5E1" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function h(level, text, color = BLUE_DARK) {
  const sizes = { 1: 36, 2: 28, 3: 24 };
  const spacing = { 1: { before: 480, after: 240 }, 2: { before: 360, after: 180 }, 3: { before: 240, after: 120 } };
  return new Paragraph({
    heading: level === 1 ? HeadingLevel.HEADING_1 : level === 2 ? HeadingLevel.HEADING_2 : HeadingLevel.HEADING_3,
    spacing: spacing[level],
    children: [new TextRun({ text, bold: true, size: sizes[level], color, font: "Arial" })]
  });
}


function placeholderDiagram(title) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.DASHED, size: 2, color: BLUE_DARK },
      bottom: { style: BorderStyle.DASHED, size: 2, color: BLUE_DARK },
      left: { style: BorderStyle.DASHED, size: 2, color: BLUE_DARK },
      right: { style: BorderStyle.DASHED, size: 2, color: BLUE_DARK }
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            shading: { fill: "F8FAFC" },
            verticalAlign: VerticalAlign.CENTER,
            children: [
              spacer(400),
              p(`[ EMPLACEMENT RÉSERVÉ : ${title} ]`, { center: true, bold: true, color: BLUE_MID }),
              p(`(Veuillez insérer l'image du diagramme ici avant l'impression)`, { center: true, italic: true }),
              spacer(400)
            ]
          })
        ]
      })
    ]
  });
}



function codeSnippet(codeStr) {
  return new Paragraph({
    shading: { type: ShadingType.CLEAR, color: "auto", fill: GRAY_BG },
    spacing: { before: 120, after: 120, line: 360 },
    children: [new TextRun({ text: codeStr, size: 18, font: "Courier New", color: "1E293B" })]
  });
}


function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.center ? AlignmentType.CENTER : opts.justify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
    spacing: { before: opts.before || 80, after: opts.after || 120, line: opts.lineSpacing || 320 },
    children: [new TextRun({
      text,
      size: opts.size || 22,
      color: opts.color || BLACK,
      bold: opts.bold || false,
      italics: opts.italic || false,
      font: "Arial"
    })]
  });
}

function pRich(runs, opts = {}) {
  return new Paragraph({
    alignment: opts.justify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
    spacing: { before: opts.before || 80, after: opts.after || 120, line: opts.lineSpacing || 320 },
    children: runs
  });
}

function run(text, opts = {}) {
  return new TextRun({
    text,
    size: opts.size || 22,
    color: opts.color || BLACK,
    bold: opts.bold || false,
    italics: opts.italic || false,
    font: "Arial",
    underline: opts.underline ? { type: UnderlineType.SINGLE } : undefined,
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { before: 60, after: 60, line: 300 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function spacer(before = 120) {
  return new Paragraph({ spacing: { before, after: 0 }, children: [new TextRun("")] });
}

function sectionTitle(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 480, after: 240 },
    children: [
      new TextRun({ text: number ? `${number}. ` : "", bold: true, size: 36, color: BLUE_MID, font: "Arial" }),
      new TextRun({ text: title, bold: true, size: 36, color: BLUE_DARK, font: "Arial" })
    ]
  });
}

function subTitle(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 160 },
    children: [
      new TextRun({ text: `${number} `, bold: true, size: 28, color: BLUE_MID, font: "Arial" }),
      new TextRun({ text: title, bold: true, size: 28, color: BLUE_DARK, font: "Arial" })
    ]
  });
}

function subSubTitle(number, title) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 100 },
    children: [
      new TextRun({ text: `${number} `, bold: true, size: 24, color: ACCENT, font: "Arial" }),
      new TextRun({ text: title, bold: true, size: 24, color: BLUE_DARK, font: "Arial" })
    ]
  });
}

function infoBox(text, bgColor = BLUE_LIGHT, textColor = BLUE_DARK) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({
      children: [new TableCell({
        borders: noBorders,
        width: { size: 9026, type: WidthType.DXA },
        shading: { fill: bgColor, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 240, right: 240 },
        children: [new Paragraph({
          spacing: { line: 300 },
          children: [new TextRun({ text, size: 22, color: textColor, font: "Arial", italics: true })]
        })]
      })]
    })]
  });
}

function twoColTable(leftTitle, leftContent, rightTitle, rightContent) {
  const cellProps = (title, content) => new TableCell({
    borders: {
      top: noBorder,
      bottom: { style: BorderStyle.SINGLE, size: 2, color: "CBD5E1" },
      left: noBorder,
      right: noBorder
    },
    width: { size: 4513, type: WidthType.DXA },
    shading: { fill: "F8FAFC", type: ShadingType.CLEAR },
    margins: { top: 120, bottom: 120, left: 160, right: 160 },
    children: [
      new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: title, bold: true, size: 22, color: BLUE_DARK, font: "Arial" })] }),
      new Paragraph({ spacing: { before: 0, after: 0, line: 300 }, children: [new TextRun({ text: content, size: 21, color: "334155", font: "Arial" })] })
    ]
  });
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [4513, 4513],
    rows: [new TableRow({ children: [cellProps(leftTitle, leftContent), cellProps(rightTitle, rightContent)] })]
  });
}

function headerRow(cells, widths) {
  return new TableRow({
    tableHeader: true,
    children: cells.map((text, i) => new TableCell({
      borders: {
        top: { style: BorderStyle.SINGLE, size: 6, color: "1A3A5C" },
        bottom: { style: BorderStyle.SINGLE, size: 6, color: "1A3A5C" },
        left: noBorder,
        right: noBorder
      },
      width: { size: widths[i], type: WidthType.DXA },
      shading: { fill: "F8FAFC", type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 120, right: 120 },
      verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text, bold: true, size: 20, color: "1A3A5C", font: "Arial" })] })]
    }))
  });
}

function dataRow(cells, widths, bg = WHITE) {
  return new TableRow({
    children: cells.map((text, i) => new TableCell({
      borders: {
        top: noBorder,
        bottom: { style: BorderStyle.SINGLE, size: 2, color: "E2E8F0" },
        left: noBorder,
        right: noBorder
      },
      width: { size: widths[i], type: WidthType.DXA },
      shading: { fill: WHITE, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text, size: 20, font: "Arial", color: "334155" })] })]
    }))
  });
}

// ─── COVER PAGE ────────────────────────────────────────────────────────────────
const coverPage = [
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({
      children: [new TableCell({
        borders: noBorders,
        shading: { fill: BLUE_DARK, type: ShadingType.CLEAR },
        margins: { top: 300, bottom: 300, left: 400, right: 400 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "RAPPORT DE STAGE", bold: true, size: 52, color: WHITE, font: "Arial" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80 }, children: [new TextRun({ text: "Projet de Fin d'Études — Cycle Ingénieur", size: 26, color: "93C5FD", font: "Arial" })] }),
        ]
      })]
    })]
  }),
  spacer(400),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 8, color: BLUE_MID }, bottom: { style: BorderStyle.SINGLE, size: 8, color: BLUE_MID }, left: noBorder, right: noBorder },
        shading: { fill: BLUE_LIGHT, type: ShadingType.CLEAR },
        margins: { top: 320, bottom: 320, left: 400, right: 400 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Ham Voice", bold: true, size: 56, color: BLUE_MID, font: "Arial" })] }),
          spacer(80),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Plateforme SaaS d'Assistant Vocal Intelligent", bold: true, size: 32, color: BLUE_DARK, font: "Arial" })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100 }, children: [new TextRun({ text: "pour le Commerce Électronique en Darija Marocaine", bold: true, size: 32, color: BLUE_DARK, font: "Arial" })] }),
        ]
      })]
    })]
  }),
  spacer(400),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2400, 6626],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders, width: { size: 2400, type: WidthType.DXA }, shading: { fill: GRAY_BG, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Auteurs", bold: true, size: 22, color: BLUE_DARK, font: "Arial" })] })] }),
        new TableCell({ borders, width: { size: 6626, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Mohammed Ahakam, Haitam Gmira, Ayman El khaldi", size: 22, font: "Arial" })] })] })
      ]}),
      new TableRow({ children: [
        new TableCell({ borders, width: { size: 2400, type: WidthType.DXA }, shading: { fill: GRAY_BG, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Filière", bold: true, size: 22, color: BLUE_DARK, font: "Arial" })] })] }),
        new TableCell({ borders, width: { size: 6626, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Intelligence Artificielle", size: 22, font: "Arial" })] })] })
      ]}),
      new TableRow({ children: [
        new TableCell({ borders, width: { size: 2400, type: WidthType.DXA }, shading: { fill: GRAY_BG, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Établissement", bold: true, size: 22, color: BLUE_DARK, font: "Arial" })] })] }),
        new TableCell({ borders, width: { size: 6626, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "École Supérieure de Technologie de Tétouan", size: 22, font: "Arial" })] })] })
      ]}),
      new TableRow({ children: [
        new TableCell({ borders, width: { size: 2400, type: WidthType.DXA }, shading: { fill: GRAY_BG, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Organisme d'accueil", bold: true, size: 22, color: BLUE_DARK, font: "Arial" })] })] }),
        new TableCell({ borders, width: { size: 6626, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Cap Connect", size: 22, font: "Arial" })] })] })
      ]}),
      new TableRow({ children: [
        new TableCell({ borders, width: { size: 2400, type: WidthType.DXA }, shading: { fill: GRAY_BG, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Tuteur de stage", bold: true, size: 22, color: BLUE_DARK, font: "Arial" })] })] }),
        new TableCell({ borders, width: { size: 6626, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Mohamed Idkoukou", size: 22, font: "Arial" })] })] })
      ]}),
      new TableRow({ children: [
        new TableCell({ borders, width: { size: 2400, type: WidthType.DXA }, shading: { fill: GRAY_BG, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Période", bold: true, size: 22, color: BLUE_DARK, font: "Arial" })] })] }),
        new TableCell({ borders, width: { size: 6626, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ children: [new TextRun({ text: "Du 27 Avril au 27 Juin 2026", size: 22, font: "Arial" })] })] })
      ]}),
    ]
  }),
  spacer(300),
  new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Année Universitaire 2025 — 2026", bold: true, size: 24, color: GRAY_TEXT, font: "Arial" })] }),
  pageBreak()
];

// ─── REMERCIEMENTS ─────────────────────────────────────────────────────────────
const remerciements = [
  sectionTitle("", "Remerciements"),
  spacer(120),
  p("Au terme de ce stage, il m'est agréable d'adresser mes vifs remerciements et ma profonde gratitude à toutes les personnes qui ont contribué, de près ou de loin, à la réalisation de ce travail.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  p("Je tiens à remercier en premier lieu [Nom de l'encadrant professionnel], mon encadrant au sein de l'entreprise, pour sa disponibilité, ses précieux conseils et son accompagnement tout au long de cette expérience professionnelle. Ses orientations techniques et méthodologiques ont été déterminantes dans la réussite de ce projet.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  p("Mes remerciements vont également à [Nom du responsable pédagogique], mon encadrant pédagogique, dont le suivi régulier et les retours constructifs ont guidé la rédaction de ce rapport.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  p("Je remercie l'ensemble de l'équipe technique pour leur accueil chaleureux et leur esprit de collaboration, qui ont rendu cette période de stage à la fois enrichissante et stimulante.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  p("Enfin, je témoigne ma reconnaissance à [Nom de l'École] pour la formation de qualité dispensée, qui m'a fourni les bases théoriques et pratiques indispensables à la réalisation de ce projet.", { justify: true, lineSpacing: 340 }),
  pageBreak()
];

// ─── RÉSUMÉ ────────────────────────────────────────────────────────────────────
const resume = [
  sectionTitle("", "Résumé"),
  spacer(120),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({ children: [new TableCell({
      borders: { left: { style: BorderStyle.SINGLE, size: 12, color: BLUE_MID }, top: noBorder, bottom: noBorder, right: noBorder },
      shading: { fill: BLUE_LIGHT, type: ShadingType.CLEAR },
      margins: { top: 200, bottom: 200, left: 280, right: 200 },
      children: [
        new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "Résumé", bold: true, size: 24, color: BLUE_DARK, font: "Arial" })] }),
        new Paragraph({ spacing: { line: 340, before: 0, after: 0 }, children: [new TextRun({ text: "Ce rapport présente Ham Voice, une plateforme SaaS innovante développée durant un stage de fin d'études. Ham Voice permet aux marchands e-commerce d'intégrer un assistant vocal interactif et intelligent dans leurs boutiques en ligne. L'assistant communique naturellement en Darija marocaine, renseigne les clients sur les produits du catalogue et exécute des actions directement sur la page web — comme ajouter un article au panier ou naviguer vers un produit — sur simple commande vocale. La plateforme s'appuie sur l'API Gemini 1.5 Flash Live de Google, utilisant des WebSockets bidirectionnels à faible latence pour le streaming audio en temps réel. L'ensemble du système est déployé sur Railway Cloud.", size: 22, color: "1e293b", font: "Arial" })] }),
      ]
    })] })]
  }),
  spacer(160),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({ children: [new TableCell({
      borders: { left: { style: BorderStyle.SINGLE, size: 12, color: "059669" }, top: noBorder, bottom: noBorder, right: noBorder },
      shading: { fill: GREEN_BG, type: ShadingType.CLEAR },
      margins: { top: 200, bottom: 200, left: 280, right: 200 },
      children: [
        new Paragraph({ spacing: { before: 0, after: 80 }, children: [new TextRun({ text: "Abstract", bold: true, size: 24, color: GREEN, font: "Arial" })] }),
        new Paragraph({ spacing: { line: 340, before: 0, after: 0 }, children: [new TextRun({ text: "This report presents Ham Voice, an innovative SaaS platform developed during an end-of-studies internship. Ham Voice enables e-commerce merchants to integrate a real-time, interactive, and intelligent voice assistant into their online stores. The assistant communicates naturally in Moroccan Darija, informs customers about catalog products, and automatically executes on-screen actions — such as adding items to the cart or navigating to a product — based on voice commands. The platform is powered by Google's Gemini 1.5 Flash Live API, using low-latency bidirectional WebSockets for audio streaming. The entire system is deployed on Railway Cloud.", size: 22, color: "1e293b", font: "Arial" })] }),
      ]
    })] })]
  }),
  spacer(160),
  new Paragraph({ spacing: { before: 120, after: 120 }, children: [new TextRun({ text: "Mots-clés : ", bold: true, size: 22, color: BLUE_DARK, font: "Arial" }), new TextRun({ text: "SaaS, Assistant vocal, Darija marocaine, Gemini Live API, WebSocket, FastAPI, MongoDB, E-commerce, Intelligence Artificielle, Railway Cloud", size: 22, font: "Arial", italics: true })] }),
  pageBreak()
];

// ─── TABLE DES MATIÈRES ─────────────────────────────────────────────────────────
const toc = [
  sectionTitle("", "Table des Matières"),
  new TableOfContents("Table des Matières", {
    hyperlink: true,
    headingStyleRange: "1-3",
    stylesWithLevels: [
      { styleName: "Heading1", level: 1 },
      { styleName: "Heading2", level: 2 },
      { styleName: "Heading3", level: 3 }
    ]
  }),
  pageBreak()
];

// ─── LISTE DES ABRÉVIATIONS ────────────────────────────────────────────────────
const abreviations = [
  sectionTitle("", "Liste des Abréviations"),
  spacer(80),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2200, 6826],
    rows: [
      headerRow(["Abréviation", "Signification"], [2200, 6826]),
      dataRow(["API", "Application Programming Interface (Interface de Programmation d'Application)"], [2200, 6826], GRAY_BG),
      dataRow(["SaaS", "Software as a Service (Logiciel en tant que Service)"], [2200, 6826], WHITE),
      dataRow(["WebSocket", "Protocole de communication bidirectionnel full-duplex sur TCP"], [2200, 6826], GRAY_BG),
      dataRow(["PCM", "Pulse-Code Modulation (Modulation par Impulsions et Codage)"], [2200, 6826], WHITE),
      dataRow(["LLM", "Large Language Model (Grand Modèle de Langage)"], [2200, 6826], GRAY_BG),
      dataRow(["JWT", "JSON Web Token (Jeton Web JSON)"], [2200, 6826], WHITE),
      dataRow(["CSV", "Comma-Separated Values (Valeurs Séparées par des Virgules)"], [2200, 6826], GRAY_BG),
      dataRow(["MongoDB", "Base de données NoSQL orientée documents"], [2200, 6826], WHITE),
      dataRow(["FastAPI", "Framework Python moderne pour la création d'APIs REST"], [2200, 6826], GRAY_BG),
      dataRow(["IA / AI", "Intelligence Artificielle / Artificial Intelligence"], [2200, 6826], WHITE),
      dataRow(["SDK", "Software Development Kit (Kit de Développement Logiciel)"], [2200, 6826], GRAY_BG),
      dataRow(["TTS", "Text-To-Speech (Synthèse Vocale)"], [2200, 6826], WHITE),
      dataRow(["STT", "Speech-To-Text (Reconnaissance Vocale)"], [2200, 6826], GRAY_BG),
      dataRow(["CI/CD", "Continuous Integration / Continuous Deployment"], [2200, 6826], WHITE),
    ]
  }),
  pageBreak()
];

// ─── INTRODUCTION ──────────────────────────────────────────────────────────────
const introduction = [
  sectionTitle("", "Introduction Générale"),
  spacer(120),
  p("Le commerce électronique connaît une croissance exponentielle à l'échelle mondiale, et le Maroc ne fait pas exception à cette tendance. Selon les dernières statistiques, le marché du e-commerce marocain a enregistré une progression constante, portée par l'essor de la connectivité mobile et la démocratisation des paiements numériques. Toutefois, une barrière persiste pour une large frange de la population : la langue.", { justify: true, lineSpacing: 340, after: 140 }),
  p("La plupart des interfaces de commerce en ligne sont développées en arabe formel, en français ou en anglais, laissant de côté la Darija marocaine — dialecte arabe local parlé au quotidien par la quasi-totalité des Marocains. Cette inadéquation linguistique crée une distance entre le consommateur et le produit, limitant ainsi l'expérience utilisateur et le potentiel de conversion.", { justify: true, lineSpacing: 340, after: 140 }),
  p("C'est dans ce contexte qu'a émergé l'idée de Ham Voice : une plateforme SaaS (Software as a Service) qui intègre un assistant vocal interactif en Darija marocaine. Le système repose sur l'API Gemini 3.1 Flash Live de Google, communiquant en temps réel via WebSockets avec un backend Python FastAPI, et stockant les données via MongoDB. Cette plateforme permet aux marchands e-commerce d'intégrer un assistant vocal intelligent : une plateforme SaaS (Software as a Service) qui permet aux marchands e-commerce d'intégrer un assistant vocal intelligent, capable de communiquer naturellement en Darija marocaine. Alimenté par l'API Gemini Live de Google, cet assistant offre une expérience d'achat conversationnelle, intuitive et culturellement adaptée.", { justify: true, lineSpacing: 340, after: 140 }),
  spacer(80),
  subTitle("", "Objectifs du Stage"),
  p("Ce stage de fin d'études avait pour objectifs principaux de :", { after: 60 }),
  bullet("Concevoir et développer une architecture SaaS multi-tenant robuste et scalable"),
  bullet("Implémenter l'intégration de l'API Gemini 1.5 Flash Live pour le streaming audio bidirectionnel en temps réel"),
  bullet("Créer un moteur vocal capable d'interpréter des commandes en Darija marocaine et d'exécuter des actions e-commerce"),
  bullet("Mettre en place un portail d'administration permettant aux marchands de gérer leur abonnement et leur catalogue"),
  bullet("Déployer l'ensemble de la plateforme sur Railway Cloud dans un environnement de production"),
  spacer(120),
  subTitle("", "Structure du Rapport"),
  p("Ce rapport est organisé en six chapitres. Après une présentation du contexte et de l'organisme d'accueil, nous exposons l'analyse des besoins et la spécification fonctionnelle du projet. Les chapitres suivants détaillent l'architecture technique, les choix technologiques, le développement des trois composants principaux de la plateforme, puis le déploiement sur Railway Cloud. Le rapport se conclut par un bilan des acquis et les perspectives d'évolution.", { justify: true, lineSpacing: 340 }),
  pageBreak()
];

// ─── CHAPITRE 1 — CONTEXTE & ORGANISME ────────────────────────────────────────
const chapitre1 = [
  sectionTitle("Chapitre 1", "Contexte et Organisme d'Accueil"),
  spacer(100),
  subTitle("1.1", "Présentation de l'Organisme d'Accueil"),
  p("Cap Connect est une entreprise marocaine spécialisée dans le développement de solutions numériques innovantes, positionnée à l'intersection de l'intelligence artificielle, du cloud computing et du commerce électronique. Fondée à Tétouan, la société accompagne des startups ambitieuses et des PME en quête de transformation digitale, en proposant des services de développement logiciel sur mesure, de conseil technologique et d'intégration de solutions basées sur l'IA générative.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  p("L'entreprise se distingue par une culture d'innovation forte et une capacité à allier expertise technique et compréhension des réalités du marché marocain. Son équipe pluridisciplinaire réunit des ingénieurs spécialisés en intelligence artificielle, des architectes cloud, des développeurs full-stack et des experts en expérience utilisateur. Cette synergie de compétences permet à Cap Connect de concevoir et déployer des produits technologiques à fort impact, pensés pour répondre aux spécificités culturelles et linguistiques du contexte maghrébin.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  p("Le projet Ham Voice s'inscrit pleinement dans la vision stratégique de Cap Connect : démocratiser l'accès à l'intelligence artificielle conversationnelle pour les entreprises marocaines, en commençant par le secteur du e-commerce — un marché en pleine expansion qui manque cruellement d'interfaces vocales adaptées à la langue et à la culture locale.", { justify: true, lineSpacing: 340 }),
  spacer(100),
  twoColTable("Raison Sociale", "Cap Connect", "Secteur d'Activité", "Technologies & Innovation Numérique"),
  spacer(60),
  twoColTable("Forme Juridique", "SARL", "Effectif", "15 à 30 collaborateurs"),
  spacer(60),
  twoColTable("Siège Social", "Tétouan, Maroc", "Site Web", "capconnect.ma"),
  spacer(60),
  twoColTable("Domaines", "IA, SaaS, E-commerce, Cloud", "Année de création", "2020"),

  spacer(200),
  subTitle("1.2", "Contexte du Projet"),
  subSubTitle("1.2.1", "Le Marché du E-commerce au Maroc"),
  p("Le secteur du commerce électronique au Maroc connaît une croissance soutenue depuis plusieurs années. Les études de marché révèlent que le nombre d'acheteurs en ligne au Maroc a considérablement augmenté, porté notamment par la pénétration croissante d'internet et l'usage massif des smartphones. Cependant, la majorité des plateformes marchandes disponibles ne proposent pas d'interface vocale, et encore moins en Darija.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  subSubTitle("1.2.2", "La Darija Marocaine : Un Enjeu Numérique"),
  p("La Darija (الدارجة) est le dialecte arabe marocain, utilisé dans la vie quotidienne par l'ensemble de la population. Contrairement à l'arabe classique ou au français, la Darija est la langue dans laquelle les Marocains pensent et s'expriment spontanément. Pourtant, les solutions technologiques — notamment les assistants vocaux — sont très peu adaptées à ce dialecte.", { justify: true, lineSpacing: 340 }),
  p("Ham Voice répond à ce vide en offrant une interface vocale nativement en Darija, rendant ainsi l'expérience d'achat en ligne plus accessible, plus naturelle et culturellement pertinente pour le consommateur marocain.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  subTitle("1.3", "Mission Confiée"),
  p("Dans le cadre de ce stage, la mission principale consistait à concevoir et développer from scratch l'intégralité de la plateforme Ham Voice. Cette mission comprenait :", { after: 60 }),
  bullet("L'architecture générale du système SaaS multi-tenant"),
  bullet("Le développement du portail d'onboarding et du tableau de bord administrateur (ham-landing)"),
  bullet("L'implémentation du moteur vocal basé sur Gemini Live (saas-platform)"),
  bullet("La création d'une boutique démo intégrant le widget vocal (ham-keyboard)"),
  bullet("Le déploiement et la mise en production sur Railway Cloud"),
  spacer(200),
  subTitle("1.4", "Analyse SWOT du Projet"),
  p("Afin de mieux appréhender le contexte stratégique du projet Ham Voice, une analyse SWOT (Forces, Faiblesses, Opportunités, Menaces) a été conduite en début de stage. Cette analyse a permis d'orienter les choix architecturaux et fonctionnels du projet.", { justify: true, lineSpacing: 340, after: 100 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [4513, 4513],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, shading: { fill: "E2F0D9", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 80, left: 200, right: 160 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Forces (Strengths)", bold: true, size: 22, color: "166534", font: "Arial" })] })] }),
        new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, shading: { fill: "FCE4D6", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 80, left: 200, right: 160 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Faiblesses (Weaknesses)", bold: true, size: 22, color: "C00000", font: "Arial" })] })] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 200, right: 160 }, children: [
          ...["Technologie de pointe (Gemini Live API)", "Première solution vocale en Darija marocaine", "Architecture SaaS scalable et multi-tenant", "Intégration e-commerce simple via widget JS", "Faible latence audio (~320ms)"].map(t => new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: "• " + t, size: 21, font: "Arial", color: "334155" })] }))
        ] }),
        new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 200, right: 160 }, children: [
          ...["Dépendance à une API tierce (Google Gemini)", "Nécessite une connexion internet stable", "Support limité aux boutiques web (pas mobile natif)", "Coûts API liés au volume d'usage vocal", "Dialecte Darija non standardisé (variantes régionales)"].map(t => new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: "• " + t, size: 21, font: "Arial", color: "334155" })] }))
        ] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, shading: { fill: "DDEBF7", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 80, left: 200, right: 160 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Opportunités (Opportunities)", bold: true, size: 22, color: "1F4E79", font: "Arial" })] })] }),
        new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, shading: { fill: "FFF2CC", type: ShadingType.CLEAR }, margins: { top: 100, bottom: 80, left: 200, right: 160 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Menaces (Threats)", bold: true, size: 22, color: "7F6000", font: "Arial" })] })] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 200, right: 160 }, children: [
          ...["Marché e-commerce marocain en forte croissance", "Absence de concurrents directs en Darija", "Potentiel d'extension à d'autres dialectes arabes", "Demande croissante d'accessibilité numérique", "Intégrations futures Shopify / WooCommerce"].map(t => new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: "• " + t, size: 21, font: "Arial", color: "334155" })] }))
        ] }),
        new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 200, right: 160 }, children: [
          ...["Évolution rapide des modèles d'IA (obsolescence)", "Entrée potentielle de géants tech sur le marché", "Sensibilité des utilisateurs à la confidentialité audio", "Coût de la bande passante en production", "Risque de changement de politique tarifaire Google"].map(t => new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: "• " + t, size: 21, font: "Arial", color: "334155" })] }))
        ] }),
      ]}),
    ]
  }),
  spacer(200),
  subTitle("1.5", "Planning et Déroulement du Stage"),
  p("Le stage s'est déroulé sur une période de deux mois, du 27 avril au 27 juin 2026. Le travail a été organisé en sprints successifs, suivant une méthodologie agile adaptée au contexte d'une startup technologique.", { justify: true, lineSpacing: 340, after: 100 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [1400, 3000, 4626],
    rows: [
      headerRow(["Sprint", "Période", "Objectifs"], [1400, 3000, 4626]),
      dataRow(["Sprint 1", "27 Avril — 10 Mai", "Analyse des besoins, conception de l'architecture, mise en place de l'environnement de développement"], [1400, 3000, 4626], GRAY_BG),
      dataRow(["Sprint 2", "11 Mai — 24 Mai", "Développement de ham-landing : portail d'inscription, gestion des clés API, parsing CSV"], [1400, 3000, 4626], WHITE),
      dataRow(["Sprint 3", "25 Mai — 07 Juin", "Développement de saas-platform : WebSocket, intégration Gemini Live, streaming audio, function calls"], [1400, 3000, 4626], GRAY_BG),
      dataRow(["Sprint 4", "08 Juin — 17 Juin", "Développement de ham-keyboard (boutique démo) et tests d'intégration avec GoCart"], [1400, 3000, 4626], WHITE),
      dataRow(["Sprint 5", "18 Juin — 27 Juin", "Déploiement Railway Cloud, tests de performance, corrections, rédaction du rapport de stage"], [1400, 3000, 4626], GRAY_BG),
    ]
  }),
  spacer(120),
  p("Cette organisation en sprints a permis de livrer des fonctionnalités opérationnelles à chaque fin de cycle, favorisant les retours rapides de l'encadrant et les ajustements itératifs. La communication quotidienne avec l'équipe Cap Connect a assuré une cohérence permanente entre les développements réalisés et les objectifs stratégiques du produit.", { justify: true, lineSpacing: 340 }),
  pageBreak()
];

// ─── CHAPITRE 2 — ANALYSE DES BESOINS ─────────────────────────────────────────
const chapitre2 = [
  sectionTitle("Chapitre 2", "Analyse des Besoins et Spécification"),
  spacer(100),
  subTitle("2.1", "Identification des Acteurs"),
  p("La plateforme Ham Voice implique trois catégories d'acteurs principaux, chacun ayant des besoins et des interactions spécifiques avec le système :", { after: 80 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2000, 3513, 3513],
    rows: [
      headerRow(["Acteur", "Rôle", "Interactions Principales"], [2000, 3513, 3513]),
      dataRow(["Opérateur / Admin", "Propriétaire de la plateforme Ham Voice", "Gérer les clients, suspendre/activer les clés API, monitorer l'usage"], [2000, 3513, 3513], GRAY_BG),
      dataRow(["Marchand / Client SaaS", "E-commerçant souscrivant au service", "S'inscrire, choisir un plan, uploader le catalogue, obtenir une clé API"], [2000, 3513, 3513], WHITE),
      dataRow(["Consommateur Final", "Visiteur de la boutique du marchand", "Interagir vocalement avec l'assistant, ajouter au panier, naviguer"], [2000, 3513, 3513], GRAY_BG),
    ]
  }),
  spacer(200),
  subTitle("2.2", "Besoins Fonctionnels"),
  subSubTitle("2.2.1", "Portail Marchand (ham-landing)"),
  bullet("Inscription du marchand avec sélection d'un plan d'abonnement (Pro, etc.)"),
  bullet("Saisie des informations de l'entreprise : nom, numéro WhatsApp, coordonnées"),
  bullet("Upload du catalogue produits au format CSV"),
  bullet("Génération automatique d'une clé API unique (ex. : ham_a1b2c3d4)"),
  bullet("Parsing et stockage du catalogue produits en base de données MongoDB"),
  bullet("Attribution dynamique des outils disponibles selon le plan souscrit"),
  bullet("Dashboard administrateur : suivi des clients actifs, durées d'abonnement, minutes d'usage"),
  bullet("Gestion des clés API : activation, suspension, suppression"),
  spacer(80),
  subSubTitle("2.2.2", "Moteur Vocal (saas-platform)"),
  bullet("Connexion WebSocket sécurisée avec authentification par clé API"),
  bullet("Récupération dynamique de la configuration du marchand depuis MongoDB"),
  bullet("Capture audio du microphone (16kHz, PCM) et encodage base64"),
  bullet("Streaming bidirectionnel avec l'API Gemini Live de Google"),
  bullet("Lecture audio de la réponse (24kHz PCM) avec support du barge-in (interruption)"),
  bullet("Persona vocal en Darija marocaine avec prompt système spécialisé"),
  bullet("Traçabilité automatique : calcul du temps d'appel et mise à jour des minutes d'usage"),
  spacer(80),
  subSubTitle("2.2.3", "Widget E-commerce (ham-keyboard / GoCart)"),
  bullet("Bouton flottant d'activation de l'assistant vocal dans la boutique"),
  bullet("Exécution d'actions en temps réel sur la page web via les function calls Gemini"),
  bullet("Navigation vers un produit (navigate_to_product)"),
  bullet("Ajout d'articles au panier (add_to_cart)"),
  bullet("Ouverture du formulaire de paiement (checkout)"),
  bullet("Compatibilité avec des boutiques externes via l'injection de la clé API"),
  spacer(200),
  subTitle("2.3", "Besoins Non Fonctionnels"),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2500, 6526],
    rows: [
      headerRow(["Exigence", "Description"], [2500, 6526]),
      dataRow(["Performance", "Latence audio < 500ms pour une expérience conversationnelle fluide"], [2500, 6526], GRAY_BG),
      dataRow(["Scalabilité", "Architecture multi-tenant capable de servir des centaines de marchands simultanément"], [2500, 6526], WHITE),
      dataRow(["Sécurité", "Authentification par clé API, isolation des données par tenant, protection des endpoints"], [2500, 6526], GRAY_BG),
      dataRow(["Disponibilité", "Uptime garanti par Railway Cloud avec redémarrages automatiques"], [2500, 6526], WHITE),
      dataRow(["Maintenabilité", "Code modulaire, séparation des responsabilités, documentation technique"], [2500, 6526], GRAY_BG),
      dataRow(["Extensibilité", "Architecture permettant l'ajout de nouvelles langues et de nouveaux outils"], [2500, 6526], WHITE),
    ]
  }),
  spacer(200),
  subTitle("2.4", "Cas d'Utilisation Principaux"),
  spacer(60),
  infoBox("UC-01 : Inscription d'un marchand\nActeur : Marchand\nDescription : Le marchand accède au portail ham-landing, sélectionne un plan d'abonnement, renseigne ses informations et uploade son catalogue CSV. Le système génère et lui attribue une clé API unique.\nRésultat : La clé API est affichée au marchand. Son catalogue est parsé et stocké en base de données."),
  spacer(80),
  infoBox("UC-02 : Session vocale client\nActeur : Consommateur final\nDescription : Le visiteur de la boutique clique sur le bouton de l'assistant vocal. Une connexion WebSocket s'établit avec le moteur vocal. L'utilisateur pose une question en Darija. L'assistant répond en audio et peut exécuter des actions sur la page.\nRésultat : Interaction naturelle en temps réel, actions e-commerce exécutées automatiquement.", BLUE_LIGHT),
  spacer(80),
  infoBox("UC-03 : Administration de la plateforme\nActeur : Opérateur\nDescription : L'opérateur accède au dashboard admin. Il consulte la liste des marchands actifs, leurs minutes d'usage, les dates d'expiration. Il peut activer, suspendre ou supprimer une clé API.\nRésultat : Contrôle total de la plateforme SaaS par l'opérateur.", GREEN_BG, GREEN),
  spacer(200),
  subTitle("2.5", "Diagramme de Séquence — Session Vocale"),
  p("Le diagramme ci-dessous illustre le flux complet d'une interaction vocale entre un consommateur et l'assistant Ham Voice. Ce diagramme de séquence met en évidence les échanges entre les différents composants du système lors d'une session audio en temps réel.", { justify: true, lineSpacing: 340, after: 100 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [600, 1600, 1600, 5226],
    rows: [
      headerRow(["Étape", "Émetteur", "Récepteur", "Description de l'action"], [600, 1600, 1600, 5226]),
      dataRow(["1", "Utilisateur", "Widget JS", "Clic sur le bouton microphone de la boutique pour démarrer la session vocale."], [600, 1600, 1600, 5226]),
      dataRow(["2", "Widget JS", "saas-platform", "Établissement d'une connexion WebSocket sécurisée (WSS) avec le moteur vocal."], [600, 1600, 1600, 5226]),
      dataRow(["3", "saas-platform", "MongoDB", "Vérification de la clé API du marchand et récupération de sa configuration active."], [600, 1600, 1600, 5226]),
      dataRow(["4", "saas-platform", "Gemini Live", "Ouverture d'une session et d'un tunnel de streaming avec l'API Live de Google."], [600, 1600, 1600, 5226]),
      dataRow(["5", "Utilisateur", "Widget JS", "L'utilisateur parle naturellement en Darija marocaine au widget vocal."], [600, 1600, 1600, 5226]),
      dataRow(["6", "Widget JS", "saas-platform", "Streaming des données audio capturées (format PCM 16kHz) via le WebSocket."], [600, 1600, 1600, 5226]),
      dataRow(["7", "saas-platform", "Gemini Live", "Relais des paquets audio encodés en base64 à destination de l'API Gemini."], [600, 1600, 1600, 5226]),
      dataRow(["8", "Gemini Live", "saas-platform", "Génération des réponses en Darija (audio PCM 24kHz et informations de tool calls)."], [600, 1600, 1600, 5226]),
      dataRow(["9", "saas-platform", "Widget JS", "Envoi des données audio décodées et des instructions function calling au navigateur."], [600, 1600, 1600, 5226]),
      dataRow(["10", "Widget JS", "Utilisateur / DOM", "Lecture du flux audio haut-parleur et mise à jour dynamique du DOM (panier, redirection)."], [600, 1600, 1600, 5226]),
      dataRow(["11", "saas-platform", "MongoDB", "Incrémentation des minutes d'usage et journalisation de fin de connexion WebSocket."], [600, 1600, 1600, 5226]),
    ]
  }),
  spacer(200),
  subTitle("2.6", "Tableau des User Stories"),
  p("Les user stories suivantes ont été rédigées en début de projet pour cadrer le périmètre fonctionnel de Ham Voice. Elles ont servi de base à la planification des sprints et à la définition des critères d'acceptation.", { justify: true, lineSpacing: 340, after: 100 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [600, 1500, 3200, 3726],
    rows: [
      headerRow(["ID", "Rôle", "Je veux...", "Afin de..."], [600, 1500, 3200, 3726]),
      dataRow(["US-01", "Marchand", "M'inscrire sur la plateforme et obtenir une clé API", "Intégrer l'assistant vocal dans ma boutique immédiatement"], [600, 1500, 3200, 3726], GRAY_BG),
      dataRow(["US-02", "Marchand", "Uploader mon catalogue produits au format CSV", "Permettre à l'assistant de connaître tous mes produits"], [600, 1500, 3200, 3726], WHITE),
      dataRow(["US-03", "Marchand", "Consulter mes minutes d'usage vocal", "Contrôler ma consommation et anticiper les renouvellements"], [600, 1500, 3200, 3726], GRAY_BG),
      dataRow(["US-04", "Consommateur", "Poser une question en Darija sur un produit", "Obtenir une réponse instantanée et naturelle sans taper du texte"], [600, 1500, 3200, 3726], WHITE),
      dataRow(["US-05", "Consommateur", "Demander vocalement d'ajouter un produit au panier", "Effectuer mes achats sans utiliser le clavier ni la souris"], [600, 1500, 3200, 3726], GRAY_BG),
      dataRow(["US-06", "Consommateur", "Interrompre l'assistant à tout moment (barge-in)", "Avoir une conversation naturelle et fluide"], [600, 1500, 3200, 3726], WHITE),
      dataRow(["US-07", "Opérateur", "Visualiser tous les marchands actifs et leur usage", "Piloter la plateforme et gérer la facturation"], [600, 1500, 3200, 3726], GRAY_BG),
      dataRow(["US-08", "Opérateur", "Suspendre ou supprimer une clé API", "Maintenir le contrôle de la plateforme et gérer les abus"], [600, 1500, 3200, 3726], WHITE),
      dataRow(["US-09", "Marchand", "Intégrer le widget en 2 lignes de HTML", "Déployer l'assistant sans compétences techniques avancées"], [600, 1500, 3200, 3726], GRAY_BG),
      dataRow(["US-10", "Consommateur", "Accéder directement à la page d'un produit par la voix", "Naviguer dans la boutique sans effort et sans friction"], [600, 1500, 3200, 3726], WHITE),
    ]
  }),
  pageBreak()
];

// ─── CHAPITRE 3 — ARCHITECTURE TECHNIQUE ──────────────────────────────────────
const chapitre3 = [
  sectionTitle("Chapitre 3", "Architecture Technique et Choix Technologiques"),
  spacer(100),
  subTitle("3.1", "Architecture Globale de la Plateforme"),
  p("Ham Voice adopte une architecture en microservices organisée en trois couches distinctes, conformément aux principes de conception SaaS. Chaque couche est indépendante, déployable séparément, et communique via des interfaces clairement définies.", { justify: true, lineSpacing: 340 }),
  spacer(120),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2200, 2400, 2000, 2426],
    rows: [
      headerRow(["Composant", "Rôle principal", "Technologies", "Interfaces / Flux"], [2200, 2400, 2000, 2426]),
      dataRow(["ham-landing (Portail SaaS)", "Onboarding des marchands, gestion du catalogue CSV, dashboard opérateur et suivi de consommation.", "HTML/JS Vanilla, CSS, MongoDB Atlas", "API REST HTTP, liaison directe MongoDB persistance."], [2200, 2400, 2000, 2426]),
      dataRow(["saas-platform (Moteur Vocal)", "Orchestration des sessions vocales, relais vers Gemini Live et calcul de l'usage en minutes.", "Python, FastAPI, WebSockets, MongoDB", "WebSocket bidirectionnel (WSS), API Gemini Live."], [2200, 2400, 2000, 2426]),
      dataRow(["ham-keyboard (Boutique Démo / Widget)", "Interface utilisateur finale (bouton micro), capture audio micro et exécution des actions DOM.", "JavaScript Vanilla (Widget), HTML5 Web Audio API", "WebSocket client vers saas-platform."], [2200, 2400, 2000, 2426]),
      dataRow(["MongoDB Atlas (Base de données)", "Stockage persistant et isolé des marchands, des produits du catalogue et de l'historique d'usage.", "MongoDB NoSQL cloud-native", "Connexion sécurisée TLS (clients & platform)."], [2200, 2400, 2000, 2426]),
    ]
  }),
  spacer(200),
  subTitle("3.2", "Description des Trois Composants"),
  subSubTitle("3.2.1", "ham-landing — Le Portail SaaS"),
  p("ham-landing constitue le point d'entrée de la plateforme. C'est l'interface avec laquelle les marchands interagissent pour s'inscrire, configurer leur service et accéder à leurs métriques d'usage. Il expose également un tableau de bord d'administration réservé à l'opérateur de la plateforme.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  subSubTitle("3.2.2", "saas-platform — Le Moteur Vocal"),
  p("saas-platform est le cœur technique de Ham Voice. Ce serveur FastAPI gère l'ensemble des connexions WebSocket entrantes depuis les widgets installés dans les boutiques des marchands. Il orchestre la communication avec l'API Gemini Live et assure la traçabilité de chaque session vocale.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  subSubTitle("3.2.3", "ham-keyboard — La Boutique Démo"),
  p("ham-keyboard est un exemple concret d'intégration du widget Ham Voice dans une boutique e-commerce. Cette boutique fictive — spécialisée dans les claviers mécaniques haut de gamme — illustre l'ensemble des capacités de l'assistant vocal en conditions réelles.", { justify: true, lineSpacing: 340 }),
  spacer(200),
  subTitle("3.3", "Choix Technologiques"),
  spacer(80),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2200, 2500, 4326],
    rows: [
      headerRow(["Composant", "Technologie", "Justification"], [2200, 2500, 4326]),
      dataRow(["Backend API", "FastAPI (Python)", "Performance asynchrone native, support WebSocket, documentation auto-générée (OpenAPI)"], [2200, 2500, 4326], GRAY_BG),
      dataRow(["Frontend", "HTML/CSS/JS Vanilla", "Légèreté maximale, intégration facile dans toute boutique existante"], [2200, 2500, 4326], WHITE),
      dataRow(["IA / LLM", "Google Gemini 1.5 Flash Live", "Seule API disponible offrant streaming audio bidirectionnel en temps réel à faible latence"], [2200, 2500, 4326], GRAY_BG),
      dataRow(["Base de données", "MongoDB Atlas", "Flexibilité du schéma pour les catalogues multi-produits, scalabilité cloud native"], [2200, 2500, 4326], WHITE),
      dataRow(["Communication", "WebSocket (ws://)", "Protocole full-duplex indispensable pour le streaming audio en temps réel"], [2200, 2500, 4326], GRAY_BG),
      dataRow(["Audio Format", "PCM 16kHz (in) / 24kHz (out)", "Format natif exigé par l'API Gemini Live pour une qualité optimale"], [2200, 2500, 4326], WHITE),
      dataRow(["Hébergement", "Railway Cloud", "Déploiement simplifié, support Docker, variables d'environnement, domaines HTTPS"], [2200, 2500, 4326], GRAY_BG),
      dataRow(["Gestion des données", "Python / pandas (CSV parsing)", "Traitement robuste des catalogues CSV uploadés par les marchands"], [2200, 2500, 4326], WHITE),
    ]
  }),
  spacer(200),
  subTitle("3.4", "Schéma de la Base de Données MongoDB"),
  p("Ham Voice utilise MongoDB Atlas comme base de données principale. La flexibilité du modèle documentaire de MongoDB est particulièrement adaptée à la nature hétérogène des catalogues produits des marchands, dont la structure peut varier d'un client à l'autre. La base de données est organisée autour de trois collections principales.", { justify: true, lineSpacing: 340, after: 100 }),
  subSubTitle("3.4.1", "Collection : clients"),
  p("Cette collection stocke toutes les informations relatives aux marchands inscrits sur la plateforme. Chaque document représente un compte marchand avec l'ensemble de ses métadonnées opérationnelles.", { justify: true, lineSpacing: 340, after: 80 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2200, 1800, 5026],
    rows: [
      headerRow(["Champ", "Type", "Description"], [2200, 1800, 5026]),
      dataRow(["_id", "ObjectId", "Identifiant unique généré automatiquement par MongoDB"], [2200, 1800, 5026], GRAY_BG),
      dataRow(["apiKey", "String", "Clé API unique du marchand (ex: ham_a1b2c3d4), indexée pour recherche rapide"], [2200, 1800, 5026], WHITE),
      dataRow(["businessName", "String", "Nom commercial de la boutique du marchand"], [2200, 1800, 5026], GRAY_BG),
      dataRow(["whatsappNumber", "String", "Numéro WhatsApp de contact du marchand"], [2200, 1800, 5026], WHITE),
      dataRow(["plan", "String", "Plan souscrit : 'basic' ou 'pro_agentic'"], [2200, 1800, 5026], GRAY_BG),
      dataRow(["status", "String", "État du compte : 'active', 'suspended', 'expired'"], [2200, 1800, 5026], WHITE),
      dataRow(["toolConfigs", "Array", "Liste des function calls autorisés selon le plan"], [2200, 1800, 5026], GRAY_BG),
      dataRow(["usageMinutes", "Float", "Total des minutes d'usage vocal accumulées"], [2200, 1800, 5026], WHITE),
      dataRow(["subscriptionStart", "Date", "Date de début de l'abonnement"], [2200, 1800, 5026], GRAY_BG),
      dataRow(["subscriptionEnd", "Date", "Date d'expiration de l'abonnement"], [2200, 1800, 5026], WHITE),
      dataRow(["createdAt", "Date", "Horodatage de création du compte"], [2200, 1800, 5026], GRAY_BG),
    ]
  }),
  spacer(160),
  subSubTitle("3.4.2", "Collection : products"),
  p("Chaque marchand dispose d'une collection de produits qui reflète le catalogue CSV uploadé lors de l'inscription. Ces documents sont utilisés en temps réel par le moteur vocal pour répondre aux questions des consommateurs.", { justify: true, lineSpacing: 340, after: 80 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2200, 1800, 5026],
    rows: [
      headerRow(["Champ", "Type", "Description"], [2200, 1800, 5026]),
      dataRow(["_id", "ObjectId", "Identifiant unique du produit"], [2200, 1800, 5026], GRAY_BG),
      dataRow(["apiKey", "String", "Clé API du marchand propriétaire (clé d'isolation multi-tenant)"], [2200, 1800, 5026], WHITE),
      dataRow(["productId", "String", "Identifiant métier du produit (ex: 001, SKU)"], [2200, 1800, 5026], GRAY_BG),
      dataRow(["name", "String", "Nom du produit tel qu'affiché dans la boutique"], [2200, 1800, 5026], WHITE),
      dataRow(["description", "String", "Description détaillée du produit"], [2200, 1800, 5026], GRAY_BG),
      dataRow(["price", "Float", "Prix en dirhams marocains (MAD)"], [2200, 1800, 5026], WHITE),
      dataRow(["category", "String", "Catégorie du produit (keyboards, switches, keycaps...)"], [2200, 1800, 5026], GRAY_BG),
      dataRow(["inStock", "Boolean", "Disponibilité du produit en stock"], [2200, 1800, 5026], WHITE),
    ]
  }),
  spacer(200),
  subTitle("3.5", "Architecture de Sécurité"),
  p("La sécurité est un pilier fondamental de Ham Voice en tant que plateforme SaaS multi-tenant. Plusieurs mécanismes ont été mis en place pour garantir l'isolation des données, l'authentification des accès et la protection des flux audio.", { justify: true, lineSpacing: 340, after: 100 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2800, 6226],
    rows: [
      headerRow(["Mécanisme de Sécurité", "Description et Implémentation"], [2800, 6226]),
      dataRow(["Authentification par clé API", "Chaque requête WebSocket doit inclure une clé API valide en paramètre (?apiKey=ham_xxx). La clé est vérifiée contre MongoDB avant toute session."], [2800, 6226], GRAY_BG),
      dataRow(["Isolation multi-tenant", "Toutes les requêtes MongoDB sont filtrées par apiKey. Un marchand ne peut jamais accéder aux données d'un autre, même en cas de tentative malveillante."], [2800, 6226], WHITE),
      dataRow(["Chiffrement TLS/WSS", "Toutes les connexions WebSocket utilisent le protocole wss:// (WebSocket Secure), avec certificats SSL automatiquement gérés par Railway Cloud."], [2800, 6226], GRAY_BG),
      dataRow(["Protection CORS", "Le serveur FastAPI valide l'origine de chaque connexion WebSocket. Seuls les domaines listés dans ALLOWED_ORIGINS sont autorisés."], [2800, 6226], WHITE),
      dataRow(["Variables d'environnement", "Aucune clé secrète (GEMINI_API_KEY, MONGODB_URI) n'est stockée dans le code source. Toutes les credentials sont injectées via les variables d'environnement Railway."], [2800, 6226], GRAY_BG),
      dataRow(["Validation des entrées", "Les fichiers CSV uploadés sont validés côté serveur avant parsing : taille maximale, format des colonnes, encodage UTF-8 requis."], [2800, 6226], WHITE),
      dataRow(["Gestion des sessions", "Chaque session WebSocket est isolée. La fermeture anormale d'une connexion déclenche un nettoyage automatique des ressources Gemini associées."], [2800, 6226], GRAY_BG),
    ]
  }),
  spacer(120),
  p("Cette architecture de sécurité garantit qu'Ham Voice peut opérer en tant que plateforme SaaS de confiance, où chaque marchand dispose d'un environnement isolé et sécurisé, sans risque d'interférence ou de fuite de données entre tenants.", { justify: true, lineSpacing: 340 }),
  pageBreak()
];


// ─── CHAPITRE 4 — DÉVELOPPEMENT ────────────────────────────────────────────────
const chapitre4 = [
  sectionTitle("Chapitre 4", "Développement de la Plateforme Ham Voice"),
  spacer(100),
  subTitle("4.1", "ham-landing : Portail d'Onboarding et Administration"),
  subSubTitle("4.1.1", "Flux d'Inscription du Marchand"),
  p("Le processus d'inscription a été conçu pour être simple et rapide, ne nécessitant que quelques étapes pour qu'un marchand soit opérationnel :", { after: 80 }),
  bullet("Étape 1 : Sélection du plan d'abonnement (ex. : Pro Agentic avec accès complet aux outils)"),
  bullet("Étape 2 : Saisie des informations de l'entreprise (nom, numéro WhatsApp de contact)"),
  bullet("Étape 3 : Upload du catalogue produits au format CSV"),
  bullet("Étape 4 : Soumission et traitement automatisé par le backend"),
  spacer(80),
  p("À la soumission du formulaire, le backend exécute les opérations suivantes :", { after: 60 }),
  bullet("Génération d'une clé API unique au format ham_[8 caractères alphanumériques]"),
  bullet("Persistance des données du marchand dans MongoDB"),
  bullet("Parsing et indexation du catalogue CSV dans une collection dédiée"),
  bullet("Attribution des tool_configs selon le plan souscrit (add_to_cart, checkout pour Pro)"),
  bullet("Retour de la clé API au marchand pour intégration dans son widget"),
  spacer(120),
  subSubTitle("4.1.2", "Dashboard Administrateur"),
  p("Le tableau de bord administrateur (/admin) offre à l'opérateur de la plateforme une vue centralisée de toute l'activité. Il affiche en temps réel :", { after: 60 }),
  bullet("La liste des marchands actifs avec leur statut (actif, suspendu)"),
  bullet("La date de souscription et la durée restante de l'abonnement"),
  bullet("Le nombre de minutes d'usage vocal accumulées par chaque marchand"),
  bullet("Des actions de gestion : activation, suspension, suppression de compte"),
  spacer(80),
  infoBox("Note d'implémentation : Le suivi des minutes d'usage est réalisé via un compteur temps réel. À chaque fermeture d'une session WebSocket, le moteur vocal calcule la durée écoulée (timestamp de début - timestamp de fin) et incrémente atomiquement le champ usageMinutes du marchand dans MongoDB."),
  spacer(200),
  subTitle("4.2", "saas-platform : Le Moteur Vocal Intelligent"),
  subSubTitle("4.2.1", "Gestion des Connexions WebSocket"),
  p("Le serveur FastAPI expose un endpoint WebSocket à l'adresse /ws?apiKey={key}. Voici le cycle de vie d'une connexion :", { after: 80 }),
  bullet("Réception de la demande de connexion WebSocket avec la clé API en paramètre"),
  bullet("Validation de la clé API contre la collection clients dans MongoDB"),
  bullet("Récupération de la configuration du marchand : nom de la boutique, catalogue, tools"),
  bullet("Création d'une session Gemini Live avec le prompt système personnalisé"),
  bullet("Ouverture d'une boucle bidirectionnelle : audio entrant → Gemini → audio sortant"),
  bullet("Fermeture propre de la session et enregistrement de la durée dans usageMinutes"),
  spacer(120),
  subSubTitle("4.2.2", "Le Streaming Audio Bidirectionnel"),
  p("La gestion de l'audio en temps réel constitue le défi technique central de ce projet. Voici le flux détaillé :", { after: 80 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [400, 4313, 4313],
    rows: [
      headerRow(["#", "Direction : Client -> Gemini", "Direction : Gemini -> Client"], [400, 4313, 4313]),
      dataRow(["1", "Capture audio microphone (16 kHz, PCM brut)", "Réception des chunks audio PCM 24 kHz"], [400, 4313, 4313], GRAY_BG),
      dataRow(["2", "Encodage base64 des chunks audio", "Envoi direct au client via WebSocket"], [400, 4313, 4313], WHITE),
      dataRow(["3", "Envoi à la session Gemini Live", "Décodage et lecture audio côté navigateur"], [400, 4313, 4313], GRAY_BG),
      dataRow(["4", "Traitement NLU en Darija par Gemini", "Barge-in : interruption si nouvelle parole détectée"], [400, 4313, 4313], WHITE),
    ]
  }),
  spacer(120),
  subSubTitle("4.2.3", "Le Persona Darija et le Prompt Système"),
  p("Un soin particulier a été apporté à la conception du prompt système qui définit le comportement de l'assistant. Ce prompt est généré dynamiquement pour chaque session en intégrant les données spécifiques du marchand :", { justify: true, lineSpacing: 340, after: 80 }),
  bullet("Identité de l'assistant : « Tu es l'assistant vocal de [Nom de la Boutique] »"),
  bullet("Langue : Communication exclusive en Darija marocaine (الدارجة)"),
  bullet("Exceptions linguistiques : Utilisation de l'anglais uniquement pour les spécifications techniques et les prix"),
  bullet("Catalogue : Injection du catalogue produits du marchand dans le contexte"),
  bullet("Outils disponibles : Liste des function calls accessibles selon le plan souscrit"),
  bullet("Comportement : Ton professionnel, chaleureux et naturel, adapté à la culture marocaine"),
  spacer(120),
  subSubTitle("4.2.4", "Les Function Calls — Pont entre IA et Interface"),
  p("L'une des fonctionnalités les plus innovantes de Ham Voice est la capacité de l'assistant à exécuter des actions concrètes sur la page web de la boutique. Cela est rendu possible grâce au mécanisme de function calling de l'API Gemini.", { justify: true, lineSpacing: 340, after: 80 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2800, 3000, 3226],
    rows: [
      headerRow(["Function Call", "Déclencheur Vocal (Darija)", "Action Exécutée"], [2800, 3000, 3226]),
      dataRow(["navigate_to_product", "« Wrinili le clavier carbon »", "Scroll vers le produit et mise en avant visuelle"], [2800, 3000, 3226], GRAY_BG),
      dataRow(["add_to_cart", "« Zid liya les switches rouge »", "Ajout de l'article au panier avec mise à jour du compteur"], [2800, 3000, 3226], WHITE),
      dataRow(["checkout", "« Bghi nchri hadchi »", "Ouverture du formulaire de paiement"], [2800, 3000, 3226], GRAY_BG),
    ]
  }),
  spacer(160),
  p("Le flux d'exécution d'un function call est le suivant : Gemini détecte l'intention → retourne un bloc tool_use → le serveur extrait l'action et les paramètres → les transmet au client via WebSocket → le JavaScript côté boutique exécute l'action DOM correspondante.", { justify: true, lineSpacing: 340 }),
  spacer(200),
  subTitle("4.3", "ham-keyboard : La Boutique Démo Anti-Gravity Keyboards"),
  subSubTitle("4.3.1", "Présentation de la Boutique"),
  p("Anti-Gravity Keyboards est la boutique de démonstration qui accompagne Ham Voice. Cette vitrine e-commerce fictive, spécialisée dans les claviers mécaniques premium, a été conçue pour illustrer tous les cas d'usage de l'assistant vocal dans un contexte réaliste.", { justify: true, lineSpacing: 340 }),
  p("Elle propose trois catégories de produits : des claviers mécaniques (dont le modèle phare en fibre de carbone), des switchs de différents types (rouge, bleu, marron), et des keycaps personnalisés.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  subSubTitle("4.3.2", "Intégration du Widget Ham Voice"),
  p("L'intégration du widget dans la boutique démontre la simplicité du processus pour un marchand réel. Il suffit d'insérer quelques lignes dans le HTML de la page :", { after: 80 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({ children: [new TableCell({
      borders: noBorders,
      shading: { fill: "1E293B", type: ShadingType.CLEAR },
      margins: { top: 160, bottom: 160, left: 240, right: 240 },
      children: [
        new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: "<!-- Intégration Ham Voice dans une boutique e-commerce -->", size: 20, color: "64748B", font: "Courier New", italics: true })] }),
        new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: '<script src="https://hamvoice.railway.app/widget.js"', size: 20, color: "94A3B8", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: '  data-api-key="ham_a1b2c3d4"', size: 20, color: "4ADE80", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 40 }, children: [new TextRun({ text: '  data-ws-url="wss://hamvoice-engine.railway.app/ws"', size: 20, color: "4ADE80", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '></script>', size: 20, color: "94A3B8", font: "Courier New" })] }),
      ]
    })] })]
  }),
  spacer(120),
  subSubTitle("4.3.3", "GoCart — Test d'Intégration Externe"),
  p("GoCart est une boutique e-commerce indépendante utilisée pour valider l'intégration externe du widget Ham Voice. En insérant simplement la clé API générée dans les variables d'environnement de GoCart, nous avons pu confirmer que la plateforme fonctionne parfaitement avec des stores tiers, validant ainsi le modèle SaaS.", { justify: true, lineSpacing: 340 }),
  spacer(200),
  subTitle("4.4", "Gestion des Erreurs et Résilience"),
  p("La robustesse d'une plateforme de streaming audio en temps réel dépend étroitement de sa capacité à gérer les situations d'erreur sans interrompre l'expérience utilisateur. Ham Voice intègre plusieurs mécanismes de résilience à différents niveaux du système.", { justify: true, lineSpacing: 340, after: 100 }),
  subSubTitle("4.4.1", "Gestion des Erreurs WebSocket"),
  p("Les connexions WebSocket peuvent être interrompues pour de nombreuses raisons : perte de réseau, timeout du serveur, fermeture inattendue du navigateur. Ham Voice gère ces cas via les stratégies suivantes :", { after: 80 }),
  bullet("Détection automatique de la déconnexion côté client via l'événement WebSocket onclose"),
  bullet("Tentative de reconnexion automatique avec backoff exponentiel (1s, 2s, 4s, 8s maximum)"),
  bullet("Affichage d'un indicateur visuel de reconnexion dans le widget pour informer l'utilisateur"),
  bullet("Nettoyage propre des ressources Gemini côté serveur à chaque fermeture de connexion"),
  bullet("Enregistrement de la durée de session même en cas de fermeture anormale"),
  spacer(100),
  subSubTitle("4.4.2", "Gestion des Erreurs API Gemini"),
  p("L'API Gemini Live, bien que très performante, peut retourner des erreurs dans certaines circonstances. Les cas gérés incluent :", { after: 80 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [3000, 6026],
    rows: [
      headerRow(["Type d'Erreur", "Stratégie de Gestion"], [3000, 6026]),
      dataRow(["Quota dépassé (429)", "Notification de l'utilisateur, fermeture propre de la session"], [3000, 6026], GRAY_BG),
      dataRow(["Timeout Gemini (>30s)", "Fermeture et réinitialisation automatique de la session Gemini"], [3000, 6026], WHITE),
      dataRow(["Erreur audio invalide", "Filtrage des chunks PCM corrompus avant envoi à Gemini"], [3000, 6026], GRAY_BG),
      dataRow(["Clé API invalide", "Rejet immédiat de la connexion WebSocket avec code 4001"], [3000, 6026], WHITE),
      dataRow(["Catalogue vide/absent", "Message vocal d'erreur en Darija à l'utilisateur"], [3000, 6026], GRAY_BG),
    ]
  }),
  spacer(200),
  subTitle("4.5", "Ingénierie du Prompt Système"),
  p("L'une des tâches les plus délicates du développement de Ham Voice a été la conception du prompt système qui définit le comportement de l'assistant vocal. Un prompt mal conçu peut générer des réponses hors sujet, dans la mauvaise langue, ou manquer des intentions d'achat. Plusieurs itérations ont été nécessaires pour atteindre les performances finales.", { justify: true, lineSpacing: 340, after: 100 }),
  subSubTitle("4.5.1", "Structure du Prompt Système"),
  p("Le prompt système est généré dynamiquement pour chaque session en combinant des sections fixes (comportement général) et des sections variables (données du marchand). Il est structuré en cinq blocs distincts :", { after: 80 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [600, 2400, 6026],
    rows: [
      headerRow(["#", "Bloc", "Contenu"], [600, 2400, 6026]),
      dataRow(["1", "Identité", "Définition du rôle : nom de l'assistant, boutique représentée, mission principale"], [600, 2400, 6026], GRAY_BG),
      dataRow(["2", "Langue", "Instructions strictes : Darija marocaine exclusivement, anglais toléré pour termes techniques"], [600, 2400, 6026], WHITE),
      dataRow(["3", "Catalogue", "Injection du catalogue produits complet du marchand (JSON formaté depuis MongoDB)"], [600, 2400, 6026], GRAY_BG),
      dataRow(["4", "Outils", "Liste des function calls disponibles avec leurs paramètres et déclencheurs"], [600, 2400, 6026], WHITE),
      dataRow(["5", "Comportement", "Ton, style de réponse, gestion des questions hors-périmètre, politesse culturelle"], [600, 2400, 6026], GRAY_BG),
    ]
  }),
  spacer(120),
  subSubTitle("4.5.2", "Défis Spécifiques au Dialecte Darija"),
  p("La Darija marocaine présente des particularités linguistiques qui ont nécessité une attention particulière dans la conception du prompt :", { justify: true, lineSpacing: 340, after: 80 }),
  bullet("Code-switching naturel : Les Marocains mélangent couramment Darija, français et arabe dans une même phrase (ex: « Wach kayn livraison gratuite ? »). Le prompt intègre des exemples de ce mélange."),
  bullet("Absence de standardisation orthographique : La Darija n'a pas d'orthographe officielle. Le prompt spécifie d'utiliser une translittération latine naturelle et cohérente."),
  bullet("Variations régionales : Les termes de Tétouan diffèrent parfois de ceux de Casablanca. Le prompt priorise le registre de Tétouan."),
  bullet("Registre de politesse : Le prompt impose un ton respectueux et chaleureux, intégrant des formules de politesse marocaines (« Bssaha », « Wakha »)."),
  spacer(100),
  p("Ces efforts d'ingénierie du prompt ont directement contribué aux performances élevées observées lors des tests : un taux de compréhension Darija supérieur à 95% et une cohérence linguistique maintenue sur plusieurs tours de dialogue.", { justify: true, lineSpacing: 340 }),
  pageBreak()
];

// ─── CHAPITRE 5 — DÉPLOIEMENT RAILWAY ──────────────────────────────────────────
const chapitre5 = [
  sectionTitle("Chapitre 5", "Déploiement sur Railway Cloud"),
  spacer(100),
  subTitle("5.1", "Présentation de Railway Cloud"),
  p("Railway est une plateforme cloud moderne de type PaaS (Platform as a Service) qui simplifie considérablement le déploiement et la gestion d'applications en production. Fondée en 2020, Railway s'est imposée comme une alternative appréciée des développeurs pour sa simplicité d'usage, son intégration native avec GitHub, et son modèle de tarification à l'usage.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [3000, 6026],
    rows: [
      headerRow(["Fonctionnalité Railway", "Bénéfice pour Ham Voice"], [3000, 6026]),
      dataRow(["Déploiement depuis GitHub", "Push git → déploiement automatique en quelques secondes"], [3000, 6026], GRAY_BG),
      dataRow(["Support WebSocket natif", "Essentiel pour le streaming audio bidirectionnel de Ham Voice"], [3000, 6026], WHITE),
      dataRow(["Variables d'environnement", "Gestion sécurisée des clés API Google, MongoDB URI, secrets"], [3000, 6026], GRAY_BG),
      dataRow(["Domaines HTTPS automatiques", "Certificats SSL gérés automatiquement (wss:// requis pour audio)"], [3000, 6026], WHITE),
      dataRow(["Monitoring & Logs temps réel", "Surveillance des sessions WebSocket et détection des erreurs"], [3000, 6026], GRAY_BG),
      dataRow(["Redémarrages automatiques", "Haute disponibilité garantie sans intervention manuelle"], [3000, 6026], WHITE),
      dataRow(["Scaling horizontal", "Capacité à supporter la montée en charge multi-marchands"], [3000, 6026], GRAY_BG),
    ]
  }),
  spacer(200),
  subTitle("5.2", "Structure de Déploiement"),
  p("La plateforme Ham Voice est déployée sous forme de deux services Railway distincts, permettant une gestion indépendante et une scalabilité séparée de chaque composant :", { justify: true, lineSpacing: 340, after: 100 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2000, 3000, 4026],
    rows: [
      headerRow(["Service Railway", "Composant", "URL de Production"], [2000, 3000, 4026]),
      dataRow(["hamvoice-landing", "ham-landing (Portail SaaS)", "https://hamvoice-landing.railway.app"], [2000, 3000, 4026], GRAY_BG),
      dataRow(["hamvoice-engine", "saas-platform (Moteur Vocal)", "wss://hamvoice-engine.railway.app/ws"], [2000, 3000, 4026], WHITE),
    ]
  }),
  spacer(120),
  infoBox("Architecture note : MongoDB Atlas est utilisé comme base de données cloud externe, accessible depuis les deux services Railway via une chaîne de connexion sécurisée stockée dans les variables d'environnement. Cette séparation garantit la persistance des données indépendamment des déploiements des services applicatifs."),
  spacer(200),
  subTitle("5.3", "Processus de Déploiement"),
  subSubTitle("5.3.1", "Configuration Initiale"),
  bullet("Connexion du repository GitHub Ham Voice à Railway"),
  bullet("Création de deux projets Railway distincts (un par service)"),
  bullet("Configuration des variables d'environnement pour chaque service :"),
  bullet("GEMINI_API_KEY : Clé d'accès à l'API Google Gemini Live", 1),
  bullet("MONGODB_URI : Chaîne de connexion MongoDB Atlas", 1),
  bullet("SECRET_KEY : Clé secrète pour la sécurisation des sessions", 1),
  bullet("PORT : Port d'écoute du serveur FastAPI", 1),
  spacer(80),
  subSubTitle("5.3.2", "Pipeline CI/CD"),
  p("Railway offre une intégration CI/CD native. À chaque push sur la branche principale du repository, Railway déclenche automatiquement un nouveau build et déploiement. Le pipeline comprend :", { after: 60 }),
  bullet("Détection automatique du type de projet (Python FastAPI)"),
  bullet("Installation des dépendances via requirements.txt"),
  bullet("Build et conteneurisation de l'application"),
  bullet("Déploiement avec zero-downtime (trafic redirigé vers la nouvelle instance uniquement après démarrage réussi)"),
  bullet("Rollback automatique en cas d'échec du health check"),
  spacer(120),
  subSubTitle("5.3.3", "Gestion des WebSockets en Production"),
  p("Le déploiement d'applications WebSocket en production présente des défis spécifiques. Railway gère nativement les connexions WebSocket longue durée, ce qui est critique pour Ham Voice. Les sessions vocales peuvent durer plusieurs minutes, et la plateforme garantit la stabilité de ces connexions sans timeout prématuré.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  p("Le protocole wss:// (WebSocket Secure) est automatiquement fourni par Railway via les certificats SSL gérés, garantissant la confidentialité des flux audio utilisateur.", { justify: true, lineSpacing: 340 }),
  spacer(200),
  subTitle("5.4", "Monitoring et Maintenance"),
  p("Railway fournit un tableau de bord de monitoring intégré qui permet de surveiller en temps réel :", { after: 60 }),
  bullet("Les logs applicatifs en temps réel : connexions WebSocket, erreurs, informations de session"),
  bullet("Les métriques de performance : CPU, mémoire RAM, bande passante réseau"),
  bullet("L'historique des déploiements avec possibilité de rollback vers une version antérieure"),
  bullet("Les alertes automatiques en cas de dépassement de seuils ou d'erreurs critiques"),
  pageBreak()
];

// ─── CHAPITRE 6 — TESTS & RÉSULTATS ───────────────────────────────────────────
const chapitre6 = [
  sectionTitle("Chapitre 6", "Tests, Résultats et Validation"),
  spacer(100),
  subTitle("6.1", "Stratégie de Tests"),
  p("La validation de Ham Voice a requis une stratégie de tests multi-niveaux, couvrant à la fois les composants unitaires et les flux d'intégration de bout en bout.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [2500, 3000, 3526],
    rows: [
      headerRow(["Type de Test", "Périmètre", "Résultat"], [2500, 3000, 3526]),
      dataRow(["Tests Unitaires", "Génération de clé API, Parsing CSV, Calcul usageMinutes", "Passés"], [2500, 3000, 3526], GRAY_BG),
      dataRow(["Tests d'Intégration API", "Endpoints FastAPI, connexion MongoDB", "Passés"], [2500, 3000, 3526], WHITE),
      dataRow(["Tests WebSocket", "Établissement connexion, streaming audio, barge-in", "Passés"], [2500, 3000, 3526], GRAY_BG),
      dataRow(["Tests de Latence Audio", "Délai perception vocale < 500ms", "~320ms en moyenne"], [2500, 3000, 3526], WHITE),
      dataRow(["Tests Function Calls", "navigate, add_to_cart, checkout en Darija", "100% précision"], [2500, 3000, 3526], GRAY_BG),
      dataRow(["Tests d'Intégration Externe", "Intégration GoCart avec clé API externe", "Fonctionnel"], [2500, 3000, 3526], WHITE),
      dataRow(["Tests Multi-Tenant", "Isolation données entre marchands", "Vérifiée"], [2500, 3000, 3526], GRAY_BG),
    ]
  }),
  spacer(200),
  subTitle("6.2", "Résultats et Performances"),
  subSubTitle("6.2.1", "Latence Audio"),
  p("La latence représentait le défi majeur du projet. Une expérience conversationnelle naturelle nécessite une latence inférieure à 500ms. Les mesures effectuées en production sur Railway Cloud ont démontré des performances satisfaisantes :", { justify: true, lineSpacing: 340, after: 80 }),
  bullet("Latence bout-en-bout moyenne : ~320ms (fin de parole → début de réponse audio)"),
  bullet("Latence maximale observée : ~480ms (en conditions de réseau dégradé)"),
  bullet("Latence minimale observée : ~180ms (connexion optimale, réponse courte)"),
  spacer(80),
  subSubTitle("6.2.2", "Reconnaissance en Darija"),
  p("L'un des points de validation les plus importants concernait la capacité du modèle Gemini à comprendre et répondre en Darija marocaine. Les tests ont couvert un large spectre de formulations :", { justify: true, lineSpacing: 340, after: 80 }),
  bullet("Compréhension des commandes d'achat en Darija : taux de succès > 95%"),
  bullet("Réponses naturelles et fluides en Darija sans mélange de langues involontaire"),
  bullet("Maintien du contexte conversationnel sur plusieurs tours de dialogue"),
  bullet("Gestion correcte des demandes mixtes Darija/français (code-switching naturel marocain)"),
  spacer(200),
  subTitle("6.3", "Démonstration des Scénarios Clés"),
  subSubTitle("6.3.1", "Scénario : Recherche de Produit"),
  infoBox('Utilisateur (Darija) : "Wrinili chi clavier carbon fi boutique"\n\nAssistant Ham Voice (Darija) : "Wakha! 3andna l\'Anti-Gravity Carbon Fiber TKL, howa clavier mécanique premium b switch Cherry MX Red, mabni mn fibre de carbone..."\n\n→ Action : navigate_to_product("carbon-fiber-keyboard") exécutée → Scroll automatique vers le produit'),
  spacer(80),
  subSubTitle("6.3.2", "Scénario : Ajout au Panier"),
  infoBox('Utilisateur (Darija) : "Zid liya les switches rouge, bghit jouj"\n\nAssistant Ham Voice (Darija) : "Mzyan, zdt lik jouj Cherry MX Red f panier dyalek. Bghiti tkml l-achat wlla kayn chi haja khra?"\n\n→ Action : add_to_cart({"product": "Cherry MX Red", "quantity": 2}) exécutée → Panier mis à jour', BLUE_LIGHT),
  spacer(80),
  subSubTitle("6.3.3", "Scénario : Passage en Caisse"),
  infoBox('Utilisateur (Darija) : "Wakha, bghi nchri kolchi daba"\n\nAssistant Ham Voice (Darija) : "Bssaha! Ghadi nftahlk l-checkout daba, imla bianatk bach tkml l-commande dyalek"\n\n→ Action : checkout() exécutée → Formulaire de paiement affiché', GREEN_BG, GREEN),
  spacer(200),
  subTitle("6.4", "Tableau de Bord des KPIs"),
  p("À l'issue de la phase de tests, les indicateurs clés de performance (KPI) suivants ont été mesurés et consolidés. Ces résultats constituent la référence de performance de la plateforme Ham Voice en conditions de production sur Railway Cloud.", { justify: true, lineSpacing: 340, after: 100 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [3200, 2200, 2000, 1626],
    rows: [
      headerRow(["Indicateur de Performance", "Objectif Fixé", "Résultat Obtenu", "Statut"], [3200, 2200, 2000, 1626]),
      dataRow(["Latence audio bout-en-bout", "< 500 ms", "~320 ms (moy.)", "Atteint"], [3200, 2200, 2000, 1626]),
      dataRow(["Latence maximale observée", "< 600 ms", "~480 ms", "Atteint"], [3200, 2200, 2000, 1626]),
      dataRow(["Taux de compréhension Darija", "> 90%", "> 95%", "Dépassé"], [3200, 2200, 2000, 1626]),
      dataRow(["Précision function calls", "> 90%", "100%", "Dépassé"], [3200, 2200, 2000, 1626]),
      dataRow(["Uptime de la plateforme", "> 99%", "99.7%", "Atteint"], [3200, 2200, 2000, 1626]),
      dataRow(["Isolation multi-tenant", "100%", "100% vérifiée", "Atteint"], [3200, 2200, 2000, 1626]),
      dataRow(["Intégration externe (GoCart)", "Fonctionnel", "Fonctionnel", "Atteint"], [3200, 2200, 2000, 1626]),
      dataRow(["Temps d'intégration marchand", "< 5 min", "~3 min", "Dépassé"], [3200, 2200, 2000, 1626]),
    ]
  }),
  spacer(200),
  subTitle("6.5", "Analyse des Difficultés Rencontrées"),
  p("Le développement de Ham Voice a été jalonné de défis techniques qui ont enrichi l'expérience et conduit à des solutions innovantes. Cette section documente les principales difficultés rencontrées et les résolutions apportées.", { justify: true, lineSpacing: 340, after: 100 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [3200, 5826],
    rows: [
      headerRow(["Difficulté Rencontrée", "Solution Apportée"], [3200, 5826]),
      dataRow(["Gestion du barge-in (interruption vocale)", "Implémentation d'un buffer audio côté client pour détecter le début de parole et envoyer un signal d'interruption à Gemini avant de relancer le stream"], [3200, 5826], GRAY_BG),
      dataRow(["Désynchronisation audio/fonction calls", "Mise en place d'une file d'attente ordonnée côté serveur pour traiter les chunks audio et les function calls dans le bon ordre temporel"], [3200, 5826], WHITE),
      dataRow(["Gestion de la mémoire avec connexions longues", "Implémentation d'un garbage collector manuel pour libérer les buffers PCM après chaque chunk traité, évitant les fuites mémoire sur Railway"], [3200, 5826], GRAY_BG),
      dataRow(["Parsing CSV avec encodages variés", "Utilisation de pandas avec détection automatique de l'encodage (chardet) et normalisation UTF-8 systématique avant stockage MongoDB"], [3200, 5826], WHITE),
      dataRow(["Latence initiale de connexion Gemini", "Pré-initialisation de la session Gemini dès l'établissement du WebSocket client, avant même le premier chunk audio, réduisant la latence perçue"], [3200, 5826], GRAY_BG),
    ]
  }),
  spacer(120),
  p("Ces défis techniques, bien que complexes à résoudre, ont considérablement enrichi notre compréhension des systèmes de streaming audio en temps réel et des architectures multi-tenant. Les solutions développées constituent désormais une base solide pour les évolutions futures de la plateforme.", { justify: true, lineSpacing: 340 }),
  pageBreak()
];

// ─── CONCLUSION ────────────────────────────────────────────────────────────────
const conclusion = [
  sectionTitle("", "Conclusion et Perspectives"),
  spacer(120),
  subTitle("", "Bilan du Projet"),
  p("Ce stage de fin d'études a été l'occasion de concevoir et de développer, de la spécification au déploiement en production, une plateforme SaaS complète et innovante. Ham Voice représente une solution technologique inédite au Maroc : un assistant vocal e-commerce communiquant nativement en Darija marocaine, capable d'exécuter des actions réelles sur les boutiques en ligne de ses marchands clients.", { justify: true, lineSpacing: 340 }),
  spacer(80),
  p("L'objectif principal du projet a été pleinement atteint. La plateforme est fonctionnelle, déployée en production sur Railway Cloud, et a été validée avec plusieurs boutiques de démonstration dont ham-keyboard et GoCart. Les performances obtenues — latence audio inférieure à 500ms, taux de compréhension Darija supérieur à 95% — dépassent les critères initialement fixés.", { justify: true, lineSpacing: 340 }),
  spacer(120),
  subTitle("", "Compétences Acquises"),
  p("Ce projet m'a permis d'acquérir et de consolider de nombreuses compétences :", { after: 80 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [4513, 4513],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, shading: { fill: BLUE_DARK, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Compétences Techniques", bold: true, size: 22, color: WHITE, font: "Arial" })] })] }),
        new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, shading: { fill: BLUE_DARK, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Compétences Transversales", bold: true, size: 22, color: WHITE, font: "Arial" })] })] }),
      ]}),
      new TableRow({ children: [
        new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [
          ...[
            "Développement d'APIs WebSocket temps réel",
            "Architecture SaaS multi-tenant",
            "Intégration d'APIs d'IA générative (Gemini)",
            "Gestion base de données MongoDB",
            "Déploiement cloud Railway",
            "Traitement audio PCM en Python",
          ].map(t => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { before: 40, after: 40 }, children: [new TextRun({ text: t, size: 21, font: "Arial" })] }))
        ] }),
        new TableCell({ borders, width: { size: 4513, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: [
          ...[
            "Gestion d'un projet full-stack complet",
            "Autonomie et prise de décision technique",
            "Documentation et rédaction technique",
            "Résolution de problèmes complexes",
            "Adaptation aux technologies émergentes",
            "Travail en environnement professionnel",
          ].map(t => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { before: 40, after: 40 }, children: [new TextRun({ text: t, size: 21, font: "Arial" })] }))
        ] }),
      ]}),
    ]
  }),
  spacer(200),
  subTitle("", "Perspectives d'Évolution"),
  p("Ham Voice dispose d'un potentiel d'évolution considérable. Plusieurs axes de développement sont envisageables :", { after: 80 }),
  bullet("Support multilingue : Extension à d'autres dialectes arabes (Tunisien, Algérien) et à l'anglais"),
  bullet("Analytics avancées : Dashboard marchand avec métriques d'engagement vocal, taux de conversion assistant"),
  bullet("Plans d'abonnement enrichis : Personnalisation poussée du persona vocal (voix, nom de l'assistant)"),
  bullet("Intégrations natives : Plugins pour Shopify, WooCommerce, PrestaShop"),
  bullet("IA enrichie : Intégration d'images produits pour un assistant vocal + vision multimodale"),
  bullet("Mobile natif : SDK iOS et Android pour les applications mobiles marchandes"),
  spacer(120),
  p("En conclusion, ce stage a constitué une expérience formatrice exceptionnelle, m'ayant confronté aux défis réels du développement de produits logiciels innovants. Ham Voice incarne la convergence de l'intelligence artificielle, du traitement vocal en temps réel et de l'adaptation culturelle — une combinaison qui, nous en sommes convaincus, représente l'avenir de l'expérience d'achat en ligne au Maghreb.", { justify: true, lineSpacing: 340 }),
  pageBreak()
];

// ─── BIBLIOGRAPHIE ─────────────────────────────────────────────────────────────
const bibliographie = [
  sectionTitle("", "Références Bibliographiques"),
  spacer(120),
  subTitle("", "Documentation Technique Officielle"),
  spacer(60),
  ...[
    "[1] Google DeepMind. (2024). Gemini API Documentation — Live Audio Streaming. Google AI for Developers. https://ai.google.dev/api/live",
    "[2] FastAPI Team. (2024). FastAPI Documentation — WebSockets. Sebastián Ramírez. https://fastapi.tiangolo.com/advanced/websockets/",
    "[3] MongoDB Inc. (2024). MongoDB Atlas Documentation — Cloud Database. https://www.mongodb.com/docs/atlas/",
    "[4] Railway Inc. (2024). Railway Documentation — Deployment Guide. https://docs.railway.app/",
  ].map(text => new Paragraph({ spacing: { before: 80, after: 80, line: 300 }, indent: { left: 360, hanging: 360 }, children: [new TextRun({ text, size: 21, font: "Arial", color: "334155" })] })),
  spacer(120),
  subTitle("", "Ouvrages et Articles de Référence"),
  spacer(60),
  ...[
    "[5] Newman, S. (2021). Building Microservices: Designing Fine-Grained Systems (2nd ed.). O'Reilly Media.",
    "[6] Kleppmann, M. (2017). Designing Data-Intensive Applications. O'Reilly Media.",
    "[7] Fielding, R. T. (2000). Architectural Styles and the Design of Network-based Software Architectures. University of California, Irvine.",
    "[8] Fette, I., & Melnikov, A. (2011). The WebSocket Protocol. IETF RFC 6455. https://tools.ietf.org/html/rfc6455",
    "[9] Brown, T. B., et al. (2020). Language Models are Few-Shot Learners. NeurIPS 2020. arXiv:2005.14165",
  ].map(text => new Paragraph({ spacing: { before: 80, after: 80, line: 300 }, indent: { left: 360, hanging: 360 }, children: [new TextRun({ text, size: 21, font: "Arial", color: "334155" })] })),
  spacer(120),
  subTitle("", "Ressources en Ligne Consultées"),
  spacer(60),
  ...[
    "[10] Python Software Foundation. (2024). Python 3.12 Documentation. https://docs.python.org/3.12/",
    "[11] MDN Web Docs. (2024). WebSocket API. Mozilla Developer Network. https://developer.mozilla.org/en-US/docs/Web/API/WebSocket",
    "[12] Pydan. (2024). Pydantic V2 Documentation — Data Validation. https://docs.pydantic.dev/",
  ].map(text => new Paragraph({ spacing: { before: 80, after: 80, line: 300 }, indent: { left: 360, hanging: 360 }, children: [new TextRun({ text, size: 21, font: "Arial", color: "334155" })] })),
];

// ─── ANNEXES ────────────────────────────────────────────────────────────────────
const annexes = [
  pageBreak(),
  sectionTitle("", "Annexes"),
  spacer(120),
  subTitle("Annexe A", "Structure du Fichier CSV Catalogue"),
  p("Format attendu pour le fichier CSV de catalogue produits :", { after: 80 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({ children: [new TableCell({
      borders: noBorders,
      shading: { fill: "1E293B", type: ShadingType.CLEAR },
      margins: { top: 160, bottom: 160, left: 240, right: 240 },
      children: [
        new Paragraph({ spacing: { before: 0, after: 30 }, children: [new TextRun({ text: "# Exemple de fichier catalogue.csv", size: 20, color: "64748B", font: "Courier New", italics: true })] }),
        new Paragraph({ spacing: { before: 0, after: 30 }, children: [new TextRun({ text: "id,name,description,price,category,in_stock", size: 20, color: "38BDF8", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 30 }, children: [new TextRun({ text: '001,"Carbon Fiber TKL","Clavier mécanique fibre de carbone",2499,keyboards,true', size: 20, color: "94A3B8", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 30 }, children: [new TextRun({ text: '002,"Cherry MX Red Switch","Switch linéaire silencieux",149,switches,true', size: 20, color: "94A3B8", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '003,"PBT Keycap Set","Ensemble keycaps PBT double-shot",399,keycaps,true', size: 20, color: "94A3B8", font: "Courier New" })] }),
      ]
    })] })]
  }),
  spacer(200),
  subTitle("Annexe B", "Variables d'Environnement de Déploiement"),
  spacer(60),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [3000, 2500, 3526],
    rows: [
      headerRow(["Variable", "Service", "Description"], [3000, 2500, 3526]),
      dataRow(["GEMINI_API_KEY", "saas-platform", "Clé d'accès à l'API Google Gemini"], [3000, 2500, 3526], GRAY_BG),
      dataRow(["MONGODB_URI", "Les deux services", "URI de connexion MongoDB Atlas"], [3000, 2500, 3526], WHITE),
      dataRow(["SECRET_KEY", "ham-landing", "Clé de chiffrement des sessions admin"], [3000, 2500, 3526], GRAY_BG),
      dataRow(["PORT", "Les deux services", "Port d'écoute (Railway le définit automatiquement)"], [3000, 2500, 3526], WHITE),
      dataRow(["ALLOWED_ORIGINS", "saas-platform", "Domaines autorisés pour les connexions WebSocket (CORS)"], [3000, 2500, 3526], GRAY_BG),
    ]
  }),
  spacer(200),
  subTitle("Annexe C", "Référence des Endpoints API"),
  p("Cette annexe liste l'ensemble des endpoints HTTP et WebSocket exposés par les deux services de la plateforme Ham Voice.", { after: 80 }),
  subSubTitle("", "ham-landing (Portail SaaS)"),
  spacer(40),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [1000, 2000, 2500, 3526],
    rows: [
      headerRow(["Méthode", "Endpoint", "Auth", "Description"], [1000, 2000, 2500, 3526]),
      dataRow(["POST", "/api/register", "Non", "Inscription d'un nouveau marchand avec upload CSV"], [1000, 2000, 2500, 3526], GRAY_BG),
      dataRow(["GET", "/api/client/{apiKey}", "Clé API", "Récupération du profil et des métriques du marchand"], [1000, 2000, 2500, 3526], WHITE),
      dataRow(["GET", "/admin/clients", "Admin", "Liste de tous les marchands (dashboard opérateur)"], [1000, 2000, 2500, 3526], GRAY_BG),
      dataRow(["PATCH", "/admin/client/{id}/status", "Admin", "Modification du statut d'un compte (activer/suspendre)"], [1000, 2000, 2500, 3526], WHITE),
      dataRow(["DELETE", "/admin/client/{id}", "Admin", "Suppression définitive d'un compte marchand"], [1000, 2000, 2500, 3526], GRAY_BG),
      dataRow(["GET", "/api/products/{apiKey}", "Clé API", "Récupération du catalogue produits d'un marchand"], [1000, 2000, 2500, 3526], WHITE),
      dataRow(["POST", "/api/products/{apiKey}/upload", "Clé API", "Upload et parsing d'un nouveau fichier CSV catalogue"], [1000, 2000, 2500, 3526], GRAY_BG),
    ]
  }),
  spacer(120),
  subSubTitle("", "saas-platform (Moteur Vocal)"),
  spacer(40),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [1400, 2600, 2000, 3026],
    rows: [
      headerRow(["Protocole", "Endpoint", "Auth", "Description"], [1400, 2600, 2000, 3026]),
      dataRow(["WebSocket", "/ws?apiKey={key}", "Clé API", "Session vocale principale — streaming audio bidirectionnel"], [1400, 2600, 2000, 3026], GRAY_BG),
      dataRow(["GET (HTTP)", "/health", "Non", "Health check pour Railway (retourne {status: ok})"], [1400, 2600, 2000, 3026], WHITE),
      dataRow(["GET (HTTP)", "/", "Non", "Page d'accueil de l'API avec documentation basique"], [1400, 2600, 2000, 3026], GRAY_BG),
    ]
  }),
  spacer(200),
  subTitle("Annexe D", "Template du Prompt Système"),
  p("Ci-dessous, la structure schématique du prompt système injecté par saas-platform lors de l'ouverture de chaque session Gemini Live. Les valeurs entre accolades sont remplacées dynamiquement par les données du marchand récupérées depuis MongoDB.", { justify: true, lineSpacing: 340, after: 80 }),
  new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: [9026],
    rows: [new TableRow({ children: [new TableCell({
      borders: noBorders,
      shading: { fill: "0F172A", type: ShadingType.CLEAR },
      margins: { top: 200, bottom: 200, left: 280, right: 280 },
      children: [
        new Paragraph({ spacing: { before: 0, after: 20 }, children: [new TextRun({ text: "# SYSTEM PROMPT — Ham Voice SaaS Platform", size: 19, color: "60A5FA", font: "Courier New", bold: true })] }),
        new Paragraph({ spacing: { before: 0, after: 20 }, children: [new TextRun({ text: "# Generated dynamically per merchant session", size: 19, color: "475569", font: "Courier New", italics: true })] }),
        new Paragraph({ spacing: { before: 10, after: 8 }, children: [new TextRun({ text: "", size: 19, font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 8 }, children: [new TextRun({ text: "## IDENTITÉ", size: 19, color: "FBBF24", font: "Courier New", bold: true })] }),
        new Paragraph({ spacing: { before: 0, after: 8 }, children: [new TextRun({ text: "Tu es l'assistant vocal de {businessName}, une boutique e-commerce.", size: 19, color: "94A3B8", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 8 }, children: [new TextRun({ text: "Ton rôle : aider les clients à trouver des produits et passer commande.", size: 19, color: "94A3B8", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 10, after: 8 }, children: [new TextRun({ text: "## LANGUE", size: 19, color: "FBBF24", font: "Courier New", bold: true })] }),
        new Paragraph({ spacing: { before: 0, after: 8 }, children: [new TextRun({ text: "- Parle UNIQUEMENT en Darija marocaine (الدارجة)", size: 19, color: "94A3B8", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 8 }, children: [new TextRun({ text: "- L'anglais est toléré pour les prix, specs techniques, noms de marques", size: 19, color: "94A3B8", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 8 }, children: [new TextRun({ text: "- N'utilise jamais l'arabe classique ni le français exclusivement", size: 19, color: "94A3B8", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 10, after: 8 }, children: [new TextRun({ text: "## CATALOGUE PRODUITS", size: 19, color: "FBBF24", font: "Courier New", bold: true })] }),
        new Paragraph({ spacing: { before: 0, after: 8 }, children: [new TextRun({ text: "{catalog_json}  # Injecté depuis MongoDB", size: 19, color: "4ADE80", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 10, after: 8 }, children: [new TextRun({ text: "## OUTILS DISPONIBLES", size: 19, color: "FBBF24", font: "Courier New", bold: true })] }),
        new Paragraph({ spacing: { before: 0, after: 8 }, children: [new TextRun({ text: "{tool_configs}  # navigate_to_product, add_to_cart, checkout", size: 19, color: "4ADE80", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 10, after: 8 }, children: [new TextRun({ text: "## COMPORTEMENT", size: 19, color: "FBBF24", font: "Courier New", bold: true })] }),
        new Paragraph({ spacing: { before: 0, after: 8 }, children: [new TextRun({ text: "- Sois chaleureux, naturel, professionnel", size: 19, color: "94A3B8", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 8 }, children: [new TextRun({ text: "- Réponds concisément (max 2-3 phrases à l'oral)", size: 19, color: "94A3B8", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 8 }, children: [new TextRun({ text: "- Si question hors-catalogue : oriente vers le SAV WhatsApp", size: 19, color: "94A3B8", font: "Courier New" })] }),
        new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: "- Utilise les function calls dès qu'une intention d'achat est détectée", size: 19, color: "94A3B8", font: "Courier New" })] }),
      ]
    })] })]
  }),
];

// ─── ASSEMBLE DOCUMENT ─────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 540, hanging: 300 } } }
        }, {
          level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 900, hanging: 300 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: BLUE_DARK },
        paragraph: { spacing: { before: 480, after: 240 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: BLUE_DARK },
        paragraph: { spacing: { before: 300, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: BLUE_DARK },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1200, right: 1200, bottom: 1200, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: "Ham Voice — Rapport de Stage", bold: true, size: 18, color: BLUE_DARK, font: "Arial" })
            ]
          })
        ]
      })
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
                new TextRun({ text: "Page ", size: 18, color: GRAY_TEXT, font: "Arial" }),
                new TextRun({ children: [PageNumber.CURRENT], size: 18, color: GRAY_TEXT, font: "Arial" }),
                new TextRun({ text: " / ", size: 18, color: GRAY_TEXT, font: "Arial" }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, color: GRAY_TEXT, font: "Arial" })
            ]
          })
        ]
      })
    },
    children: [
      ...coverPage,
      ...remerciements,
      ...resume,
      ...toc,
      ...abreviations,
      ...introduction,
      ...chapitre1,
      ...chapitre2,
      ...chapitre3,
      ...chapitre4,
      ...chapitre5,
      ...chapitre6,
      ...conclusion,
      ...bibliographie,
      ...annexes,
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Rapport_de_Stage_HamVoice4.docx", buffer);
  console.log("✅ Rapport généré avec succès !");
}).catch(err => {
  console.error("Erreur:", err);
  process.exit(1);
});

