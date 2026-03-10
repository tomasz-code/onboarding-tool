const express = require('express');
const cors    = require('cors');
const Anthropic = require('@anthropic-ai/sdk');
const { google } = require('googleapis');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
        BorderStyle, Table, TableRow, TableCell, WidthType, ShadingType } = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '2mb' }));

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ─── Google Drive Auth ────────────────────────────────────────────────────────
let drive = null;
const DRIVE_FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID;

if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) {
  try {
    const auth = new google.auth.GoogleAuth({
      credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
      scopes: ['https://www.googleapis.com/auth/drive.file'],
    });
    drive = google.drive({ version: 'v3', auth });
  } catch (e) {
    console.warn('Google Drive nicht konfiguriert:', e.message);
  }
}

// ─── /chat  ───────────────────────────────────────────────────────────────────
app.post('/chat', async (req, res) => {
  try {
    const { messages, system } = req.body;

    // Wenn keine Nachrichten → Eröffnungsnachricht holen
    const msgs = messages.length === 0
      ? [{ role: 'user', content: 'Bitte beginne das Onboarding.' }]
      : messages;

    const response = await anthropic.messages.create({
      model: 'claude-opus-4-5',
      max_tokens: 1024,
      system,
      messages: msgs,
    });

    res.json(response);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ─── /generate  ──────────────────────────────────────────────────────────────
app.post('/generate', async (req, res) => {
  try {
    const { messages } = req.body;

    // Schritt 1: Worksheet-Daten aus Gespräch extrahieren
    const extractionPrompt = `Du bekommst ein Onboarding-Gespräch zwischen einem Marketing-Berater und einem Kunden.
Extrahiere daraus alle relevanten Informationen und gib sie als JSON zurück.
Antworte NUR mit dem JSON-Objekt, ohne Erklärungen, ohne Markdown-Backticks.

JSON-Struktur:
{
  "werteversprechen": "",
  "usp": "",
  "zielgruppe_name": "",
  "zielgruppe_alter": "",
  "zielgruppe_geschlecht": "",
  "zielgruppe_wohnort": "",
  "zielgruppe_sonstiges": "",
  "probleme": ["", "", ""],
  "unangenehme_situationen": ["", "", ""],
  "nachts_wach": ["", "", ""],
  "konsequenzen": ["", "", ""],
  "werte_zielgruppe": ["", "", ""],
  "wunschleben": ["", "", ""],
  "insgeheime_wuensche": ["", "", ""],
  "kernwuensche": ["", "", ""],
  "produkt_name": "",
  "produkt_preis": "",
  "hauptbenefit": "",
  "features": ["", "", ""],
  "einwaende": ["", "", ""],
  "schutz_mechanismus": "",
  "barriere": "",
  "positionierung": "",
  "kampagnen_kriterien": ["", "", ""],
  "tonalitaet": "",
  "sprache_form": "",
  "no_gos": ["", "", ""]
}`;

    const extractionResponse = await anthropic.messages.create({
      model: 'claude-opus-4-5',
      max_tokens: 2048,
      messages: [
        ...messages,
        {
          role: 'user',
          content: extractionPrompt
        }
      ]
    });

    let worksheetData;
    try {
      let raw = extractionResponse.content[0].text.trim();
      raw = raw.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/i, '').trim();
      worksheetData = JSON.parse(raw);
    } catch {
      return res.status(500).json({ error: 'JSON-Parsing fehlgeschlagen', raw: extractionResponse.content[0].text });
    }

    // Schritt 2: .docx generieren
    const docBuffer = await generateDocx(worksheetData);

    // Schritt 3: In Google Drive hochladen
    const fileName = `Onboarding_${worksheetData.zielgruppe_name || 'Kunde'}_${new Date().toISOString().slice(0,10)}.docx`;

    if (!drive) {
      return res.status(500).json({ error: 'Google Drive nicht konfiguriert.' });
    }

    const { Readable } = require('stream');
    const stream = Readable.from(docBuffer);

    await drive.files.create({
      requestBody: {
        name: fileName,
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        parents: [DRIVE_FOLDER_ID],
      },
      media: {
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        body: stream,
      },
    });

    res.json({ success: true, fileName });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ─── DOCX Generator ───────────────────────────────────────────────────────────
async function generateDocx(d) {
  const gold  = 'C9A96E';
  const dark  = '1A1A2E';
  const light = 'F7F5F2';
  const gray  = '8A8A9A';

  function heading(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 120 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: gold } },
      children: [new TextRun({ text, font: 'Georgia', size: 26, bold: true, color: dark })]
    });
  }

  function subheading(text) {
    return new Paragraph({
      spacing: { before: 240, after: 80 },
      children: [new TextRun({ text, font: 'DM Sans', size: 22, bold: true, color: dark })]
    });
  }

  function body(text) {
    return new Paragraph({
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text: text || '—', font: 'DM Sans', size: 20, color: text ? dark : gray })]
    });
  }

  function labelValue(label, value) {
    return new Paragraph({
      spacing: { before: 80, after: 60 },
      children: [
        new TextRun({ text: label + ': ', font: 'DM Sans', size: 20, bold: true, color: dark }),
        new TextRun({ text: value || '—', font: 'DM Sans', size: 20, color: value ? dark : gray }),
      ]
    });
  }

  function bulletItem(text) {
    return new Paragraph({
      spacing: { before: 40, after: 40 },
      indent: { left: 360 },
      children: [
        new TextRun({ text: '▸  ', font: 'DM Sans', size: 20, color: gold }),
        new TextRun({ text: text || '—', font: 'DM Sans', size: 20, color: dark }),
      ]
    });
  }

  function spacer() {
    return new Paragraph({ spacing: { before: 120, after: 0 }, children: [new TextRun('')] });
  }

  function twoColTable(leftLabel, leftVal, rightLabel, rightVal) {
    const cellStyle = (label, val) => new TableCell({
      width: { size: 4500, type: WidthType.DXA },
      margins: { top: 120, bottom: 120, left: 160, right: 160 },
      shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
      borders: {
        top:    { style: BorderStyle.SINGLE, size: 2, color: 'E8E4DF' },
        bottom: { style: BorderStyle.SINGLE, size: 2, color: 'E8E4DF' },
        left:   { style: BorderStyle.SINGLE, size: 2, color: 'E8E4DF' },
        right:  { style: BorderStyle.SINGLE, size: 2, color: 'E8E4DF' },
      },
      children: [
        new Paragraph({ children: [new TextRun({ text: label, font: 'DM Sans', size: 18, bold: true, color: gold })] }),
        new Paragraph({ spacing: { before: 60 }, children: [new TextRun({ text: val || '—', font: 'DM Sans', size: 20, color: dark })] }),
      ]
    });

    return new Table({
      width: { size: 9000, type: WidthType.DXA },
      columnWidths: [4500, 4500],
      rows: [new TableRow({ children: [cellStyle(leftLabel, leftVal), cellStyle(rightLabel, rightVal)] })]
    });
  }

  const arr = (v) => Array.isArray(v) ? v : [];

  const children = [
    // Titelseite
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 600, after: 80 },
      children: [new TextRun({ text: 'Marketing Worksheet', font: 'Georgia', size: 48, bold: true, color: dark })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: 'Kundenavatar & Zielgruppen-Analyse', font: 'DM Sans', size: 24, color: gray })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 600 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: gold } },
      children: [new TextRun({ text: new Date().toLocaleDateString('de-DE', { year: 'numeric', month: 'long', day: 'numeric' }), font: 'DM Sans', size: 20, color: gold })]
    }),
    spacer(),

    // SECTION 1 — Die Basis
    heading('1 · Die Basis'),
    subheading('Werte-Versprechen'),
    body(d.werteversprechen),
    spacer(),
    subheading('USP — Alleinstellungsmerkmal'),
    body(d.usp),
    spacer(),

    // SECTION 2 — Zielgruppe
    heading('2 · Zielgruppe · ' + (d.zielgruppe_name || 'Julia')),
    labelValue('Alter', d.zielgruppe_alter),
    labelValue('Geschlecht', d.zielgruppe_geschlecht),
    labelValue('Wohnort', d.zielgruppe_wohnort),
    labelValue('Sonstiges', d.zielgruppe_sonstiges),
    spacer(),

    // SECTION 3 — Die Regen-Seite
    heading('3 · Die Regen-Seite (weg von)'),

    subheading('Top-Probleme & Herausforderungen'),
    ...arr(d.probleme).map(bulletItem),
    spacer(),

    subheading('Unangenehme Alltagssituationen'),
    ...arr(d.unangenehme_situationen).map(bulletItem),
    spacer(),

    subheading('Was hält sie nachts wach?'),
    ...arr(d.nachts_wach).map(bulletItem),
    spacer(),

    subheading('Konsequenzen ohne Hilfe'),
    ...arr(d.konsequenzen).map(bulletItem),
    spacer(),

    // SECTION 4 — Die Sonnen-Seite
    heading('4 · Die Sonnen-Seite (hin zu)'),

    subheading('Werte der Zielgruppe'),
    ...arr(d.werte_zielgruppe).map(bulletItem),
    spacer(),

    subheading('Gewünschtes Leben'),
    ...arr(d.wunschleben).map(bulletItem),
    spacer(),

    subheading('Insgeheime Wünsche'),
    ...arr(d.insgeheime_wuensche).map(bulletItem),
    spacer(),

    subheading('Kern-Wünsche & Ziele'),
    ...arr(d.kernwuensche).map(bulletItem),
    spacer(),

    // SECTION 5 — Produkt
    heading('5 · Produkt & Dienstleistung'),
    labelValue('Produkt', d.produkt_name),
    labelValue('Preis', d.produkt_preis),
    labelValue('Hauptbenefit', d.hauptbenefit),
    spacer(),
    subheading('Features & Benefits'),
    ...arr(d.features).map(bulletItem),
    spacer(),
    subheading('Typische Einwände'),
    ...arr(d.einwaende).map(bulletItem),
    spacer(),

    // SECTION 6 — Entanglement Map
    heading('6 · Entanglement-Map'),
    spacer(),
    twoColTable(
      'Was schützt die Zielgruppe', d.schutz_mechanismus,
      'Wovon hält es sie ab', d.barriere
    ),
    spacer(),

    // SECTION 7 — Positionierung & Kriterien
    heading('7 · Positionierung & Kampagnen-Kriterien'),
    subheading('Gewählte Positionierung'),
    body(d.positionierung),
    spacer(),
    subheading('Kampagnen-Kriterien'),
    ...arr(d.kampagnen_kriterien).map(bulletItem),
    spacer(),

    // SECTION 8 — Formatierung
    heading('8 · Formatierung & Schreibstil'),
    labelValue('Ansprache', d.sprache_form),
    labelValue('Tonalität', d.tonalitaet),
    spacer(),
    subheading('No-Gos / Abschreckende Worte'),
    ...arr(d.no_gos).map(bulletItem),
  ];

  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: 'DM Sans', size: 20 } }
      }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      children
    }]
  });

  return await Packer.toBuffer(doc);
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server läuft auf Port ${PORT}`));
