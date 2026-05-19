const express = require('express');
const { 
  Document, Packer, Paragraph, TextRun, ImageRun, Header, Footer,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, PageNumber,
  PageBreak, LevelFormat
} = require('docx');

const app = express();
app.use(express.json());

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_KEY;

async function supabaseGet(path) {
  const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    headers: {
      'apikey': SUPABASE_KEY,
      'authorization': `Bearer ${SUPABASE_KEY}`,
      'content-type': 'application/json'
    }
  });
  return res.json();
}

async function downloadImage(url) {
  try {
    const res = await fetch(url);
    if (!res.ok) return null;
    const arrayBuffer = await res.arrayBuffer();
    return Buffer.from(arrayBuffer);
  } catch {
    return null;
  }
}

function getImageType(url) {
  const lower = url.toLowerCase();
  if (lower.includes('.png')) return 'png';
  if (lower.includes('.jpg') || lower.includes('.jpeg')) return 'jpg';
  return 'jpg';
}

function parseBoldText(text) {
  const runs = [];
  const parts = text.split(/\*\*(.*?)\*\*/g);
  for (let i = 0; i < parts.length; i++) {
    if (!parts[i]) continue;
    runs.push(new TextRun({
      text: parts[i],
      bold: i % 2 === 1,
      font: 'Arial',
      size: 24
    }));
  }
  return runs.length > 0 ? runs : [new TextRun({ text, font: 'Arial', size: 24 })];
}

function parseMarkdownToParagraphs(text) {
  if (!text) return [];
  const paragraphs = [];
  const lines = text.split('\n');

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      paragraphs.push(new Paragraph({ children: [new TextRun('')] }));
      continue;
    }
    if (trimmed.startsWith('## ') || trimmed.startsWith('# ')) {
      paragraphs.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: trimmed.replace(/^#+\s+/, ''), bold: true, font: 'Arial', size: 24 })],
        pageBreakBefore: true,
        spacing: { before: 240, after: 120 }
      }));
      continue;
    }
    if (trimmed.startsWith('### ')) {
      paragraphs.push(new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: trimmed.replace(/^###\s+/, ''), bold: true, font: 'Arial', size: 22 })],
        spacing: { before: 180, after: 90 }
      }));
      continue;
    }
    if (trimmed === '---') {
      paragraphs.push(new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'AAAAAA', space: 1 } },
        children: [new TextRun('')]
      }));
      continue;
    }
    const runs = parseBoldText(trimmed);
    paragraphs.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 0, after: 160, line: 360 },
      children: runs
    }));
  }
  return paragraphs;
}

function makeHeader(perfil) {
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '2E75B6', space: 4 } },
        children: [new TextRun({ text: perfil.nome_completo || 'Engenheiro Responsável', bold: true, font: 'Arial', size: 20, color: '2E75B6' })]
      }),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { after: 80 },
        children: [new TextRun({ text: `${perfil.especialidade || 'Engenheiro Civil'} CREA-${perfil.uf_crea || ''} ${perfil.crea || ''}`, font: 'Arial', size: 18, color: '555555' })]
      })
    ]
  });
}

function makeFooter() {
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } },
        children: [
          new TextRun({ text: 'Página ', font: 'Arial', size: 18, color: '555555' }),
          new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 18, color: '555555' })
        ]
      })
    ]
  });
}

app.post('/gerar-docx', async (req, res) => {
  const { laudo_id } = req.body;
  if (!laudo_id) return res.status(400).json({ error: 'laudo_id obrigatório' });

  try {
    const laudos = await supabaseGet(`laudos?id=eq.${laudo_id}&select=*`);
    const laudo = laudos[0];
    if (!laudo) return res.status(404).json({ error: 'Laudo não encontrado' });

    const perfis = await supabaseGet(`profiles?id=eq.${laudo.user_id}&select=*`);
    const perfil = perfis[0] || {};

    const fotos = await supabaseGet(`fotos?laudo_id=eq.${laudo_id}&aprovado=eq.true&order=ordem.asc&select=*`);

    const children = [];

    children.push(
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1440, after: 480 }, children: [new TextRun({ text: laudo.titulo || 'LAUDO TÉCNICO', bold: true, font: 'Arial', size: 36 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: `Tipo: ${laudo.tipo || 'Vistoria de Constatação'}`, font: 'Arial', size: 24 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: `Nº ${laudo.id}`, font: 'Arial', size: 22, color: '555555' })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: `Cliente: ${laudo.cliente || ''}`, font: 'Arial', size: 24 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: `Endereço: ${laudo.endereco || ''}`, font: 'Arial', size: 24 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: `Data da vistoria: ${laudo.data_vistoria || ''}`, font: 'Arial', size: 24 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 1440 }, children: [new TextRun({ text: `Responsável: ${perfil.nome_completo || ''} — CREA-${perfil.uf_crea || ''} ${perfil.crea || ''}`, font: 'Arial', size: 24, bold: true })] }),
      new Paragraph({ children: [new PageBreak()] })
    );

    if (laudo.texto_laudo) {
      children.push(...parseMarkdownToParagraphs(laudo.texto_laudo));
    }

    if (fotos && fotos.length > 0) {
      children.push(
        new Paragraph({ children: [new PageBreak()] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 240, after: 240 }, children: [new TextRun({ text: 'MEMORIAL FOTOGRÁFICO', bold: true, font: 'Arial', size: 24 })] })
      );

      for (let i = 0; i < fotos.length; i++) {
        const foto = fotos[i];
        let imgBuffer = null;

        if (foto.url) {
          const storageUrl = `${SUPABASE_URL}/storage/v1/object/public/fotos/${foto.url}`;
          imgBuffer = await downloadImage(storageUrl);
          if (!imgBuffer && foto.url.startsWith('http')) {
            imgBuffer = await downloadImage(foto.url);
          }
        }

        if (imgBuffer) {
          children.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 240, after: 120 },
            children: [new ImageRun({
              type: getImageType(foto.url || 'jpg'),
              data: imgBuffer,
              transformation: { width: 480, height: 360 },
              altText: { title: `Figura ${i + 1}`, description: foto.observacao_engenheiro || '', name: `figura${i + 1}` }
            })]
          }));
        }

        children.push(new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 60, after: 60 },
          children: [new TextRun({ text: `Figura ${i + 1} — ${foto.observacao_engenheiro || ''}`, font: 'Arial', size: 20, italics: true })]
        }));

        if (foto.patamar_prioridade) {
          children.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 240 },
            children: [new TextRun({ text: `Classificação: ${foto.patamar_prioridade}`, font: 'Arial', size: 20, bold: true, color: foto.patamar_prioridade.includes('1') ? 'C0392B' : foto.patamar_prioridade.includes('2') ? 'E67E22' : '27AE60' })]
          }));
        }

        if (i < fotos.length - 1) {
          children.push(new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'EEEEEE', space: 4 } }, spacing: { after: 240 }, children: [new TextRun('')] }));
        }
      }
    }

    const doc = new Document({
      styles: {
        default: { document: { run: { font: 'Arial', size: 24 } } },
        paragraphStyles: [
          { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 24, bold: true, font: 'Arial' }, paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
          { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 22, bold: true, font: 'Arial' }, paragraph: { spacing: { before: 180, after: 90 }, outlineLevel: 1 } }
        ]
      },
      numbering: { config: [{ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
      sections: [{
        properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1701, right: 1134, bottom: 1134, left: 1701 } } },
        headers: { default: makeHeader(perfil) },
        footers: { default: makeFooter() },
        children
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    const filename = `laudo_${(laudo.titulo || laudo_id).replace(/[^a-z0-9_.\-]/gi, '_')}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.send(buffer);

  } catch (err) {
    console.error('Erro ao gerar DOCX:', err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/health', (req, res) => res.json({ ok: true }));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`LaudoFlow Word Service rodando na porta ${PORT}`));
