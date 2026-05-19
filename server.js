const https_mod = require("https");
const http_mod = require("http");
const express = require("express");
const docx = require("docx");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  PageBreak, ImageRun, Header, Footer, PageNumber,
  BorderStyle, UnderlineType, TabStopType
} = docx;
const { createClient } = require('@supabase/supabase-js');

const app = express();
const PORT = process.env.PORT || 8080;
app.use(express.json({ limit: '10mb' }));
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

// Constantes visuais — padrão Roberto Franco
const F = 'Arial';
const PT12 = 24;   // corpo
const PT14 = 28;   // título capa
const PT10 = 20;   // cabeçalho/rodapé
const PT9  = 18;
const L15  = 360;  // 1,5 entrelinhas
const L1   = 240;  // simples
const PAR  = 200;  // espaço após parágrafo corpo
const MARGEM = { top: 1701, bottom: 1134, left: 1701, right: 1134 };

app.get('/health', (req, res) => res.json({ ok: true, porta: PORT }));

app.post('/gerar-docx', async (req, res) => {
  const { laudo_id } = req.body;
  if (!laudo_id) return res.status(400).json({ erro: 'laudo_id obrigatório' });
  try {
    const { data: laudo, error: le } = await supabase.from('laudos').select('*').eq('id', laudo_id).single();
    if (le || !laudo) return res.status(404).json({ erro: 'Laudo não encontrado' });
    const { data: perfil } = await supabase.from('profiles').select('*').eq('id', laudo.user_id).single();
    const { data: fotos } = await supabase.from('fotos').select('*').eq('laudo_id', laudo_id).order('ordem', { ascending: true });
    const doc = await montarDocumento(laudo, perfil, fotos || []);
    const buffer = await Packer.toBuffer(doc);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="laudo_${laudo_id.slice(0,8)}.docx"`);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: err.message });
  }
});

async function montarDocumento(laudo, perfil, fotos) {
  const nome   = perfil?.nome_completo || 'Engenheiro';
  const crea   = perfil?.crea || '';
  const uf     = perfil?.uf || 'SP';
  const tel    = perfil?.telefone || '';
  const email  = perfil?.email || '';
  const cidade = laudo.cidade || 'São Paulo';
  const titulo = (laudo.titulo || 'LAUDO DE VISTORIA DE CONSTATAÇÃO').toUpperCase();
  const filhos = [];

  // ── CAPA ─────────────────────────────────────────────────────
  // Espaço grande antes do título (como no Roberto — ~1/3 da página)
  for (let i = 0; i < 8; i++) filhos.push(pVazio());

  // Título: negrito, sublinhado, maiúsculo, grande
  filhos.push(new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 1200, line: L1 },
    children: [new TextRun({ text: titulo, font: F, size: PT14, bold: true, underline: { type: UnderlineType.SINGLE } })],
  }));

  // OBJETO DA VISTORIA
  if (laudo.endereco) {
    filhos.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 0, after: 600, line: L15 },
      children: [
        new TextRun({ text: 'OBJETO DA VISTORIA', font: F, size: PT12, bold: true, underline: { type: UnderlineType.SINGLE } }),
        new TextRun({ text: ': ' + laudo.endereco + '.', font: F, size: PT12 }),
      ],
    }));
  }

  // DATA DA VISTORIA
  if (laudo.data_vistoria) {
    filhos.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 0, after: 1200, line: L15 },
      children: [
        new TextRun({ text: 'DATA DA VISTORIA', font: F, size: PT12, bold: true, underline: { type: UnderlineType.SINGLE } }),
        new TextRun({ text: ': ' + formatarData(laudo.data_vistoria) + '.', font: F, size: PT12 }),
      ],
    }));
  }

  // Parágrafo de apresentação — justificado, com nome em caixa alta negrito
  filhos.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: PAR, line: L15 },
    children: [
      new TextRun({ text: nome.toUpperCase(), font: F, size: PT12, bold: true }),
      new TextRun({ text: `, Engenheiro Civil, portador do CREA nº. ${crea}, procedeu à vistoria técnica do imóvel identificado. Apresenta o `, font: F, size: PT12 }),
      new TextRun({ text: titulo, font: F, size: PT12, bold: true }),
      new TextRun({ text: ' em anexo.', font: F, size: PT12 }),
    ],
  }));

  if (laudo.cliente) {
    filhos.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 0, after: PAR, line: L15 },
      children: [
        new TextRun({ text: 'Solicitante: ', font: F, size: PT12, bold: true }),
        new TextRun({ text: laudo.cliente + '.', font: F, size: PT12 }),
      ],
    }));
  }

  // Quebra de página
  filhos.push(new Paragraph({ children: [new PageBreak()] }));

  // ── CORPO ────────────────────────────────────────────────────
  if (laudo.texto_laudo) {
    const linhas = laudo.texto_laudo.split('\n');
    for (const linha of linhas) {
      const trim = linha.trim();
      if (!trim || trim === '---') continue;

      if (trim.startsWith('### ')) {
        filhos.push(paraSubsub(trim.replace(/^#+\s*/, '')));
      } else if (trim.startsWith('## ')) {
        filhos.push(paraSub(trim.replace(/^#+\s*/, '')));
      } else if (trim.startsWith('# ')) {
        filhos.push(paraSecao(trim.replace(/^#+\s*/, '')));
      } else if (trim.startsWith('- ')) {
        filhos.push(paraLista(trim.slice(2), linha));
      } else if (/^[a-z]\)/.test(trim) || /^\d+\./.test(trim)) {
        filhos.push(paraListaNum(trim));
      } else {
        filhos.push(paraCorpo(trim));
      }
    }
  }

  // ── FOTOS ────────────────────────────────────────────────────
  if (fotos.length > 0) {
    filhos.push(new Paragraph({ children: [new PageBreak()] }));
    filhos.push(paraSecao('MEMORIAL FOTOGRÁFICO'));

    for (let idx = 0; idx < fotos.length; idx++) {
      const foto = fotos[idx];
      try {
        const storagePath = foto.url || '';
        const { data: urlData } = supabase.storage.from('fotos').getPublicUrl(storagePath);
        const publicUrl = urlData?.publicUrl;
        if (publicUrl) {
          const imgBuf = await downloadUrl(publicUrl);
          const ext = storagePath.split('.').pop()?.toLowerCase();
          const tipo = ext === 'png' ? 'png' : 'jpeg';

          filhos.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 400, after: 120, line: L1 },
            children: [new ImageRun({ data: imgBuf, type: tipo, transformation: { width: 400, height: 300 } })],
          }));

          const leg = foto.texto_ia
            ? String(foto.texto_ia).slice(0, 100)
            : (foto.observacao_engenheiro ? String(foto.observacao_engenheiro).slice(0, 80) : `Figura ${idx + 1}`);

          filhos.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 400, line: L1 },
            children: [new TextRun({ text: `Figura ${idx + 1} – ${leg}`, font: F, size: PT10, italics: true })],
          }));
        }
      } catch (e) { console.warn(`Foto ${idx + 1}:`, e.message); }
    }
  }

  // ── ASSINATURA ───────────────────────────────────────────────
  for (let i = 0; i < 3; i++) filhos.push(pVazio());
  filhos.push(new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before: 0, after: PAR, line: L1 },
    children: [new TextRun({ text: `${cidade}, ${formatarData(new Date().toISOString())}`, font: F, size: PT12 })],
  }));
  for (let i = 0; i < 3; i++) filhos.push(pVazio());
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80, line: L1 },
    children: [new TextRun({ text: '_'.repeat(50), font: F, size: PT12 })],
  }));
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60, line: L1 },
    children: [new TextRun({ text: nome.toUpperCase(), font: F, size: PT12, bold: true })],
  }));
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60, line: L1 },
    children: [new TextRun({ text: 'Engenheiro Civil', font: F, size: PT12 })],
  }));
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: L1 },
    children: [new TextRun({ text: `CREA-${uf} nº. ${crea}`, font: F, size: PT12 })],
  }));

  return new Document({
    styles: {
      default: {
        document: {
          run: { font: F, size: PT12 },
          paragraph: { spacing: { line: L15, before: 0, after: PAR } },
        },
      },
    },
    sections: [{
      properties: { page: { margin: MARGEM, size: { width: 11906, height: 16838 } } },
      headers: { default: cabecalho(nome, crea) },
      footers: { default: rodape(tel, email) },
      children: filhos,
    }],
  });
}

// ── Cabeçalho padrão Roberto ──────────────────────────────────
function cabecalho(nome, crea) {
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '1E3A5F', space: 4 } },
        spacing: { before: 0, after: 60, line: L1 },
        children: [new TextRun({ text: nome, font: F, size: PT10, bold: true })],
      }),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 0, after: 0, line: L1 },
        children: [new TextRun({ text: `Engenheiro Civil  CREA ${crea}`, font: F, size: PT10 })],
      }),
    ],
  });
}

// ── Rodapé padrão Roberto ─────────────────────────────────────
function rodape(tel, email) {
  return new Footer({
    children: [
      new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } },
        spacing: { before: 80, after: 0, line: L1 },
        tabStops: [{ type: TabStopType.RIGHT, position: 9026 }],
        children: [
          ...(tel   ? [new TextRun({ text: `✆ ${tel}`, font: F, size: PT9 })]       : []),
          ...(email ? [new TextRun({ text: `   ✉ ${email}`, font: F, size: PT9 })]  : []),
          new TextRun({ text: '\t', font: F, size: PT9 }),
          new TextRun({ text: 'Página ', font: F, size: PT9 }),
          new TextRun({ children: [PageNumber.CURRENT], font: F, size: PT9 }),
        ],
      }),
    ],
  });
}

// ── Helpers de parágrafo ──────────────────────────────────────

// Seção: "1 – TÍTULO" negrito, espaço antes
function paraSecao(texto) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 480, after: 240, line: L15 },
    children: [new TextRun({ text: limpa(texto), font: F, size: PT12, bold: true })],
  });
}

// Subseção: "1.1 – Título" negrito
function paraSub(texto) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 320, after: 160, line: L15 },
    children: [new TextRun({ text: limpa(texto), font: F, size: PT12, bold: true })],
  });
}

// Subsubseção itálico negrito
function paraSubsub(texto) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 200, after: 120, line: L15 },
    children: [new TextRun({ text: limpa(texto), font: F, size: PT12, bold: true, italics: true })],
  });
}

// Corpo: justificado, recuo primeira linha (padrão Roberto)
function paraCorpo(texto) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: PAR, line: L15 },
    indent: { firstLine: 720 },
    children: parseBold(texto),
  });
}

// Lista com traço
function paraLista(texto, linhaOriginal) {
  const nivel = linhaOriginal.match(/^(\s+)/)?.[1]?.length || 0;
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 100, line: L15 },
    indent: { left: nivel >= 4 ? 1440 : 720, hanging: 360 },
    children: [new TextRun({ text: '– ', font: F, size: PT12 }), ...parseBold(texto)],
  });
}

// Lista numerada
function paraListaNum(texto) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 100, line: L15 },
    indent: { left: 720, hanging: 360 },
    children: parseBold(texto),
  });
}

function pVazio() {
  return new Paragraph({
    spacing: { before: 0, after: 0, line: L1 },
    children: [new TextRun({ text: '', font: F, size: PT12 })],
  });
}

function limpa(texto) {
  return String(texto || '').replace(/\*\*/g, '');
}

function parseBold(texto) {
  if (!texto) return [new TextRun({ text: '', font: F, size: PT12 })];
  return String(texto).split(/\*\*(.*?)\*\*/g).map((p, i) =>
    new TextRun({ text: p, font: F, size: PT12, bold: i % 2 === 1 })
  );
}

function formatarData(iso) {
  if (!iso) return '';
  try { return new Date(iso).toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' }); }
  catch { return iso; }
}

function downloadUrl(url) {
  return new Promise((resolve, reject) => {
    const lib = url.startsWith('https') ? https_mod : http_mod;
    lib.get(url, (res) => {
      if (res.statusCode !== 200) { reject(new Error('HTTP ' + res.statusCode)); return; }
      const chunks = [];
      res.on('data', chunk => chunks.push(chunk));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

app.listen(PORT, '0.0.0.0', () => console.log(`LaudoFlow Word Service rodando na porta ${PORT}`));
