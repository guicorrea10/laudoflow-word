const express = require('express');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  PageBreak, ImageRun, Header, Footer, PageNumber,
  BorderStyle, TabStopType, UnderlineType, LevelFormat
} = require('docx');
const { createClient } = require('@supabase/supabase-js');

const app = express();
const PORT = process.env.PORT || 8080;
app.use(express.json({ limit: '10mb' }));

const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

const F = 'Arial';
const PT14 = 28; const PT12 = 24; const PT10 = 20; const PT9 = 18;
const LINHA_SIMPLES = 240; const LINHA_15 = 360;
const ESP_PAR = 160; const ESP_SECAO = 320;
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
  const titulo = (laudo.titulo || 'LAUDO TÉCNICO DE VISTORIA DE CONSTATAÇÃO').toUpperCase();
  const filhos = [];

  // CAPA
  filhos.push(pVazio());
  filhos.push(pVazio());

  // Título sublinhado e negrito
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 600, line: LINHA_SIMPLES },
    children: [new TextRun({ text: titulo, font: F, size: PT14, bold: true, underline: { type: UnderlineType.SINGLE } })],
  }));

  filhos.push(pVazio());

  if (laudo.endereco) filhos.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 320, line: LINHA_15 },
    children: [
      new TextRun({ text: 'OBJETO DA VISTORIA', font: F, size: PT12, bold: true, underline: { type: UnderlineType.SINGLE } }),
      new TextRun({ text: ': ' + laudo.endereco + '.', font: F, size: PT12 }),
    ],
  }));

  if (laudo.data_vistoria) filhos.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 600, line: LINHA_15 },
    children: [
      new TextRun({ text: 'DATA DA VISTORIA', font: F, size: PT12, bold: true, underline: { type: UnderlineType.SINGLE } }),
      new TextRun({ text: ': ' + formatarData(laudo.data_vistoria) + '.', font: F, size: PT12 }),
    ],
  }));

  filhos.push(pVazio());

  filhos.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: ESP_PAR, line: LINHA_15 },
    indent: { firstLine: 720 },
    children: [
      new TextRun({ text: nome.toUpperCase(), font: F, size: PT12, bold: true }),
      new TextRun({ text: `, Engenheiro Civil, portador do CREA nº. ${crea}, procedeu à vistoria técnica do imóvel. Apresenta o `, font: F, size: PT12 }),
      new TextRun({ text: titulo, font: F, size: PT12, bold: true }),
      new TextRun({ text: ' em anexo.', font: F, size: PT12 }),
    ],
  }));

  if (laudo.cliente) filhos.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: ESP_PAR, line: LINHA_15 },
    indent: { firstLine: 720 },
    children: [
      new TextRun({ text: 'Solicitante: ', font: F, size: PT12, bold: true }),
      new TextRun({ text: laudo.cliente + '.', font: F, size: PT12 }),
    ],
  }));

  // Quebra de página
  filhos.push(new Paragraph({ children: [new PageBreak()] }));

  // CORPO
  if (laudo.texto_laudo) {
    const linhas = laudo.texto_laudo.split('\n');
    for (let i = 0; i < linhas.length; i++) {
      const linha = linhas[i];
      const trim = linha.trim();
      if (!trim || trim === '---') continue;
      if (trim.startsWith('### ')) { filhos.push(paraSubsubsecao(trim.replace(/^#+\s*/, ''))); continue; }
      if (trim.startsWith('## '))  { filhos.push(paraSubsecao(trim.replace(/^#+\s*/, '')));    continue; }
      if (trim.startsWith('# '))   { filhos.push(paraSecao(trim.replace(/^#+\s*/, '')));       continue; }
      if (trim.startsWith('- '))   { filhos.push(paraLista(trim.slice(2), linha));             continue; }
      if (/^[a-z]\)/.test(trim) || /^\d+\./.test(trim)) { filhos.push(paraListaNum(trim));    continue; }
      filhos.push(paraCorpo(trim));
    }
  }

  // FOTOS
  if (fotos.length > 0) {
    filhos.push(new Paragraph({ children: [new PageBreak()] }));
    filhos.push(paraSecao('MEMORIAL FOTOGRÁFICO'));
    for (let idx = 0; idx < fotos.length; idx++) {
      const foto = fotos[idx];
      try {
        // url salva como caminho relativo no bucket 'laudos'
        // ex: "laudos/cc56bc0c-.../foto.png"
        const storagePath = foto.url || '';
        const bucketPath = storagePath.startsWith('laudos/')
          ? storagePath.slice('laudos/'.length)  // remove prefixo do bucket
          : storagePath;
        const { data: imgData, error: imgErr } = await supabase.storage.from('laudos').download(bucketPath);
        if (!imgErr && imgData) {
          const ab = await imgData.arrayBuffer();
          const uint8 = new Uint8Array(ab);
          const ext = (storagePath || '').split('.').pop()?.toLowerCase();
          const tipo = ext === 'png' ? 'png' : 'jpeg';
          filhos.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 240, after: 80, line: LINHA_SIMPLES },
            children: [new ImageRun({ data: uint8, type: tipo, transformation: { width: 420, height: 315 } })],
          }));
          const leg = foto.texto_laudo ? foto.texto_laudo.slice(0, 120) : `Figura ${idx + 1}`;
          filhos.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 280, line: LINHA_SIMPLES },
            children: [new TextRun({ text: `Figura ${idx + 1} – ${leg}`, font: F, size: PT10, italics: true })],
          }));
        }
      } catch (e) { console.warn(`Foto ${idx + 1}:`, e.message); }
    }
  }

  // ASSINATURA
  filhos.push(pVazio());
  filhos.push(new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before: 0, after: ESP_PAR, line: LINHA_SIMPLES },
    children: [new TextRun({ text: `${cidade}, ${formatarData(new Date().toISOString())}`, font: F, size: PT12 })],
  }));
  filhos.push(pVazio());
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60, line: LINHA_SIMPLES },
    children: [new TextRun({ text: '_'.repeat(48), font: F, size: PT12 })],
  }));
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60, line: LINHA_SIMPLES },
    children: [new TextRun({ text: nome.toUpperCase(), font: F, size: PT12, bold: true })],
  }));
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60, line: LINHA_SIMPLES },
    children: [new TextRun({ text: 'Engenheiro Civil', font: F, size: PT12 })],
  }));
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: LINHA_SIMPLES },
    children: [new TextRun({ text: `CREA-${uf} nº. ${crea}`, font: F, size: PT12 })],
  }));

  return new Document({
    styles: {
      default: {
        document: {
          run: { font: F, size: PT12 },
          paragraph: { spacing: { line: LINHA_15, before: 0, after: ESP_PAR } },
        },
      },
    },
    sections: [{
      properties: { page: { margin: MARGEM, size: { width: 11906, height: 16838 } } },
      headers: { default: montarCabecalho(nome, crea) },
      footers: { default: montarRodape(tel, email) },
      children: filhos,
    }],
  });
}

function montarCabecalho(nome, crea) {
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '1E3A5F', space: 4 } },
        spacing: { before: 0, after: 60, line: LINHA_SIMPLES },
        children: [new TextRun({ text: nome, font: F, size: PT10, bold: true })],
      }),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 0, after: 0, line: LINHA_SIMPLES },
        children: [new TextRun({ text: `Engenheiro Civil  CREA ${crea}`, font: F, size: PT10 })],
      }),
    ],
  });
}

function montarRodape(tel, email) {
  return new Footer({
    children: [
      new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } },
        spacing: { before: 80, after: 0, line: LINHA_SIMPLES },
        tabStops: [{ type: TabStopType.RIGHT, position: 9026 }],
        children: [
          ...(tel   ? [new TextRun({ text: `✆ ${tel}`, font: F, size: PT9 })]         : []),
          ...(email ? [new TextRun({ text: `   ✉ ${email}`, font: F, size: PT9 })]    : []),
          new TextRun({ text: '\t', font: F, size: PT9 }),
          new TextRun({ text: 'Página ', font: F, size: PT9 }),
          new TextRun({ children: [PageNumber.CURRENT], font: F, size: PT9 }),
        ],
      }),
    ],
  });
}

function paraSecao(texto) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: ESP_SECAO, after: ESP_PAR, line: LINHA_15 },
    children: [new TextRun({ text: texto.replace(/\*\*/g,''), font: F, size: PT12, bold: true })],
  });
}
function paraSubsecao(texto) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: ESP_PAR + 80, after: ESP_PAR, line: LINHA_15 },
    children: [new TextRun({ text: texto.replace(/\*\*/g,''), font: F, size: PT12, bold: true })],
  });
}
function paraSubsubsecao(texto) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: ESP_PAR, after: ESP_PAR, line: LINHA_15 },
    children: [new TextRun({ text: texto.replace(/\*\*/g,''), font: F, size: PT12, bold: true, italics: true })],
  });
}
function paraCorpo(texto) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: ESP_PAR, line: LINHA_15 },
    indent: { firstLine: 720 },
    children: parseBold(texto),
  });
}
function paraLista(texto, linhaOriginal) {
  const nivel = linhaOriginal.match(/^(\s+)/)?.[1]?.length || 0;
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 80, line: LINHA_15 },
    indent: { left: nivel >= 4 ? 1440 : 720, hanging: 360 },
    children: [new TextRun({ text: '– ', font: F, size: PT12 }), ...parseBold(texto)],
  });
}
function paraListaNum(texto) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: 80, line: LINHA_15 },
    indent: { left: 720, hanging: 360 },
    children: parseBold(texto),
  });
}
function pVazio() {
  return new Paragraph({
    spacing: { before: 0, after: 0, line: LINHA_SIMPLES },
    children: [new TextRun({ text: '', font: F, size: PT12 })],
  });
}
function parseBold(texto) {
  return texto.split(/\*\*(.*?)\*\*/g).map((p, i) =>
    new TextRun({ text: p, font: F, size: PT12, bold: i % 2 === 1 })
  );
}
function formatarData(iso) {
  if (!iso) return '';
  try { return new Date(iso).toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' }); }
  catch { return iso; }
}

app.listen(PORT, '0.0.0.0', () => console.log(`LaudoFlow Word Service rodando na porta ${PORT}`));
