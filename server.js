const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  PageBreak, ImageRun, Header, Footer, PageNumber,
  BorderStyle, UnderlineType, TabStopType
  BorderStyle, UnderlineType, TabStopType,
  Table, TableRow, TableCell, WidthType, ShadingType,
  LevelFormat, VerticalAlign
} = docx;
const { createClient } = require('@supabase/supabase-js');

@@ -14,16 +16,23 @@ const PORT = process.env.PORT || 8080;
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
// ── Constantes visuais — padrão Roberto Franco Godoi / NASF ──
const F        = 'Arial';
const AZUL     = '1F4E79';   // azul escuro do Roberto
const CINZA_LN = 'AAAAAA';  // linha cinza
const PT12     = 24;
const PT11     = 22;
const PT10     = 20;
const PT9      = 18;
const PT14     = 28;
const PT18     = 36;
const L15      = 360;  // 1,5 entrelinhas
const L1       = 240;  // simples
const PAR      = 200;  // espaço após parágrafo

// Margens A4 ABNT: esq 3cm, dir 2cm, sup 3cm, inf 2cm (DXA)
const MARGEM = { top: 1701, bottom: 1134, left: 1701, right: 1134 };
// Largura útil A4 com essas margens: 11906 - 1701 - 1134 = 9071 DXA

app.get('/health', (req, res) => res.json({ ok: true, porta: PORT }));

@@ -46,104 +55,194 @@ app.post('/gerar-docx', async (req, res) => {
  }
});

// ════════════════════════════════════════════════════════════════
async function montarDocumento(laudo, perfil, fotos) {
  const nome   = perfil?.nome_completo || 'Engenheiro';
  const crea   = perfil?.crea || '';
  const uf     = perfil?.uf || 'SP';
  const tel    = perfil?.telefone || '';
  const email  = perfil?.email || '';
  const cidade = laudo.cidade || 'São Paulo';
  const titulo = (laudo.titulo || 'LAUDO DE VISTORIA DE CONSTATAÇÃO').toUpperCase();
  const nome    = perfil?.nome_completo || 'Engenheiro';
  const crea    = perfil?.crea || '';
  const uf      = perfil?.uf || 'SP';
  const tel     = perfil?.telefone || '';
  const email   = perfil?.email || '';
  const empresa = perfil?.empresa || '';
  const cidade  = laudo.cidade || 'São Paulo';
  const tipo    = (laudo.tipo_laudo || 'LAUDO DE VISTORIA DE CONSTATAÇÃO').toUpperCase();
  const endereco = laudo.endereco || '';
  const cliente  = laudo.cliente || '';

  const filhos = [];

  // ── CAPA ─────────────────────────────────────────────────────
  // Espaço grande antes do título (como no Roberto — ~1/3 da página)
  for (let i = 0; i < 8; i++) filhos.push(pVazio());
  // ── CAPA ────────────────────────────────────────────────────
  // Linha azul acima do título (como no Roberto)
  filhos.push(new Paragraph({
    spacing: { before: 1200, after: 0, line: L1 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: AZUL, space: 1 } },
    children: [new TextRun({ text: '', font: F, size: PT12 })],
  }));

  // Título centralizado, grande, negrito, azul
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 400, line: L1 },
    children: [new TextRun({
      text: tipo,
      font: F, size: PT18, bold: true, color: '000000',
    })],
  }));

  // Título: negrito, sublinhado, maiúsculo, grande
  // Linha azul abaixo do título
  filhos.push(new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 1200, line: L1 },
    children: [new TextRun({ text: titulo, font: F, size: PT14, bold: true, underline: { type: UnderlineType.SINGLE } })],
    spacing: { before: 0, after: 600, line: L1 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: AZUL, space: 1 } },
    children: [new TextRun({ text: '', font: F, size: PT12 })],
  }));

  // OBJETO DA VISTORIA
  if (laudo.endereco) {
  // Endereço centralizado
  if (endereco) {
    filhos.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 0, after: 600, line: L15 },
      children: [
        new TextRun({ text: 'OBJETO DA VISTORIA', font: F, size: PT12, bold: true, underline: { type: UnderlineType.SINGLE } }),
        new TextRun({ text: ': ' + laudo.endereco + '.', font: F, size: PT12 }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 160, line: L1 },
      children: [new TextRun({ text: endereco, font: F, size: PT14 })],
    }));
  }

  // DATA DA VISTORIA
  if (laudo.data_vistoria) {
  // Bairro/cidade se houver
  if (cidade && cidade !== 'São Paulo') {
    filhos.push(new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { before: 0, after: 1200, line: L15 },
      children: [
        new TextRun({ text: 'DATA DA VISTORIA', font: F, size: PT12, bold: true, underline: { type: UnderlineType.SINGLE } }),
        new TextRun({ text: ': ' + formatarData(laudo.data_vistoria) + '.', font: F, size: PT12 }),
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 600, line: L1 },
      children: [new TextRun({ text: cidade, font: F, size: PT12, color: '444444' })],
    }));
  } else {
    filhos.push(pVazio());
    filhos.push(pVazio());
  }

  // Parágrafo de apresentação — justificado, com nome em caixa alta negrito
  // Tabela de metadados (Contratante / Data / Documento / Versão)
  const dataTxt = laudo.data_vistoria ? formatarData(laudo.data_vistoria) : '—';
  const bordCell = {
    top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
  };

  filhos.push(new Table({
    width: { size: 9071, type: WidthType.DXA },
    columnWidths: [4535, 4536],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: 'DDDDDD' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: 'DDDDDD' },
      left: { style: BorderStyle.NONE, size: 0 },
      right: { style: BorderStyle.NONE, size: 0 },
      insideH: { style: BorderStyle.SINGLE, size: 4, color: 'DDDDDD' },
      insideV: { style: BorderStyle.NONE, size: 0 },
    },
    rows: [
      new TableRow({ children: [
        new TableCell({
          borders: bordCell,
          width: { size: 4535, type: WidthType.DXA },
          margins: { top: 120, bottom: 60, left: 160, right: 160 },
          children: [
            new Paragraph({ children: [new TextRun({ text: 'Contratante', font: F, size: PT10, color: '666666' })] }),
            new Paragraph({ children: [new TextRun({ text: cliente || '—', font: F, size: PT12, bold: true })] }),
          ],
        }),
        new TableCell({
          borders: bordCell,
          width: { size: 4536, type: WidthType.DXA },
          margins: { top: 120, bottom: 60, left: 160, right: 160 },
          children: [
            new Paragraph({ children: [new TextRun({ text: 'Data da Vistoria', font: F, size: PT10, color: '666666' })] }),
            new Paragraph({ children: [new TextRun({ text: dataTxt, font: F, size: PT12, bold: true })] }),
          ],
        }),
      ]}),
      new TableRow({ children: [
        new TableCell({
          borders: bordCell,
          width: { size: 4535, type: WidthType.DXA },
          margins: { top: 60, bottom: 120, left: 160, right: 160 },
          children: [
            new Paragraph({ children: [new TextRun({ text: 'Documento', font: F, size: PT10, color: '666666' })] }),
            new Paragraph({ children: [new TextRun({ text: laudo_id_curto(laudo.id), font: F, size: PT12 })] }),
          ],
        }),
        new TableCell({
          borders: bordCell,
          width: { size: 4536, type: WidthType.DXA },
          margins: { top: 60, bottom: 120, left: 160, right: 160 },
          children: [
            new Paragraph({ children: [new TextRun({ text: 'Versão', font: F, size: PT10, color: '666666' })] }),
            new Paragraph({ children: [new TextRun({ text: 'final', font: F, size: PT12, italics: true })] }),
          ],
        }),
      ]}),
    ],
  }));

  filhos.push(pVazio());
  filhos.push(pVazio());
  filhos.push(pVazio());

  // Linha cinza abaixo, rodapé da capa com engenheiro
  filhos.push(new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { before: 0, after: PAR, line: L15 },
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100, line: L1 },
    border: {
      top: { style: BorderStyle.SINGLE, size: 4, color: CINZA_LN, space: 4 },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: CINZA_LN, space: 4 },
    },
    children: [
      new TextRun({ text: nome.toUpperCase(), font: F, size: PT12, bold: true }),
      new TextRun({ text: `, Engenheiro Civil, portador do CREA nº. ${crea}, procedeu à vistoria técnica do imóvel identificado. Apresenta o `, font: F, size: PT12 }),
      new TextRun({ text: titulo, font: F, size: PT12, bold: true }),
      new TextRun({ text: ' em anexo.', font: F, size: PT12 }),
      new TextRun({ text: empresa ? empresa : '', font: F, size: PT10 }),
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
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60, line: L1 },
    children: [new TextRun({ text: `Eng. ${nome} — CREA-${uf} ${crea}`, font: F, size: PT10 })],
  }));

  // Quebra de página
  filhos.push(new Paragraph({ children: [new PageBreak()] }));

  // ── CORPO ────────────────────────────────────────────────────
  // ── CORPO ───────────────────────────────────────────────────
  if (laudo.texto_laudo) {
    const linhas = laudo.texto_laudo.split('\n');
    // Filtrar bloco de cabeçalho interno gerado pela IA
    // (tudo antes da primeira seção numerada real: "# 1." ou "## 1.")
    let iniciou = false;
    for (const linha of linhas) {
      const trim = linha.trim();
      if (!trim || trim === '---') continue;

      // Detectar início do conteúdo real (seção numerada)
      if (!iniciou) {
        const ehSecaoReal = /^#{1,3}\s*\d+[\.\s]/.test(trim) ||
                            /^#{1,3}\s+[A-ZÁÉÍÓÚÃÕÂÊÔÇÜ]/.test(trim);
        if (!ehSecaoReal) continue;
        iniciou = true;
      }

      if (trim.startsWith('### ')) {
        filhos.push(paraSubsub(trim.replace(/^#+\s*/, '')));
        filhos.push(paraH3(trim.replace(/^#+\s*/, '')));
      } else if (trim.startsWith('## ')) {
        filhos.push(paraSub(trim.replace(/^#+\s*/, '')));
        filhos.push(paraH2(trim.replace(/^#+\s*/, '')));
      } else if (trim.startsWith('# ')) {
        filhos.push(paraSecao(trim.replace(/^#+\s*/, '')));
      } else if (trim.startsWith('- ')) {
        filhos.push(paraLista(trim.slice(2), linha));
      } else if (/^[a-z]\)/.test(trim) || /^\d+\./.test(trim)) {
        filhos.push(paraH1(trim.replace(/^#+\s*/, '')));
      } else if (trim.startsWith('- ') || trim.startsWith('• ')) {
        filhos.push(paraBullet(trim.replace(/^[-•]\s*/, ''), linha));
      } else if (/^[a-z]\)\s/.test(trim) || /^\d+\.\s/.test(trim)) {
        filhos.push(paraListaNum(trim));
      } else {
        filhos.push(paraCorpo(trim));
      }
    }
  }

  // ── FOTOS ────────────────────────────────────────────────────
  // ── FOTOS ───────────────────────────────────────────────────
  if (fotos.length > 0) {
    filhos.push(new Paragraph({ children: [new PageBreak()] }));
    filhos.push(paraSecao('MEMORIAL FOTOGRÁFICO'));
    filhos.push(paraH1('MEMORIAL FOTOGRÁFICO'));

    for (let idx = 0; idx < fotos.length; idx++) {
      const foto = fotos[idx];
@@ -156,38 +255,61 @@ async function montarDocumento(laudo, perfil, fotos) {
          const ext = storagePath.split('.').pop()?.toLowerCase();
          const tipo = ext === 'png' ? 'png' : 'jpeg';

          // Foto centralizada
          filhos.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 400, after: 120, line: L1 },
            children: [new ImageRun({ data: imgBuf, type: tipo, transformation: { width: 400, height: 300 } })],
            children: [new ImageRun({
              data: imgBuf,
              type: tipo,
              transformation: { width: 440, height: 330 },
            })],
          }));

          const leg = foto.texto_ia
            ? String(foto.texto_ia).slice(0, 100)
            : (foto.observacao_engenheiro ? String(foto.observacao_engenheiro).slice(0, 80) : `Figura ${idx + 1}`);
          // Legenda padrão Roberto: "Foto N — descrição" em itálico
          const descFoto = foto.observacao_engenheiro
            ? String(foto.observacao_engenheiro).slice(0, 100)
            : (foto.texto_ia ? String(foto.texto_ia).slice(0, 100) : '');

          filhos.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 60, after: 400, line: L1 },
            children: [new TextRun({ text: `Figura ${idx + 1} – ${leg}`, font: F, size: PT10, italics: true })],
            spacing: { before: 60, after: 500, line: L1 },
            children: [
              new TextRun({ text: `Foto ${idx + 1}`, font: F, size: PT10, bold: true }),
              ...(descFoto ? [new TextRun({ text: ` — `, font: F, size: PT10 }),
                              new TextRun({ text: descFoto, font: F, size: PT10, italics: true })] : []),
            ],
          }));
        }
      } catch (e) { console.warn(`Foto ${idx + 1}:`, e.message); }
    }
  }

  // ── ASSINATURA ───────────────────────────────────────────────
  for (let i = 0; i < 3; i++) filhos.push(pVazio());
  // ── ENCERRAMENTO / ASSINATURA ────────────────────────────────
  filhos.push(new Paragraph({ children: [new PageBreak()] }));
  filhos.push(paraH1('ENCERRAMENTO'));

  filhos.push(paraCorpo(
    `O presente laudo foi elaborado em conformidade com a ABNT NBR 13752:2024 e demais normas técnicas aplicáveis, representando a expressão técnica das condições verificadas na data da vistoria.`
  ));

  filhos.push(pVazio());
  filhos.push(pVazio());

  filhos.push(new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before: 0, after: PAR, line: L1 },
    children: [new TextRun({ text: `${cidade}, ${formatarData(new Date().toISOString())}`, font: F, size: PT12 })],
  }));
  for (let i = 0; i < 3; i++) filhos.push(pVazio());

  filhos.push(pVazio());
  filhos.push(pVazio());

  // Linha de assinatura
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80, line: L1 },
    children: [new TextRun({ text: '_'.repeat(50), font: F, size: PT12 })],
    children: [new TextRun({ text: '_'.repeat(52), font: F, size: PT12 })],
  }));
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
@@ -202,9 +324,10 @@ async function montarDocumento(laudo, perfil, fotos) {
  filhos.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: L1 },
    children: [new TextRun({ text: `CREA-${uf} nº. ${crea}`, font: F, size: PT12 })],
    children: [new TextRun({ text: `CREA-${uf} ${crea}`, font: F, size: PT12 })],
  }));

  // ── DOCUMENTO ───────────────────────────────────────────────
  return new Document({
    styles: {
      default: {
@@ -214,48 +337,74 @@ async function montarDocumento(laudo, perfil, fotos) {
        },
      },
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [{
            level: 0,
            format: LevelFormat.BULLET,
            text: '•',
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: { indent: { left: 720, hanging: 360 } },
              run: { font: F, size: PT12 },
            },
          }],
        },
      ],
    },
    sections: [{
      properties: { page: { margin: MARGEM, size: { width: 11906, height: 16838 } } },
      headers: { default: cabecalho(nome, crea) },
      footers: { default: rodape(tel, email) },
      properties: {
        page: {
          margin: MARGEM,
          size: { width: 11906, height: 16838 }, // A4
        },
      },
      headers: { default: montarCabecalho(tipo, endereco) },
      footers: { default: montarRodape(nome, empresa) },
      children: filhos,
    }],
  });
}

// ── Cabeçalho padrão Roberto ──────────────────────────────────
function cabecalho(nome, crea) {
// Linha 1: tipo do laudo em azul bold + linha azul abaixo
// Linha 2: endereço em cinza
function montarCabecalho(tipo, endereco) {
  return new Header({
    children: [
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '1E3A5F', space: 4 } },
        spacing: { before: 0, after: 60, line: L1 },
        children: [new TextRun({ text: nome, font: F, size: PT10, bold: true })],
        children: [new TextRun({ text: tipo, font: F, size: PT10, bold: true, color: AZUL })],
      }),
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        spacing: { before: 0, after: 0, line: L1 },
        children: [new TextRun({ text: `Engenheiro Civil  CREA ${crea}`, font: F, size: PT10 })],
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: CINZA_LN, space: 2 } },
        children: [new TextRun({ text: endereco || '', font: F, size: PT9, color: '888888' })],
      }),
    ],
  });
}

// ── Rodapé padrão Roberto ─────────────────────────────────────
function rodape(tel, email) {
// "Empresa | Engenheiro — Página X de Y"
function montarRodape(nome, empresa) {
  return new Footer({
    children: [
      new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC', space: 4 } },
        alignment: AlignmentType.CENTER,
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: CINZA_LN, space: 4 } },
        spacing: { before: 80, after: 0, line: L1 },
        tabStops: [{ type: TabStopType.RIGHT, position: 9026 }],
        tabStops: [{ type: TabStopType.RIGHT, position: 9071 }],
        children: [
          ...(tel   ? [new TextRun({ text: `✆ ${tel}`, font: F, size: PT9 })]       : []),
          ...(email ? [new TextRun({ text: `   ✉ ${email}`, font: F, size: PT9 })]  : []),
          new TextRun({ text: empresa ? `${empresa}  |  ` : '', font: F, size: PT9, color: '666666' }),
          new TextRun({ text: `Eng. ${nome}`, font: F, size: PT9, color: '666666' }),
          new TextRun({ text: '\t', font: F, size: PT9 }),
          new TextRun({ text: 'Página ', font: F, size: PT9 }),
          new TextRun({ children: [PageNumber.CURRENT], font: F, size: PT9 }),
          new TextRun({ text: 'Página ', font: F, size: PT9, color: '666666' }),
          new TextRun({ children: [PageNumber.CURRENT], font: F, size: PT9, color: '666666' }),
          new TextRun({ text: ' de ', font: F, size: PT9, color: '666666' }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: F, size: PT9, color: '666666' }),
        ],
      }),
    ],
@@ -264,55 +413,54 @@ function rodape(tel, email) {

// ── Helpers de parágrafo ──────────────────────────────────────

// Seção: "1 – TÍTULO" negrito, espaço antes
function paraSecao(texto) {
// H1: linha azul acima e abaixo, texto azul bold grande
function paraH1(texto) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 480, after: 240, line: L15 },
    children: [new TextRun({ text: limpa(texto), font: F, size: PT12, bold: true })],
    spacing: { before: 480, after: 480, line: L1 },
    border: {
      top:    { style: BorderStyle.SINGLE, size: 6, color: AZUL, space: 4 },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: AZUL, space: 4 },
    },
    children: [new TextRun({ text: limpa(texto), font: F, size: PT14, bold: true, color: AZUL })],
  });
}

// Subseção: "1.1 – Título" negrito
function paraSub(texto) {
// H2: bold, sem linha, espaço antes
function paraH2(texto) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 320, after: 160, line: L15 },
    spacing: { before: 400, after: 200, line: L15 },
    children: [new TextRun({ text: limpa(texto), font: F, size: PT12, bold: true })],
  });
}

// Subsubseção itálico negrito
function paraSubsub(texto) {
// H3: bold itálico
function paraH3(texto) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 200, after: 120, line: L15 },
    spacing: { before: 280, after: 160, line: L15 },
    children: [new TextRun({ text: limpa(texto), font: F, size: PT12, bold: true, italics: true })],
  });
}

// Corpo: justificado, recuo primeira linha (padrão Roberto)
// Corpo: justificado, sem recuo (padrão Roberto)
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
// Bullet com numbering config
function paraBullet(texto, linhaOriginal) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    numbering: { reference: 'bullets', level: 0 },
    spacing: { before: 0, after: 100, line: L15 },
    indent: { left: nivel >= 4 ? 1440 : 720, hanging: 360 },
    children: [new TextRun({ text: '– ', font: F, size: PT12 }), ...parseBold(texto)],
    children: parseBold(texto),
  });
}

// Lista numerada
// Lista numerada/letrada
function paraListaNum(texto) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
@@ -330,7 +478,7 @@ function pVazio() {
}

function limpa(texto) {
  return String(texto || '').replace(/\*\*/g, '');
  return String(texto || '').replace(/\*\*/g, '').trim();
}

function parseBold(texto) {
@@ -342,8 +490,16 @@ function parseBold(texto) {

function formatarData(iso) {
  if (!iso) return '';
  try { return new Date(iso).toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' }); }
  catch { return iso; }
  try {
    return new Date(iso).toLocaleDateString('pt-BR', {
      day: '2-digit', month: 'long', year: 'numeric',
    });
  } catch { return iso; }
}

function laudo_id_curto(id) {
  if (!id) return '—';
  return id.slice(0, 8).toUpperCase();
}

function downloadUrl(url) {
