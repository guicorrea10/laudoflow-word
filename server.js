require('dotenv').config();

const https_mod = require("https");
const http_mod = require("http");
const express = require("express");
const docx = require("docx");
const sharp = require("sharp");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  PageBreak, ImageRun, Header, Footer, PageNumber,
  BorderStyle, UnderlineType, TabStopType,
  Table, TableRow, TableCell, WidthType, ShadingType,
  LevelFormat, VerticalAlign
} = docx;
const { createClient } = require('@supabase/supabase-js');

const app = express();
const PORT = process.env.PORT || 8080;
app.use(express.json({ limit: '10mb' }));
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

const F        = 'Arial';
const AZUL     = '1F4E79';
const CINZA_LN = 'AAAAAA';
const PT12     = 24;
const PT11     = 22;
const PT10     = 20;
const PT9      = 18;
const PT14     = 28;
const PT18     = 36;
const L15      = 360;
const L1       = 240;
const PAR      = 200;

const MARGEM = { top: 1701, bottom: 1134, left: 1701, right: 1134 };

app.get('/health', (req, res) => res.json({ ok: true, porta: PORT }));

app.post('/gerar-docx', async (req, res) => {
  const { laudo_id } = req.body;
  if (!laudo_id) return res.status(400).json({ erro: 'laudo_id obrigatorio' });
  try {
    const { data: laudo, error: le } = await supabase.from('laudos').select('*').eq('id', laudo_id).single();
    if (le || !laudo) return res.status(404).json({ erro: 'Laudo nao encontrado' });
    const { data: perfil } = await supabase.from('profiles').select('*').eq('id', laudo.user_id).single();
    const { data: fotos } = await supabase.from('fotos').select('*').eq('laudo_id', laudo_id).order('ordem', { ascending: true });
    const doc = await montarDocumento(laudo, perfil, fotos || []);
    const buffer = await Packer.toBuffer(doc);
    const filename = `laudo_${laudo_id.slice(0, 8)}.docx`;
    const storagePath = `docx/${laudo_id}/${filename}`;

    const { error: uploadErr } = await supabase.storage
      .from('fotos')
      .upload(storagePath, buffer, {
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        upsert: true,
      });
    if (uploadErr) throw new Error(`Storage upload: ${uploadErr.message}`);

    const { data: signed, error: signErr } = await supabase.storage
      .from('fotos')
      .createSignedUrl(storagePath, 3600);
    if (signErr || !signed) throw new Error(`Signed URL: ${signErr?.message}`);

    res.json({ url: signed.signedUrl, filename });
  } catch (err) {
    console.error(err);
    res.status(500).json({ erro: err.message });
  }
});

// ─── CORRECAO: extrai o storage path a partir da URL pública ─────────────────
// O campo foto.url pode conter a URL pública completa do Supabase,
// ex: https://xxx.supabase.co/storage/v1/object/public/fotos/laudos/abc/foto.JPG
// O download via SDK precisa apenas do path relativo: laudos/abc/foto.JPG
function extrairStoragePath(urlOuPath) {
  if (!urlOuPath) return null;
  // Se já é um path relativo (não começa com http), retorna direto
  if (!urlOuPath.startsWith('http')) return urlOuPath;
  // Extrai tudo após "/object/public/fotos/" ou "/object/sign/fotos/"
  const match = urlOuPath.match(/\/object\/(?:public|sign)\/fotos\/(.+?)(?:\?|$)/);
  if (match) return match[1];
  // Fallback: extrai após "/fotos/"
  const fallback = urlOuPath.match(/\/fotos\/(.+?)(?:\?|$)/);
  if (fallback) return fallback[1];
  return urlOuPath;
}
// ─────────────────────────────────────────────────────────────────────────────

async function montarDocumento(laudo, perfil, fotos) {
  const nome    = perfil?.nome_completo || 'Engenheiro';
  const crea    = perfil?.crea || '';
  const uf      = perfil?.uf_crea || 'SP';
  const empresa = '';
  const cidade  = perfil?.cidade_atuacao || 'Sao Paulo';
  const rawTipo = laudo.tipo || 'LAUDO DE VISTORIA DE CONSTATACAO';
  const tipo    = rawTipo.split('\n').map(l => l.replace(/^#+\s*/, '').trim()).filter(Boolean)[0]?.toUpperCase() || 'LAUDO DE VISTORIA DE CONSTATACAO';
  const endereco = laudo.endereco || '';
  const cliente  = laudo.cliente || '';

  const filhos = [];

  filhos.push(new Paragraph({
    spacing: { before: 1200, after: 0, line: L1 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: AZUL, space: 1 } },
    children: [new TextRun({ text: '', font: F, size: PT12 })],
  }));

  filhos.push(new Paragraph({
    alignment: 'center',
    spacing: { before: 400, after: 200, line: L1 },
    children: [new TextRun({
      text: 'LAUDO TECNICO',
      font: F, size: PT18, bold: true, color: '000000',
    })],
  }));

  filhos.push(new Paragraph({
    alignment: 'center',
    spacing: { before: 0, after: 400, line: L1 },
    children: [new TextRun({
      text: tipo,
      font: F, size: PT14, color: '444444',
    })],
  }));

  filhos.push(new Paragraph({
    spacing: { before: 0, after: 600, line: L1 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: AZUL, space: 1 } },
    children: [new TextRun({ text: '', font: F, size: PT12 })],
  }));

  if (endereco) {
    filhos.push(new Paragraph({
      alignment: 'center',
      spacing: { before: 0, after: 160, line: L1 },
      children: [new TextRun({ text: endereco, font: F, size: PT14 })],
    }));
  }

  if (cidade && cidade !== 'Sao Paulo') {
    filhos.push(new Paragraph({
      alignment: 'center',
      spacing: { before: 0, after: 600, line: L1 },
      children: [new TextRun({ text: cidade, font: F, size: PT12, color: '444444' })],
    }));
  } else {
    filhos.push(pVazio());
    filhos.push(pVazio());
  }

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
            new Paragraph({ children: [new TextRun({ text: 'Versao', font: F, size: PT10, color: '666666' })] }),
            new Paragraph({ children: [new TextRun({ text: 'final', font: F, size: PT12, italics: true })] }),
          ],
        }),
      ]}),
    ],
  }));

  filhos.push(pVazio());
  filhos.push(pVazio());
  filhos.push(pVazio());

  filhos.push(new Paragraph({
    alignment: 'center',
    spacing: { before: 200, after: 100, line: L1 },
    border: {
      top: { style: BorderStyle.SINGLE, size: 4, color: CINZA_LN, space: 4 },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: CINZA_LN, space: 4 },
    },
    children: [
      new TextRun({ text: empresa ? empresa : '', font: F, size: PT10 }),
    ],
  }));
  filhos.push(new Paragraph({
    alignment: 'center',
    spacing: { before: 0, after: 60, line: L1 },
    children: [new TextRun({ text: `Eng. ${nome} — CREA-${uf} ${crea}`, font: F, size: PT10 })],
  }));

  filhos.push(new Paragraph({ children: [new PageBreak()] }));

  // ─── DOWNLOAD DAS FOTOS ───────────────────────────────────────────────────
  const fotoBuffers = new Map();
  console.log(`[docx] Iniciando download de ${fotos.length} fotos...`);

  for (let idx = 0; idx < fotos.length; idx++) {
    const foto = fotos[idx];
    try {
      const rawUrl = foto.url || '';
      if (!rawUrl) {
        console.warn(`[foto ${idx + 1}] url vazia, pulando`);
        continue;
      }

      // CORRECAO PRINCIPAL: extrai o storage path relativo da URL
      const storagePath = extrairStoragePath(rawUrl);
      if (!storagePath) {
        console.warn(`[foto ${idx + 1}] nao foi possivel extrair storage path de: ${rawUrl}`);
        continue;
      }

      console.log(`[foto ${idx + 1}] storage path: ${storagePath}`);

      const { data: fileData, error: fileErr } = await supabase.storage
        .from('fotos')
        .download(storagePath);

      if (fileErr) {
        console.error(`[foto ${idx + 1}] ERRO download:`, fileErr.message);
        continue;
      }

      const imgBuf = Buffer.from(await fileData.arrayBuffer());
      console.log(`[foto ${idx + 1}] baixada OK — ${imgBuf.length} bytes`);

      let resizedBuf = imgBuf;
      let tipoImg = 'jpeg';

      try {
        resizedBuf = await sharp(imgBuf)
          .resize(1200, 900, { fit: 'inside', withoutEnlargement: true })
          .jpeg({ quality: 85 })
          .toBuffer();
        tipoImg = 'jpeg';
        console.log(`[foto ${idx + 1}] sharp OK — ${resizedBuf.length} bytes`);
      } catch (sharpErr) {
        console.warn(`[foto ${idx + 1}] sharp falhou, usando original:`, sharpErr.message);
        // Detecta tipo pelo path como fallback
        const ext = storagePath.split('.').pop().toLowerCase();
        tipoImg = ext === 'png' ? 'png' : 'jpeg';
      }

      fotoBuffers.set(idx, { buf: resizedBuf, tipoImg });

    } catch (e) {
      console.error(`[foto ${idx + 1}] erro inesperado:`, e.message || e);
    }
  }

  console.log(`[docx] ${fotoBuffers.size} de ${fotos.length} fotos carregadas com sucesso`);
  // ─────────────────────────────────────────────────────────────────────────

  function descricaoFoto(foto) {
    const limpaDesc = s => String(s).replace(/^(?:foto|figura)\s*\d+\s*[-—:]\s*/i, '').trim().slice(0, 200);
    if (foto.titulo) return limpaDesc(foto.titulo);
    if (foto.observacao_engenheiro) return limpaDesc(foto.observacao_engenheiro);
    if (foto.texto_ia) return limpaDesc(foto.texto_ia);
    return '';
  }

  function paragrafosImagem(idx, foto) {
    const entry = fotoBuffers.get(idx);
    if (!entry) {
      console.warn(`[docx] sem buffer para foto ${idx + 1}, marcador omitido`);
      return [];
    }
    const desc = descricaoFoto(foto);
    return [
      new Paragraph({
        alignment: 'center',
        spacing: { before: 400, after: 120, line: L1 },
        children: [new ImageRun({ data: entry.buf, type: entry.tipoImg, transformation: { width: 440, height: 330 } })],
      }),
      new Paragraph({
        alignment: 'center',
        spacing: { before: 60, after: 500, line: L1 },
        children: [
          new TextRun({ text: `Foto ${idx + 1}`, font: F, size: PT10, bold: true }),
          ...(desc ? [new TextRun({ text: ' — ', font: F, size: PT10 }), new TextRun({ text: desc, font: F, size: PT10, italics: true })] : []),
        ],
      }),
    ];
  }

  if (laudo.texto_laudo) {
    const linhas = laudo.texto_laudo.split('\n');
    let primeiraSecaoIdx = -1;
    for (let i = 0; i < linhas.length; i++) {
      const trim = linhas[i].trim();
      if (/^#{1,3}\s+\d+[\.\s]/.test(trim)) {
        primeiraSecaoIdx = i;
        break;
      }
    }

    let linhasProcessar;
    if (primeiraSecaoIdx >= 0) {
      linhasProcessar = linhas.slice(primeiraSecaoIdx);
    } else {
      let skipIdx = 0;
      for (let i = 0; i < linhas.length; i++) {
        const trim = linhas[i].trim();
        if (!trim || trim === '---') { skipIdx = i + 1; continue; }
        if (/^#{1,3}\s+(?!\d)/.test(trim)) { skipIdx = i + 1; continue; }
        if (/^[A-Z\s\-]+$/.test(trim) && trim.length > 3) { skipIdx = i + 1; continue; }
        break;
      }
      linhasProcessar = linhas.slice(skipIdx);
    }

    for (const linha of linhasProcessar) {
      const trim = linha.trim();
      if (!trim || trim === '---') continue;

      // Aceita múltiplos marcadores na mesma linha: [FOTO 1] [FOTO 2] [FOTO 3]
      const todosMarcadores = [...trim.matchAll(/\[(?:FOTO|FIGURA|IMAGEM)[_\s]*(\d+)\]/gi)];
      if (todosMarcadores.length > 0) {
        for (const m of todosMarcadores) {
          const fIdx = parseInt(m[1]) - 1;
          if (fotos[fIdx]) filhos.push(...paragrafosImagem(fIdx, fotos[fIdx]));
        }
        continue;
      }

      if (trim.startsWith('##### ') || trim.startsWith('#### ') || trim.startsWith('### ')) {
        filhos.push(paraH3(trim.replace(/^#+\s*/, '')));
      } else if (trim.startsWith('## ')) {
        filhos.push(paraH2(trim.replace(/^#+\s*/, '')));
      } else if (trim.startsWith('# ')) {
        filhos.push(paraH1(trim.replace(/^#+\s*/, '')));
      } else if (trim.startsWith('- ') || trim.startsWith('• ')) {
        filhos.push(paraBullet(trim.replace(/^[-•]\s*/, '')));
      } else if (/^[a-z]\)\s/.test(trim) || /^\d+\.\s/.test(trim)) {
        filhos.push(paraListaNum(trim));
      } else {
        filhos.push(paraCorpo(trim));
      }
    }
  }

  filhos.push(new Paragraph({ children: [new PageBreak()] }));
  filhos.push(paraH1('ENCERRAMENTO'));

  filhos.push(paraCorpo(
    'O presente laudo foi elaborado em conformidade com a ABNT NBR 13752:2024 e demais normas tecnicas aplicaveis, representando a expressao tecnica das condicoes verificadas na data da vistoria.'
  ));

  filhos.push(pVazio());
  filhos.push(pVazio());

  filhos.push(new Paragraph({
    alignment: 'right',
    spacing: { before: 0, after: PAR, line: L1 },
    children: [new TextRun({ text: `${cidade}, ${formatarData(new Date().toISOString())}`, font: F, size: PT12 })],
  }));

  filhos.push(pVazio());
  filhos.push(pVazio());

  filhos.push(new Paragraph({
    alignment: 'center',
    spacing: { before: 0, after: 80, line: L1 },
    children: [new TextRun({ text: '_'.repeat(52), font: F, size: PT12 })],
  }));
  filhos.push(new Paragraph({
    alignment: 'center',
    spacing: { before: 0, after: 60, line: L1 },
    children: [new TextRun({ text: nome.toUpperCase(), font: F, size: PT12, bold: true })],
  }));
  filhos.push(new Paragraph({
    alignment: 'center',
    spacing: { before: 0, after: 60, line: L1 },
    children: [new TextRun({ text: 'Engenheiro Civil', font: F, size: PT12 })],
  }));
  filhos.push(new Paragraph({
    alignment: 'center',
    spacing: { before: 0, after: 0, line: L1 },
    children: [new TextRun({ text: `CREA-${uf} ${crea}`, font: F, size: PT12 })],
  }));

  filhos.push(new Paragraph({ children: [new PageBreak()] }));
  filhos.push(paraH1('ANEXOS'));

  filhos.push(paraCorpo(
    'Os documentos a seguir relacionados constituem anexos do presente laudo, integrando-o para todos os fins:'
  ));

  filhos.push(pVazio());
  filhos.push(paraCorpo(''));
  filhos.push(paraCorpo(''));
  filhos.push(paraCorpo(''));
  filhos.push(pVazio());

  return new Document({
    styles: {
      default: {
        document: {
          run: { font: F, size: PT12 },
          paragraph: { spacing: { line: L15, before: 0, after: PAR } },
        },
      },
    },
    numbering: {
      config: [{
        reference: 'bullets',
        levels: [{
          level: 0,
          format: 'bullet',
          text: '•',
          alignment: 'left',
          style: {
            paragraph: { indent: { left: 720, hanging: 360 } },
            run: { font: F, size: PT12 },
          },
        }],
      }],
    },
    sections: [{
      properties: {
        page: {
          margin: MARGEM,
          size: { width: 11906, height: 16838 },
        },
      },
      headers: { default: montarCabecalho(tipo, endereco) },
      footers: { default: montarRodape(nome, empresa) },
      children: filhos,
    }],
  });
}

function montarCabecalho(tipo, endereco) {
  return new Header({
    children: [
      new Paragraph({
        spacing: { before: 0, after: 60, line: L1 },
        children: [new TextRun({ text: tipo, font: F, size: PT10, bold: true, color: AZUL })],
      }),
      new Paragraph({
        spacing: { before: 0, after: 0, line: L1 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: CINZA_LN, space: 2 } },
        children: [new TextRun({ text: endereco || '', font: F, size: PT9, color: '888888' })],
      }),
    ],
  });
}

function montarRodape(nome, empresa) {
  return new Footer({
    children: [
      new Paragraph({
        alignment: 'center',
        border: { top: { style: BorderStyle.SINGLE, size: 4, color: CINZA_LN, space: 4 } },
        spacing: { before: 80, after: 0, line: L1 },
        tabStops: [{ type: TabStopType.RIGHT, position: 9071 }],
        children: [
          new TextRun({ text: empresa ? `${empresa}  |  ` : '', font: F, size: PT9, color: '666666' }),
          new TextRun({ text: `Eng. ${nome}`, font: F, size: PT9, color: '666666' }),
          new TextRun({ text: '\t', font: F, size: PT9 }),
          new TextRun({ text: 'Pagina ', font: F, size: PT9, color: '666666' }),
          new TextRun({ children: [PageNumber.CURRENT], font: F, size: PT9, color: '666666' }),
          new TextRun({ text: ' de ', font: F, size: PT9, color: '666666' }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], font: F, size: PT9, color: '666666' }),
        ],
      }),
    ],
  });
}

function paraH1(texto) {
  return new Paragraph({
    spacing: { before: 480, after: 480, line: L1 },
    border: {
      top:    { style: BorderStyle.SINGLE, size: 6, color: AZUL, space: 4 },
      bottom: { style: BorderStyle.SINGLE, size: 6, color: AZUL, space: 4 },
    },
    children: [new TextRun({ text: limpa(texto), font: F, size: PT14, bold: true, color: AZUL })],
  });
}

function paraH2(texto) {
  return new Paragraph({
    spacing: { before: 400, after: 200, line: L15 },
    children: [new TextRun({ text: limpa(texto), font: F, size: PT12, bold: true })],
  });
}

function paraH3(texto) {
  return new Paragraph({
    spacing: { before: 280, after: 160, line: L15 },
    children: [new TextRun({ text: limpa(texto), font: F, size: PT12, bold: true, italics: true })],
  });
}

function paraCorpo(texto) {
  return new Paragraph({
    alignment: 'both',
    spacing: { before: 0, after: PAR, line: L15 },
    children: parseBold(texto),
  });
}

function paraBullet(texto) {
  return new Paragraph({
    alignment: 'both',
    numbering: { reference: 'bullets', level: 0 },
    spacing: { before: 0, after: 100, line: L15 },
    children: parseBold(texto),
  });
}

function paraListaNum(texto) {
  return new Paragraph({
    alignment: 'both',
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
  return String(texto || '').replace(/\*\*/g, '').trim();
}

function parseBold(texto) {
  if (!texto) return [new TextRun({ text: '', font: F, size: PT12 })];
  return String(texto).split(/\*\*(.*?)\*\*/g).map((p, i) =>
    new TextRun({ text: p, font: F, size: PT12, bold: i % 2 === 1 })
  );
}

function formatarData(iso) {
  if (!iso) return '';
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

app.listen(PORT, '0.0.0.0', () => console.log(`LaudoFlow Word Service rodando na porta ${PORT}`));
