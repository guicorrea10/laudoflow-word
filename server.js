const express = require('express');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak, ImageRun, Header, Footer, PageNumber, NumberFormat, Table, TableRow, TableCell, WidthType, BorderStyle } = require('docx');
const { createClient } = require('@supabase/supabase-js');

const app = express();

// ✅ CRÍTICO: Railway injeta PORT dinamicamente — NUNCA hardcode 8080
const PORT = process.env.PORT || 8080;

app.use(express.json({ limit: '10mb' }));

// Supabase
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_KEY
);

// ─── Health check ───────────────────────────────────────────
app.get('/health', (req, res) => {
  res.json({ ok: true, porta: PORT, ts: new Date().toISOString() });
});

// ─── Gerar DOCX ─────────────────────────────────────────────
app.post('/gerar-docx', async (req, res) => {
  const { laudo_id } = req.body;

  if (!laudo_id) {
    return res.status(400).json({ erro: 'laudo_id obrigatório', body_recebido: req.body });
  }

  try {
    // 1. Buscar dados do laudo
    const { data: laudo, error: laudoErr } = await supabase
      .from('laudos')
      .select('*')
      .eq('id', laudo_id)
      .single();

    if (laudoErr || !laudo) {
      return res.status(404).json({ erro: 'Laudo não encontrado', detalhe: laudoErr?.message });
    }

    // 2. Buscar perfil do engenheiro
    const { data: perfil } = await supabase
      .from('profiles')
      .select('*')
      .eq('id', laudo.user_id)
      .single();

    // 3. Buscar fotos
    const { data: fotos } = await supabase
      .from('fotos')
      .select('*')
      .eq('laudo_id', laudo_id)
      .order('ordem', { ascending: true });

    // 4. Montar documento Word
    const doc = await montarDocumento(laudo, perfil, fotos || []);

    // 5. Gerar buffer e retornar
    const buffer = await Packer.toBuffer(doc);

    const nomeArquivo = `laudo_${laudo_id.slice(0, 8)}.docx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${nomeArquivo}"`);
    res.setHeader('Content-Length', buffer.length);
    res.send(buffer);

  } catch (err) {
    console.error('Erro ao gerar DOCX:', err);
    res.status(500).json({ erro: 'Erro interno ao gerar documento', detalhe: err.message });
  }
});

// ─── Montagem do documento ───────────────────────────────────
async function montarDocumento(laudo, perfil, fotos) {
  const engNome = perfil?.nome_completo || 'Engenheiro';
  const engCrea = perfil?.crea || '';
  const engUf = perfil?.uf || 'SP';

  const secoes = [];

  // CAPA
  secoes.push(
    new Paragraph({
      children: [new TextRun({ text: '', break: 1 })],
    }),
    new Paragraph({
      text: laudo.titulo || 'LAUDO TÉCNICO DE VISTORIA',
      heading: HeadingLevel.TITLE,
      alignment: AlignmentType.CENTER,
      spacing: { before: 2400, after: 400 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: laudo.tipo_laudo || '', size: 24, bold: false }),
      ],
      spacing: { after: 800 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: laudo.endereco || '', size: 24 }),
      ],
      spacing: { after: 400 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: `Cliente: ${laudo.cliente || ''}`, size: 24 }),
      ],
      spacing: { after: 400 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: `Data: ${formatarData(laudo.data_vistoria)}`, size: 24 }),
      ],
      spacing: { after: 800 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: `${engNome}`, size: 24, bold: true }),
      ],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: `CREA-${engUf} ${engCrea}`, size: 24 }),
      ],
      spacing: { after: 400 },
    }),
    // Quebra de página após capa
    new Paragraph({ children: [new PageBreak()] }),
  );

  // TEXTO DO LAUDO (gerado pela IA)
  if (laudo.texto_laudo) {
    const linhas = laudo.texto_laudo.split('\n');
    for (const linha of linhas) {
      const trimmed = linha.trim();
      if (!trimmed) {
        secoes.push(new Paragraph({ text: '' }));
        continue;
      }

      // Detectar headings markdown
      if (trimmed.startsWith('### ')) {
        secoes.push(new Paragraph({
          text: trimmed.replace('### ', ''),
          heading: HeadingLevel.HEADING_3,
          spacing: { before: 400, after: 200 },
        }));
      } else if (trimmed.startsWith('## ')) {
        secoes.push(new Paragraph({
          text: trimmed.replace('## ', ''),
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 600, after: 200 },
        }));
      } else if (trimmed.startsWith('# ')) {
        secoes.push(new Paragraph({
          text: trimmed.replace('# ', ''),
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 800, after: 400 },
          pageBreakBefore: true,
        }));
      } else if (trimmed.startsWith('**') && trimmed.endsWith('**')) {
        secoes.push(new Paragraph({
          children: [new TextRun({ text: trimmed.replace(/\*\*/g, ''), bold: true, size: 24 })],
          spacing: { before: 200, after: 200 },
        }));
      } else {
        secoes.push(new Paragraph({
          children: [new TextRun({ text: trimmed, size: 24 })],
          spacing: { after: 200 },
          indent: { left: 0 },
        }));
      }
    }
  }

  // FOTOS (se houver)
  if (fotos.length > 0) {
    secoes.push(
      new Paragraph({ children: [new PageBreak()] }),
      new Paragraph({
        text: 'REGISTROS FOTOGRÁFICOS',
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 400, after: 400 },
      }),
    );

    for (let i = 0; i < fotos.length; i++) {
      const foto = fotos[i];
      try {
        // Baixar imagem do Supabase Storage
        const { data: imgData, error: imgErr } = await supabase.storage
          .from('fotos')
          .download(foto.storage_path || foto.url?.split('/fotos/')[1]);

        if (!imgErr && imgData) {
          const arrayBuffer = await imgData.arrayBuffer();
          const uint8 = new Uint8Array(arrayBuffer);

          secoes.push(
            new Paragraph({
              children: [
                new ImageRun({
                  data: uint8,
                  transformation: { width: 400, height: 300 },
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { before: 400, after: 200 },
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: `Foto ${i + 1}${foto.texto_laudo ? ' — ' + foto.texto_laudo.slice(0, 80) : ''}`,
                  size: 20,
                  italics: true,
                }),
              ],
              spacing: { after: 400 },
            }),
          );
        }
      } catch (e) {
        console.warn(`Não foi possível incluir foto ${i + 1}:`, e.message);
      }
    }
  }

  // ASSINATURA FINAL
  secoes.push(
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({
      text: `${laudo.cidade || 'São Paulo'}, ${formatarData(new Date().toISOString())}`,
      alignment: AlignmentType.RIGHT,
      spacing: { before: 800, after: 800 },
      children: [new TextRun({ text: `${laudo.cidade || 'São Paulo'}, ${formatarData(new Date().toISOString())}`, size: 24 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: '_'.repeat(50), size: 24 })],
      spacing: { before: 1200, after: 200 },
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: engNome, bold: true, size: 24 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: `Engenheiro Civil — CREA-${engUf} ${engCrea}`, size: 24 })],
    }),
  );

  // Montar doc com formatação ABNT
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: 'Arial', size: 24 }, // 12pt = 24 half-points
          paragraph: { spacing: { line: 360 } }, // 1,5 entrelinhas
        },
      },
    },
    sections: [{
      properties: {
        page: {
          margin: {
            top: 1701,    // 3cm
            bottom: 1134, // 2cm
            left: 1701,   // 3cm
            right: 1134,  // 2cm
          },
        },
      },
      headers: {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: `${engNome} — CREA-${engUf} ${engCrea}`, size: 20, color: '2563EB' }),
              ],
              border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '2563EB' } },
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [
                new TextRun({ children: [PageNumber.CURRENT], size: 20 }),
                new TextRun({ text: ' / ', size: 20 }),
                new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 20 }),
              ],
            }),
          ],
        }),
      },
      children: secoes,
    }],
  });

  return doc;
}

// ─── Helpers ─────────────────────────────────────────────────
function formatarData(iso) {
  if (!iso) return '';
  try {
    return new Date(iso).toLocaleDateString('pt-BR', {
      day: '2-digit', month: 'long', year: 'numeric'
    });
  } catch {
    return iso;
  }
}

// ─── Start ───────────────────────────────────────────────────
app.listen(PORT, '0.0.0.0', () => {
  console.log(`LaudoFlow Word Service rodando na porta ${PORT}`);
});
