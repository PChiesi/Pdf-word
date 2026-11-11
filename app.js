/* Conversor PDF ⇄ Word — tudo roda no navegador; nenhum arquivo é enviado. */
(function () {
  'use strict';

  pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';

  const $ = (id) => document.getElementById(id);
  const el = {
    abas: document.querySelectorAll('.abas button'),
    soltar: $('soltar'), arquivo: $('arquivo'), soltarT: $('soltar-t'),
    opPdf: $('op-pdf2word'), opWord: $('op-word2pdf'),
    progresso: $('progresso'), barra: $('barra'), status: $('status'),
    pronto: $('pronto'), prontoT: $('pronto-t'), prontoS: $('pronto-s'),
    baixar: $('baixar'), outro: $('outro'), erro: $('erro'),
    previa: $('previa'), previaC: $('previa-conteudo'),
  };

  let modo = 'pdf2word';
  let resultado = null; // { acao: () => void, rotulo }
  let ocupado = false;

  // ---------- Interface ----------
  function trocarModo(novo) {
    if (ocupado) return;
    modo = novo;
    el.abas.forEach((b) => b.setAttribute('aria-selected', String(b.dataset.modo === novo)));
    el.opPdf.hidden = novo !== 'pdf2word';
    el.opWord.hidden = novo !== 'word2pdf';
    el.soltarT.textContent = novo === 'pdf2word' ? 'Arraste o PDF aqui' : 'Arraste o arquivo Word (.docx) aqui';
    el.arquivo.accept = novo === 'pdf2word' ? '.pdf,application/pdf' : '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    reiniciar();
  }
  el.abas.forEach((b) => b.addEventListener('click', () => trocarModo(b.dataset.modo)));

  function reiniciar() {
    resultado = null;
    el.progresso.hidden = true; el.pronto.hidden = true; el.erro.hidden = true;
    el.previa.hidden = true; el.previaC.innerHTML = '';
    el.soltar.hidden = false; el.arquivo.value = '';
  }
  el.outro.addEventListener('click', reiniciar);
  el.baixar.addEventListener('click', () => resultado && resultado.acao());

  function progresso(frac, texto) {
    el.barra.style.width = Math.round(Math.max(0, Math.min(1, frac)) * 100) + '%';
    if (texto) el.status.textContent = texto;
  }
  function mostrarErro(msg) {
    el.progresso.hidden = true;
    el.soltar.hidden = false;
    el.erro.hidden = false;
    el.erro.textContent = msg;
  }

  ['dragenter', 'dragover'].forEach((ev) => el.soltar.addEventListener(ev, (e) => { e.preventDefault(); el.soltar.classList.add('sobre'); }));
  ['dragleave', 'drop'].forEach((ev) => el.soltar.addEventListener(ev, (e) => { e.preventDefault(); el.soltar.classList.remove('sobre'); }));
  el.soltar.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) iniciar(f); });
  el.arquivo.addEventListener('change', () => { const f = el.arquivo.files[0]; if (f) iniciar(f); });
  // Permite soltar em qualquer lugar da página
  document.addEventListener('dragover', (e) => e.preventDefault());
  document.addEventListener('drop', (e) => { e.preventDefault(); const f = e.dataTransfer && e.dataTransfer.files[0]; if (f && !ocupado) iniciar(f); });

  async function iniciar(arquivo) {
    if (ocupado) return;
    const nome = arquivo.name.toLowerCase();
    // Se a pessoa soltou o tipo "errado", troca de aba sozinho.
    if (nome.endsWith('.pdf') && modo !== 'pdf2word') trocarModo('pdf2word');
    if (nome.endsWith('.docx') && modo !== 'word2pdf') trocarModo('word2pdf');
    reiniciar();
    if (nome.endsWith('.doc')) return mostrarErro('Arquivos .doc (formato antigo) não são suportados. Abra no Word, use "Salvar como" → "Documento do Word (.docx)" e tente de novo.');
    if (modo === 'pdf2word' && !nome.endsWith('.pdf')) return mostrarErro('Escolha um arquivo .pdf.');
    if (modo === 'word2pdf' && !nome.endsWith('.docx')) return mostrarErro('Escolha um arquivo .docx.');

    ocupado = true;
    el.soltar.hidden = true;
    el.progresso.hidden = false;
    progresso(0, 'Lendo ' + arquivo.name + '…');
    try {
      if (modo === 'pdf2word') await pdfParaWord(arquivo);
      else await wordParaPdf(arquivo);
    } catch (err) {
      console.error(err);
      let msg = 'Não foi possível converter este arquivo. ' + (err && err.message ? '(' + err.message + ')' : '');
      if (err && err.name === 'PasswordException') msg = 'Este PDF é protegido por senha. Remova a senha e tente de novo.';
      mostrarErro(msg);
    } finally {
      ocupado = false;
    }
  }

  function concluir(titulo, sub, rotulo, acao) {
    el.progresso.hidden = true;
    el.pronto.hidden = false;
    el.prontoT.textContent = titulo;
    el.prontoS.textContent = sub;
    el.baixar.textContent = rotulo;
    resultado = { acao };
  }

  function baixarBlob(blob, nome) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = nome;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
  const semExt = (n) => n.replace(/\.[^.]+$/, '');
  const pausa = () => new Promise((r) => setTimeout(r, 0));

  // =====================================================================
  // PDF → Word
  // =====================================================================
  const FONTES = [
    [/^arial.*narrow/i, 'Arial Narrow'], [/^arial/i, 'Arial'], [/^helvetica/i, 'Arial'],
    [/^times|timesnewroman/i, 'Times New Roman'], [/^courier/i, 'Courier New'],
    [/^calibri/i, 'Calibri'], [/^cambria/i, 'Cambria'], [/^candara/i, 'Candara'],
    [/^consolas/i, 'Consolas'], [/^constantia/i, 'Constantia'], [/^corbel/i, 'Corbel'],
    [/^georgia/i, 'Georgia'], [/^verdana/i, 'Verdana'], [/^tahoma/i, 'Tahoma'],
    [/^trebuchet/i, 'Trebuchet MS'], [/^segoeui|^segoe/i, 'Segoe UI'], [/^garamond/i, 'Garamond'],
    [/^bookantiqua/i, 'Book Antiqua'], [/^centurygothic/i, 'Century Gothic'], [/^century/i, 'Century'],
    [/^aptos/i, 'Aptos'], [/^symbol/i, 'Symbol'], [/^wingdings/i, 'Wingdings'],
    [/^roboto/i, 'Roboto'], [/^opensans/i, 'Open Sans'], [/^lato/i, 'Lato'], [/^montserrat/i, 'Montserrat'],
  ];

  function infoFonte(page, fontName, estilo) {
    let obj = null;
    try { obj = page.commonObjs.get(fontName); } catch (e) { /* ainda não carregada */ }
    const bruto = ((obj && obj.name) || '').replace(/^[A-Z]{6}\+/, '');
    const [fam, sufixo = ''] = bruto.split(/[-,]/);
    const familiaLimpa = (fam || '').replace(/(PS)?MT$/, '').replace(/PS$/, '');
    let familia = null;
    for (const [re, nome] of FONTES) if (re.test(familiaLimpa)) { familia = nome; break; }
    if (!familia && familiaLimpa) familia = familiaLimpa.replace(/([a-z])([A-Z])/g, '$1 $2');
    const generica = (estilo && estilo.fontFamily) || 'sans-serif';
    if (!familia) familia = generica === 'serif' ? 'Times New Roman' : generica === 'monospace' ? 'Courier New' : 'Arial';
    const s = (sufixo + ' ' + bruto).toLowerCase();
    const negrito = !!(obj && (obj.bold || obj.black)) || /bold|black|heavy|semibold|demi/.test(s);
    const italico = !!(obj && obj.italic) || /italic|oblique/.test(s);
    return { familia, generica, negrito, italico };
  }

  const medidor = document.createElement('canvas').getContext('2d');
  function medir(texto, f, tamanho) {
    medidor.font = `${f.italico ? 'italic ' : ''}${f.negrito ? 'bold ' : ''}${tamanho}px "${f.familia}", ${f.generica}`;
    return medidor.measureText(texto).width;
  }

  // Cor de fundo (mediana da borda) e cor do texto (pixels mais distantes do fundo).
  function analisarCaixa(dados, W, H, x0, y0, x1, y1) {
    x0 = Math.max(0, Math.floor(x0)); y0 = Math.max(0, Math.floor(y0));
    x1 = Math.min(W - 1, Math.ceil(x1)); y1 = Math.min(H - 1, Math.ceil(y1));
    if (x1 <= x0 || y1 <= y0) return null;
    const r = [], g = [], b = [];
    const amostra = (x, y) => { const i = (y * W + x) * 4; r.push(dados[i]); g.push(dados[i + 1]); b.push(dados[i + 2]); };
    const passoX = Math.max(1, Math.floor((x1 - x0) / 60));
    for (let x = x0; x <= x1; x += passoX) { amostra(x, y0); amostra(x, y1); }
    for (let y = y0; y <= y1; y++) { amostra(x0, y); amostra(x1, y); }
    const med = (v) => v.sort((p, q) => p - q)[v.length >> 1];
    const fundo = [med(r), med(g), med(b)];
    let maxD = 0;
    const passo = Math.max(1, Math.floor(Math.sqrt(((x1 - x0) * (y1 - y0)) / 4000)));
    for (let y = y0; y <= y1; y += passo) for (let x = x0; x <= x1; x += passo) {
      const i = (y * W + x) * 4;
      const d = Math.abs(dados[i] - fundo[0]) + Math.abs(dados[i + 1] - fundo[1]) + Math.abs(dados[i + 2] - fundo[2]);
      if (d > maxD) maxD = d;
    }
    // Cor mais frequente entre os pixels de "tinta" (ignora bordas suavizadas)
    const bins = new Map();
    for (let y = y0; y <= y1; y += passo) for (let x = x0; x <= x1; x += passo) {
      const i = (y * W + x) * 4;
      const d = Math.abs(dados[i] - fundo[0]) + Math.abs(dados[i + 1] - fundo[1]) + Math.abs(dados[i + 2] - fundo[2]);
      if (d < maxD * 0.6) continue;
      const k = (dados[i] >> 4) << 8 | (dados[i + 1] >> 4) << 4 | (dados[i + 2] >> 4);
      const v = bins.get(k) || [0, 0, 0, 0];
      v[0] += dados[i]; v[1] += dados[i + 1]; v[2] += dados[i + 2]; v[3]++;
      bins.set(k, v);
    }
    let melhor = null;
    for (const v of bins.values()) if (!melhor || v[3] > melhor[3]) melhor = v;
    const cor = melhor && maxD > 30 ? [melhor[0] / melhor[3], melhor[1] / melhor[3], melhor[2] / melhor[3]] : [0, 0, 0];
    return { fundo, cor };
  }
  const hex = (c) => c.map((v) => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase();

  async function pdfParaWord(arquivo) {
    const D = window.DocxLib;
    const editavel = document.querySelector('input[name="m1"]:checked').value === 'editavel';
    const dadosArq = new Uint8Array(await arquivo.arrayBuffer());
    const pdf = await pdfjsLib.getDocument({
      data: dadosArq, cMapUrl: 'vendor/cmaps/', cMapPacked: true, standardFontDataUrl: 'vendor/standard_fonts/',
    }).promise;

    const secoes = [];
    let semTexto = 0;
    for (let p = 1; p <= pdf.numPages; p++) {
      progresso((p - 1) / pdf.numPages, `Convertendo página ${p} de ${pdf.numPages}…`);
      await pausa();
      const page = await pdf.getPage(p);
      const vp = page.getViewport({ scale: 1 }); // unidades em pontos (1/72")
      const Wpt = vp.width, Hpt = vp.height;
      // ~200 dpi, limitado para não estourar a memória
      let S = 200 / 72;
      const maxPx = 5000;
      S = Math.min(S, maxPx / Wpt, maxPx / Hpt);
      const canvas = document.createElement('canvas');
      canvas.width = Math.round(Wpt * S); canvas.height = Math.round(Hpt * S);
      const ctx = canvas.getContext('2d', { willReadFrequently: true });
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport: page.getViewport({ scale: S }) }).promise;

      const filhos = [];
      if (editavel) {
        const tc = await page.getTextContent();
        const itens = [];
        for (const it of tc.items) {
          if (!it.str || !it.str.trim()) continue;
          const t = pdfjsLib.Util.transform(vp.transform, it.transform);
          const tam = Math.hypot(t[2], t[3]);
          // Só texto horizontal vira texto editável; o resto fica na imagem de fundo.
          if (!(t[0] > 0 && Math.abs(t[1]) < 0.01 * tam && Math.abs(t[2]) < 0.01 * tam) || tam < 1) continue;
          const est = tc.styles[it.fontName] || {};
          const asc = est.ascent && est.ascent > 0.3 ? est.ascent : 0.8;
          const desc = est.descent && est.descent < 0 ? -est.descent : 0.2;
          const largura = it.width > 0 ? it.width : tam * it.str.length * 0.5;
          itens.push({
            str: it.str, x: t[4], base: t[5], tam, largura,
            topo: t[5] - tam * asc, fundo: t[5] + tam * desc,
            fonte: infoFonte(page, it.fontName, est),
          });
        }
        if (!itens.length) semTexto++;

        // Cores e "apagar" o texto da imagem de fundo
        const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const pad = 1.5;
        for (const it of itens) {
          const a = analisarCaixa(img.data, canvas.width, canvas.height,
            it.x * S, it.topo * S - pad, (it.x + it.largura) * S, it.fundo * S + pad);
          it.cor = a ? hex(a.cor) : '000000';
          it.fundoCor = a ? a.fundo : [255, 255, 255];
        }
        for (const it of itens) {
          ctx.fillStyle = `rgb(${it.fundoCor.join(',')})`;
          ctx.fillRect(it.x * S - pad, it.topo * S - pad, it.largura * S + 2 * pad, (it.fundo - it.topo) * S + 2 * pad);
        }

        // Agrupa em linhas
        itens.sort((a, b) => a.base - b.base || a.x - b.x);
        const linhas = [];
        for (const it of itens) {
          const l = linhas.find((L) => Math.abs(L.base - it.base) < 0.3 * Math.min(L.tam, it.tam)
            && it.x >= L.fim - 0.5 * it.tam && it.x - L.fim < 2.5 * Math.max(L.tam, it.tam));
          if (l) { l.itens.push(it); l.fim = Math.max(l.fim, it.x + it.largura); l.tam = Math.max(l.tam, it.tam); l.topo = Math.min(l.topo, it.topo); }
          else linhas.push({ base: it.base, fim: it.x + it.largura, tam: it.tam, topo: it.topo, itens: [it] });
        }

        for (const L of linhas) {
          L.itens.sort((a, b) => a.x - b.x);
          const runs = [];
          let cursor = L.itens[0].x;
          for (const it of L.itens) {
            const meio = Math.round(it.tam * 2) / 2; // Word aceita meio ponto
            const base = { font: it.fonte.familia, size: Math.max(2, Math.round(meio * 2)), bold: it.fonte.negrito, italics: it.fonte.italico, color: it.cor };
            const gap = it.x - cursor;
            if (gap > 0.05 * it.tam && runs.length) {
              const esp = medir(' ', it.fonte, meio);
              runs.push(new D.TextRun({ ...base, text: ' ', characterSpacing: Math.round((gap - esp) * 20) }));
            }
            // Ajusta o espaçamento para a linha ocupar exatamente a mesma largura do PDF.
            // Texto justificado: a sobra vai para os espaços entre palavras.
            const natural = medir(it.str, it.fonte, meio);
            const sobra = it.largura - natural;
            const nEsp = (it.str.match(/ /g) || []).length;
            if (sobra > 0.02 * it.tam * it.str.length && nEsp > 0 && it.str.trim().length > 0) {
              const extra = Math.min(meio * 40, (sobra / nEsp) * 20);
              for (const parte of it.str.split(/( +)/)) {
                if (!parte) continue;
                runs.push(new D.TextRun({ ...base, text: parte, characterSpacing: parte[0] === ' ' ? Math.round(extra) : 0 }));
              }
            } else {
              let cs = (sobra / it.str.length) * 20; // twips por caractere
              cs = Math.max(-meio * 6, Math.min(meio * 10, cs));
              runs.push(new D.TextRun({ ...base, text: it.str, characterSpacing: Math.round(cs) }));
            }
            cursor = it.x + it.largura;
          }
          const alturaLinha = L.tam * 1.2;
          filhos.push(new D.Paragraph({
            children: runs,
            spacing: { before: 0, after: 0, line: Math.round(alturaLinha * 20), lineRule: D.LineRuleType.EXACT },
            frame: {
              type: 'absolute',
              position: { x: Math.round(L.itens[0].x * 20), y: Math.round((L.base - L.tam * 0.95) * 20) },
              width: Math.round((L.fim - L.itens[0].x + L.tam * 3) * 20),
              height: Math.round(alturaLinha * 20),
              anchor: { horizontal: D.FrameAnchorType.PAGE, vertical: D.FrameAnchorType.PAGE },
              wrap: D.FrameWrap.NONE,
              rule: D.HeightRule.AT_LEAST,
            },
          }));
        }
      }

      const blob = await new Promise((r) => canvas.toBlob(r, 'image/jpeg', 0.92));
      const bytes = new Uint8Array(await blob.arrayBuffer());
      const fundo = new D.Paragraph({
        spacing: { before: 0, after: 0, line: 20, lineRule: D.LineRuleType.EXACT },
        children: [new D.ImageRun({
          data: bytes,
          transformation: { width: Wpt * 96 / 72, height: Hpt * 96 / 72 },
          floating: {
            horizontalPosition: { relative: D.HorizontalPositionRelativeFrom.PAGE, offset: 0 },
            verticalPosition: { relative: D.VerticalPositionRelativeFrom.PAGE, offset: 0 },
            behindDocument: true, allowOverlap: true, lockAnchor: true,
          },
        })],
      });

      secoes.push({
        properties: {
          page: {
            size: { width: Math.round(Wpt * 20), height: Math.round(Hpt * 20), orientation: Wpt > Hpt ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT },
            margin: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0, gutter: 0 },
          },
        },
        children: [fundo, ...filhos],
      });

      if (p <= 3) mostrarPreviaCanvas(canvas, p === 1);
      page.cleanup();
    }

    progresso(1, 'Montando o arquivo Word…');
    await pausa();
    const doc = new D.Document({
      creator: '', title: semExt(arquivo.name),
      styles: { default: { document: { paragraph: { spacing: { before: 0, after: 0 } } } } },
      sections: secoes,
    });
    const blob = await D.Packer.toBlob(doc);
    const nome = semExt(arquivo.name) + '.docx';
    let sub = `${pdf.numPages} página(s) · ${nome}`;
    if (editavel && semTexto === pdf.numPages) sub += ' · Este PDF não tem texto (parece escaneado): as páginas foram mantidas como imagem.';
    concluir('Seu Word está pronto!', sub, 'Baixar Word', () => baixarBlob(blob, nome));
    baixarBlob(blob, nome);
  }

  function mostrarPreviaCanvas(canvas, limpar) {
    if (limpar) el.previaC.innerHTML = '';
    el.previa.hidden = false;
    const c = document.createElement('canvas');
    const k = Math.min(1, 900 / canvas.width);
    c.width = Math.round(canvas.width * k); c.height = Math.round(canvas.height * k);
    c.getContext('2d').drawImage(canvas, 0, 0, c.width, c.height);
    el.previaC.appendChild(c);
  }

  // =====================================================================
  // Word → PDF
  // =====================================================================
  // Valores-padrão do Word quando o arquivo não define (e isola do estilo desta página).
  const BASE_WORD = `.docx-wrapper{font-family:"Times New Roman",Times,serif;font-size:10pt;line-height:normal;color:#000;text-align:left;letter-spacing:normal;word-spacing:normal;-webkit-text-size-adjust:none;text-size-adjust:none}`;
  const pt = (v) => { const m = /([\d.]+)pt/.exec(v || ''); return m ? parseFloat(m[1]) : null; };

  async function wordParaPdf(arquivo) {
    const textoSelecionavel = document.querySelector('input[name="m2"]:checked').value === 'texto';
    const area = el.previaC;
    area.innerHTML = '';
    el.previa.hidden = false;
    progresso(0.15, 'Lendo o documento…');
    // O Word só hifeniza se a opção estiver ligada no arquivo.
    let hifen = false;
    try {
      const zip = await JSZip.loadAsync(arquivo);
      const cfg = zip.file('word/settings.xml');
      if (cfg) hifen = /<w:autoHyphenation(?![^>]*w:val="(false|0|off)")[^>]*\/>/.test(await cfg.async('string'));
    } catch (e) { /* segue sem */ }
    const cssBase = BASE_WORD + (hifen ? '' : '.docx{hyphens:manual !important}');
    await window.DocxPreview.renderAsync(arquivo, area, null, {
      className: 'docx', inWrapper: true, breakPages: true, ignoreLastRenderedPageBreak: false,
      experimental: true, useBase64URL: true, renderHeaders: true, renderFooters: true,
      renderFootnotes: true, renderEndnotes: true, renderChanges: false, ignoreWidth: false, ignoreHeight: false,
    });
    const st = document.createElement('style');
    st.textContent = cssBase;
    area.prepend(st);
    progresso(0.4, 'Carregando fontes e imagens…');
    await esperarRecursos(document, area);

    progresso(0.5, 'Montando as páginas…');
    paginar(area);
    const paginas = Array.from(area.querySelectorAll('section.docx'));
    if (!paginas.length) throw new Error('documento vazio');
    const nome = semExt(arquivo.name) + '.pdf';

    if (textoSelecionavel) {
      progresso(1);
      concluir('Documento pronto!', `Na janela que abrir, escolha "Salvar como PDF" (ou "Microsoft Print to PDF") e salve como ${nome}.`,
        'Salvar PDF', () => imprimir(area, paginas, semExt(arquivo.name)));
      return;
    }

    const { jsPDF } = window.jspdf;
    let pdf = null;
    for (let i = 0; i < paginas.length; i++) {
      progresso(0.4 + 0.6 * (i / paginas.length), `Gerando página ${i + 1} de ${paginas.length}…`);
      await pausa();
      const sec = paginas[i];
      const wPt = pt(sec.style.width) || sec.offsetWidth * 0.75;
      const hPt = pt(sec.style.minHeight) || pt(sec.style.height) || sec.offsetHeight * 0.75;
      const cv = await html2canvas(sec, { scale: 2.5, backgroundColor: '#ffffff', useCORS: true, logging: false });
      const pxPorPt = cv.width / wPt;
      const fatiaPx = Math.round(hPt * pxPorPt);
      for (let y = 0; y < cv.height - 2; y += fatiaPx) {
        const f = document.createElement('canvas');
        f.width = cv.width; f.height = fatiaPx;
        const fc = f.getContext('2d');
        fc.fillStyle = '#fff'; fc.fillRect(0, 0, f.width, f.height);
        fc.drawImage(cv, 0, -y);
        const orient = wPt > hPt ? 'landscape' : 'portrait';
        if (!pdf) pdf = new jsPDF({ unit: 'pt', format: [wPt, hPt], orientation: orient, compress: true });
        else pdf.addPage([wPt, hPt], orient);
        pdf.addImage(f.toDataURL('image/jpeg', 0.95), 'JPEG', 0, 0, wPt, hPt, undefined, 'FAST');
      }
    }
    pdf.setProperties({ title: semExt(arquivo.name) });
    const blob = pdf.output('blob');
    concluir('Seu PDF está pronto!', `${pdf.getNumberOfPages()} página(s) · ${nome}`, 'Baixar PDF', () => baixarBlob(blob, nome));
    baixarBlob(blob, nome);
  }

  // Arquivos sem as marcas de página do Word chegam numa "página" comprida:
  // aqui quebramos entre parágrafos/tabelas, repetindo cabeçalho e rodapé.
  function paginar(area) {
    const fila = Array.from(area.querySelectorAll('section.docx'));
    let guarda = 0;
    while (fila.length && guarda++ < 2000) {
      const sec = fila.shift();
      const hPx = (pt(sec.style.minHeight) || pt(sec.style.height) || 792) / 0.75;
      if (sec.offsetHeight <= hPx + 1) continue;
      const art = sec.querySelector(':scope > article');
      if (!art || art.children.length < 2) continue;
      const cs = getComputedStyle(sec);
      const rodape = sec.querySelector(':scope > footer');
      const limite = sec.getBoundingClientRect().top + hPx - parseFloat(cs.paddingBottom)
        - (rodape ? rodape.offsetHeight : 0);
      const blocos = Array.from(art.children);
      let corte = blocos.findIndex((b, i) => i > 0 && b.getBoundingClientRect().bottom > limite + 0.5);
      if (corte < 1) continue;
      const nova = sec.cloneNode(true);
      const artNova = nova.querySelector(':scope > article');
      artNova.innerHTML = '';
      blocos.slice(corte).forEach((b) => artNova.appendChild(b));
      sec.after(nova);
      fila.unshift(nova);
    }
  }

  async function esperarRecursos(doc, raiz) {
    const tudo = (async () => {
      try { await doc.fonts.ready; } catch (e) { /* ok */ }
      const imgs = Array.from(raiz.querySelectorAll('img'));
      await Promise.all(imgs.map((im) => im.complete ? 0 : new Promise((r) => { im.onload = im.onerror = r; })));
    })();
    await Promise.race([tudo, new Promise((r) => setTimeout(r, 5000))]);
  }

  function imprimir(area, paginas, titulo) {
    const s0 = paginas[0].style;
    const w = pt(s0.width) || 612, h = pt(s0.minHeight) || 792;
    // Se as páginas já vêm do tamanho certo (arquivos salvos pelo Word), imprime 1:1.
    // Senão, deixa o navegador quebrar as páginas usando as margens do documento.
    const cabe = paginas.every((p) => p.scrollHeight * 0.75 <= h + 4);
    const m = ['paddingTop', 'paddingRight', 'paddingBottom', 'paddingLeft'].map((k) => pt(s0[k]) || 72);
    const cssPagina = cabe
      ? `@page { size: ${w}pt ${h}pt; margin: 0; }
         section.docx { margin: 0 !important; box-shadow: none !important; height: ${h}pt !important; overflow: hidden; break-after: page; }`
      : `@page { size: ${w}pt ${h}pt; margin: ${m[0]}pt ${m[1]}pt ${m[2]}pt ${m[3]}pt; }
         section.docx { margin: 0 !important; box-shadow: none !important; padding: 0 !important; width: auto !important; min-height: 0 !important; break-after: page; }`;

    const iframe = document.createElement('iframe');
    iframe.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;';
    document.body.appendChild(iframe);
    const d = iframe.contentDocument;
    d.open();
    d.write(`<!doctype html><html><head><meta charset="utf-8"><title>${titulo.replace(/</g, '&lt;')}</title>
      <style>html,body{margin:0;padding:0;background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact}
      .docx-wrapper{background:#fff !important;padding:0 !important;display:block !important}
      section.docx:last-child{break-after:auto}
      ${cssPagina}</style></head><body>${area.innerHTML}</body></html>`);
    d.close();
    const vai = async () => {
      await esperarRecursos(d, d.body);
      iframe.contentWindow.focus();
      iframe.contentWindow.print();
      setTimeout(() => iframe.remove(), 60000);
    };
    if (d.readyState === 'complete') vai(); else iframe.onload = vai;
  }

  trocarModo('pdf2word');
})();
