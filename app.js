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
  // PDF → Word (motor em pdf2word.js)
  // =====================================================================
  async function pdfParaWord(arquivo) {
    const editavel = document.querySelector('input[name="m1"]:checked').value === 'editavel';
    const dadosArq = new Uint8Array(await arquivo.arrayBuffer());
    const pdf = await pdfjsLib.getDocument({
      data: dadosArq, cMapUrl: 'vendor/cmaps/', cMapPacked: true, standardFontDataUrl: 'vendor/standard_fonts/',
    }).promise;
    const { blob, semTexto } = await window.PdfWord.converter(pdf, {
      editavel, progresso, previa: mostrarPreviaCanvas, titulo: semExt(arquivo.name),
    });
    const nome = semExt(arquivo.name) + '.docx';
    let sub = `${pdf.numPages} página(s) · ${nome}`;
    if (semTexto) sub += ' · Este PDF não tem texto (parece escaneado): as páginas foram mantidas como imagem.';
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
