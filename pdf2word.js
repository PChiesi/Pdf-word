/* PDF → Word: reconstrói o documento em formato nativo do Word
   (parágrafos, títulos, listas, sumário com tabulação, tabelas, cabeçalho e rodapé).
   Tudo roda no navegador. */
(function () {
  'use strict';

  const D = () => window.DocxLib;
  const TW = 20;          // twips por ponto
  const EMU = 12700;      // EMUs por ponto
  const pausa = () => new Promise((r) => setTimeout(r, 0));

  // ---------------------------------------------------------------------
  // Fontes
  // ---------------------------------------------------------------------
  const FONTES = [
    [/^arial.*narrow/i, 'Arial Narrow'], [/^arial/i, 'Arial'], [/^helvetica/i, 'Arial'], [/^liberationsans/i, 'Arial'],
    [/^times|timesnewroman|^liberationserif/i, 'Times New Roman'], [/^courier|^liberationmono/i, 'Courier New'],
    [/^calibri|^carlito/i, 'Calibri'], [/^cambria|^caladea/i, 'Cambria'], [/^candara/i, 'Candara'],
    [/^consolas/i, 'Consolas'], [/^constantia/i, 'Constantia'], [/^corbel/i, 'Corbel'],
    [/^georgia/i, 'Georgia'], [/^verdana/i, 'Verdana'], [/^tahoma/i, 'Tahoma'],
    [/^trebuchet/i, 'Trebuchet MS'], [/^segoeui|^segoe/i, 'Segoe UI'], [/^garamond/i, 'Garamond'],
    [/^bookantiqua/i, 'Book Antiqua'], [/^centurygothic/i, 'Century Gothic'], [/^century/i, 'Century'],
    [/^aptos/i, 'Aptos'], [/^symbol/i, 'Symbol'], [/^wingdings/i, 'Wingdings'],
    [/^roboto/i, 'Roboto'], [/^opensans/i, 'Open Sans'], [/^lato/i, 'Lato'], [/^montserrat/i, 'Montserrat'],
    [/^dejavusans/i, 'DejaVu Sans'],
  ];
  // Altura da linha "simples" do Word, em múltiplos do tamanho da fonte
  const ALTURA_LINHA = { 'Calibri': 1.22, 'Cambria': 1.17, 'Arial': 1.15, 'Times New Roman': 1.15, 'Aptos': 1.2,
    'Segoe UI': 1.33, 'Verdana': 1.22, 'Tahoma': 1.21, 'Georgia': 1.14, 'Courier New': 1.13, 'Garamond': 1.12 };
  const alturaLinha = (fam) => ALTURA_LINHA[fam] || 1.17;

  function infoFonte(page, fontName, estilo) {
    let obj = null;
    try { obj = page.commonObjs.get(fontName); } catch (e) { /* ainda não carregada */ }
    const bruto = ((obj && obj.name) || '').replace(/^[A-Z]{6}\+/, '');
    const [fam, sufixo = ''] = bruto.split(/[-,]/);
    const limpa = (fam || '').replace(/(PS)?MT$/, '').replace(/PS$/, '');
    let familia = null;
    for (const [re, nome] of FONTES) if (re.test(limpa)) { familia = nome; break; }
    if (!familia && limpa) familia = limpa.replace(/([a-z])([A-Z])/g, '$1 $2');
    const generica = (estilo && estilo.fontFamily) || 'sans-serif';
    if (!familia) familia = generica === 'serif' ? 'Times New Roman' : generica === 'monospace' ? 'Courier New' : 'Arial';
    const s = (sufixo + ' ' + bruto).toLowerCase();
    const negrito = !!(obj && (obj.bold || obj.black)) || /bold|black|heavy|semibold|demi/.test(s);
    const italico = !!(obj && obj.italic) || /italic|oblique/.test(s);
    return { familia, generica, negrito, italico };
  }

  const medidor = document.createElement('canvas').getContext('2d');
  function medir(texto, f, tam) {
    medidor.font = `${f.italico ? 'italic ' : ''}${f.negrito ? 'bold ' : ''}${tam}px "${f.familia}", ${f.generica}`;
    return medidor.measureText(texto).width;
  }

  // ---------------------------------------------------------------------
  // Cores: fundo (mediana da borda) e tinta (cor mais frequente)
  // ---------------------------------------------------------------------
  function analisarCaixa(dados, W, H, x0, y0, x1, y1) {
    x0 = Math.max(0, Math.floor(x0)); y0 = Math.max(0, Math.floor(y0));
    x1 = Math.min(W - 1, Math.ceil(x1)); y1 = Math.min(H - 1, Math.ceil(y1));
    if (x1 <= x0 || y1 <= y0) return null;
    const r = [], g = [], b = [];
    const amostra = (x, y) => { const i = (y * W + x) * 4; r.push(dados[i]); g.push(dados[i + 1]); b.push(dados[i + 2]); };
    const px = Math.max(1, Math.floor((x1 - x0) / 60));
    for (let x = x0; x <= x1; x += px) { amostra(x, y0); amostra(x, y1); }
    for (let y = y0; y <= y1; y++) { amostra(x0, y); amostra(x1, y); }
    const med = (v) => v.sort((p, q) => p - q)[v.length >> 1];
    const fundo = [med(r), med(g), med(b)];
    const passo = Math.max(1, Math.floor(Math.sqrt(((x1 - x0) * (y1 - y0)) / 4000)));
    let maxD = 0;
    const dist = (i) => Math.abs(dados[i] - fundo[0]) + Math.abs(dados[i + 1] - fundo[1]) + Math.abs(dados[i + 2] - fundo[2]);
    for (let y = y0; y <= y1; y += passo) for (let x = x0; x <= x1; x += passo) maxD = Math.max(maxD, dist((y * W + x) * 4));
    const bins = new Map();
    for (let y = y0; y <= y1; y += passo) for (let x = x0; x <= x1; x += passo) {
      const i = (y * W + x) * 4;
      if (dist(i) < maxD * 0.6) continue;
      const k = (dados[i] >> 4) << 8 | (dados[i + 1] >> 4) << 4 | (dados[i + 2] >> 4);
      const v = bins.get(k) || [0, 0, 0, 0];
      v[0] += dados[i]; v[1] += dados[i + 1]; v[2] += dados[i + 2]; v[3]++;
      bins.set(k, v);
    }
    let m = null;
    for (const v of bins.values()) if (!m || v[3] > m[3]) m = v;
    const cor = m && maxD > 30 ? [m[0] / m[3], m[1] / m[3], m[2] / m[3]] : [0, 0, 0];
    return { fundo, cor };
  }
  const hex = (c) => c.map((v) => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase();
  const quaseBranco = (c) => c[0] > 245 && c[1] > 245 && c[2] > 245;

  // ---------------------------------------------------------------------
  // Utilidades de texto
  // ---------------------------------------------------------------------
  const RE_LIDER = /^[\s.·•…_\-–]*$/;            // só pontinhos/tracinhos (linha guia do sumário)
  const RE_MARCADOR = /^[•●○◦▪■□►▸‣⁃–\-*]$/;
  const RE_NUM_LISTA = /^(\d{1,3}[.)]|[a-zA-Z][.)]|[ivxIVX]{1,4}[.)])$/;
  const RE_NUM_PAG = /^(\d{1,4}|[ivxlcdm]{1,6})$/i;
  const chave = (t) => t.toLowerCase().replace(/\d+/g, '#').replace(/\s+/g, ' ').trim();
  const percentil = (v, p) => { if (!v.length) return null; const s = [...v].sort((a, b) => a - b); return s[Math.min(s.length - 1, Math.max(0, Math.round((s.length - 1) * p)))]; };
  const mediana = (v) => percentil(v, 0.5);

  // Separa "Título ........ 3" que às vezes vem num único pedaço
  function dividirLider(it) {
    const m = /^(.*?\S)\s*([.·…_]{4,}[.·…_\s]*)(\S.*)?$/.exec(it.str);
    if (!m || it.str.length < 5) return [it];
    const n = it.str.length, k = it.largura / n;
    const partes = [];
    const a = m[1], l = m[2], c = m[3] || '';
    partes.push({ ...it, str: a, largura: a.length * k });
    partes.push({ ...it, str: l, x: it.x + a.length * k, largura: l.length * k, lider: true });
    if (c) partes.push({ ...it, str: c, x: it.x + (a.length + l.length) * k, largura: c.length * k });
    return partes;
  }

  // ---------------------------------------------------------------------
  // 1) Leitura de uma página: texto, cores, linhas e áreas gráficas
  // ---------------------------------------------------------------------
  async function lerPagina(page, num) {
    const vp = page.getViewport({ scale: 1 });
    const W = vp.width, H = vp.height;
    const S = Math.min(200 / 72, 5000 / W, 5000 / H);
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(W * S); canvas.height = Math.round(H * S);
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport: page.getViewport({ scale: S }) }).promise;
    const previa = canvas;

    const tc = await page.getTextContent();
    let itens = [];
    for (const it of tc.items) {
      if (!it.str || !it.str.trim()) continue;
      const t = pdfjsLib.Util.transform(vp.transform, it.transform);
      const tam = Math.hypot(t[2], t[3]);
      if (tam < 1 || !(t[0] > 0 && Math.abs(t[1]) < 0.01 * tam && Math.abs(t[2]) < 0.01 * tam)) continue; // texto girado fica como imagem
      const est = tc.styles[it.fontName] || {};
      const asc = est.ascent && est.ascent > 0.3 ? est.ascent : 0.8;
      const desc = est.descent && est.descent < 0 ? -est.descent : 0.2;
      const largura = it.width > 0 ? it.width : tam * it.str.length * 0.5;
      const base = { str: it.str, x: t[4], base: t[5], tam, largura, topo: t[5] - tam * asc, fundo: t[5] + tam * desc, fonte: infoFonte(page, it.fontName, est) };
      itens.push(...dividirLider(base));
    }

    // Cores do texto e do fundo; depois "apaga" o texto da imagem
    const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const pad = 1.5;
    for (const it of itens) {
      const a = analisarCaixa(img.data, canvas.width, canvas.height, it.x * S, it.topo * S - pad, (it.x + it.largura) * S, it.fundo * S + pad);
      it.cor = a ? hex(a.cor) : '000000';
      it.fundoCor = a ? a.fundo : [255, 255, 255];
    }
    const copia = document.createElement('canvas');
    copia.width = canvas.width; copia.height = canvas.height;
    const cc = copia.getContext('2d', { willReadFrequently: true });
    cc.drawImage(canvas, 0, 0);
    for (const it of itens) {
      cc.fillStyle = `rgb(${it.fundoCor.map(Math.round).join(',')})`;
      cc.fillRect(it.x * S - pad, it.topo * S - pad, it.largura * S + 2 * pad, (it.fundo - it.topo) * S + 2 * pad);
    }

    const linhas = agruparLinhas(itens);
    const graficos = await areasGraficas(copia, S, W, H);
    page.cleanup();
    return { num, W, H, linhas, graficos, previa, temTexto: itens.length > 0 };
  }

  function agruparLinhas(itens) {
    itens.sort((a, b) => a.topo - b.topo || a.x - b.x);
    const linhas = [];
    for (const it of itens) {
      const centro = (it.topo + it.fundo) / 2;
      let alvo = null;
      for (const L of linhas) {
        if (centro > L.topo && centro < L.fundo && Math.abs(L.base - it.base) < 0.8 * Math.max(L.tam, it.tam)) { alvo = L; break; }
      }
      if (alvo) {
        alvo.itens.push(it);
        alvo.topo = Math.min(alvo.topo, it.topo); alvo.fundo = Math.max(alvo.fundo, it.fundo);
      } else linhas.push({ itens: [it], topo: it.topo, fundo: it.fundo, base: it.base, tam: it.tam });
    }
    for (const L of linhas) prepararLinha(L);
    linhas.sort((a, b) => a.base - b.base || a.x0 - b.x0);
    return linhas;
  }

  function prepararLinha(L) {
    L.itens.sort((a, b) => a.x - b.x);
    // Tamanho dominante (por nº de caracteres) e linha de base
    const porTam = new Map();
    for (const it of L.itens) { const k = Math.round(it.tam * 2) / 2; porTam.set(k, (porTam.get(k) || 0) + it.str.length); }
    L.tam = [...porTam.entries()].sort((a, b) => b[1] - a[1])[0][0];
    const principal = L.itens.find((it) => Math.abs(it.tam - L.tam) < 0.6) || L.itens[0];
    L.base = principal.base;
    L.fonte = principal.fonte;
    for (const it of L.itens) it.sobrescrito = it.tam < L.tam * 0.8 && it.base < L.base - 0.2 * L.tam;
    L.x0 = Math.min(...L.itens.map((i) => i.x));
    L.x1 = Math.max(...L.itens.map((i) => i.x + i.largura));
    // Segmentos: blocos separados por espaço grande ou por linha guia
    L.segs = [];
    let seg = null, liderAntes = false;
    for (const it of L.itens) {
      if (it.lider || (RE_LIDER.test(it.str) && it.str.replace(/\s/g, '').length >= 3)) { liderAntes = true; continue; }
      if (!seg || liderAntes || it.x - seg.x1 > 2.2 * L.tam) {
        seg = { itens: [], x0: it.x, x1: it.x + it.largura, lider: liderAntes };
        L.segs.push(seg);
      }
      seg.itens.push(it); seg.x1 = Math.max(seg.x1, it.x + it.largura);
      liderAntes = false;
    }
    for (const s of L.segs) s.texto = textoItens(s.itens);
    L.texto = L.segs.map((s) => s.texto).join(' ');
  }

  function textoItens(itens) {
    let t = '', fim = null;
    for (const it of itens) {
      if (fim !== null && it.x - fim > 0.12 * it.tam && !/\s$/.test(t) && !/^\s/.test(it.str)) t += ' ';
      t += it.str; fim = it.x + it.largura;
    }
    return t.replace(/\s+/g, ' ').trim();
  }

  // Áreas com desenho (imagens, linhas, fundos) depois de apagar o texto
  async function areasGraficas(canvas, S, W, H) {
    const cw = canvas.width, ch = canvas.height;
    const dados = canvas.getContext('2d').getImageData(0, 0, cw, ch).data;
    const cel = Math.max(2, Math.round(2.5 * S));
    const gx = Math.ceil(cw / cel), gy = Math.ceil(ch / cel);
    const grade = new Uint8Array(gx * gy);
    let marcadas = 0;
    for (let j = 0; j < gy; j++) for (let i = 0; i < gx; i++) {
      let tinta = false;
      const yMax = Math.min(ch, (j + 1) * cel), xMax = Math.min(cw, (i + 1) * cel);
      for (let y = j * cel; y < yMax && !tinta; y++) for (let x = i * cel; x < xMax; x++) {
        const k = (y * cw + x) * 4;
        if (765 - dados[k] - dados[k + 1] - dados[k + 2] > 45) { tinta = true; break; }
      }
      if (tinta) { grade[j * gx + i] = 1; marcadas++; }
    }
    if (!marcadas) return [];
    // Componentes conectados
    const vis = new Uint8Array(gx * gy);
    let caixas = [];
    for (let s = 0; s < grade.length; s++) {
      if (!grade[s] || vis[s]) continue;
      let x0 = 1e9, y0 = 1e9, x1 = -1, y1 = -1, n = 0;
      const pilha = [s]; vis[s] = 1;
      while (pilha.length) {
        const c = pilha.pop(); const cx = c % gx, cy = (c / gx) | 0; n++;
        x0 = Math.min(x0, cx); x1 = Math.max(x1, cx); y0 = Math.min(y0, cy); y1 = Math.max(y1, cy);
        for (let dy = -1; dy <= 1; dy++) for (let dx = -1; dx <= 1; dx++) {
          const nx = cx + dx, ny = cy + dy;
          if (nx < 0 || ny < 0 || nx >= gx || ny >= gy) continue;
          const v = ny * gx + nx;
          if (grade[v] && !vis[v]) { vis[v] = 1; pilha.push(v); }
        }
      }
      caixas.push({ x0: x0 * cel, y0: y0 * cel, x1: Math.min(cw, (x1 + 1) * cel), y1: Math.min(ch, (y1 + 1) * cel), n });
    }
    // Junta caixas que se sobrepõem
    let mudou = true;
    while (mudou) {
      mudou = false;
      for (let a = 0; a < caixas.length && !mudou; a++) for (let b = a + 1; b < caixas.length; b++) {
        const A = caixas[a], B = caixas[b];
        if (A.x0 <= B.x1 && B.x0 <= A.x1 && A.y0 <= B.y1 && B.y0 <= A.y1) {
          A.x0 = Math.min(A.x0, B.x0); A.y0 = Math.min(A.y0, B.y0); A.x1 = Math.max(A.x1, B.x1); A.y1 = Math.max(A.y1, B.y1); A.n += B.n;
          caixas.splice(b, 1); mudou = true; break;
        }
      }
    }
    caixas = caixas.filter((c) => (c.x1 - c.x0) / S >= 4 || (c.y1 - c.y0) / S >= 4);
    const out = [];
    for (const c of caixas) {
      const w = c.x1 - c.x0, h = c.y1 - c.y0;
      const recorte = document.createElement('canvas');
      recorte.width = w; recorte.height = h;
      recorte.getContext('2d').drawImage(canvas, c.x0, c.y0, w, h, 0, 0, w, h);
      const tipo = w * h > 400000 ? 'image/jpeg' : 'image/png';
      const blob = await new Promise((r) => recorte.toBlob(r, tipo, 0.9));
      out.push({ x: c.x0 / S, y: c.y0 / S, w: w / S, h: h / S, bytes: new Uint8Array(await blob.arrayBuffer()) });
    }
    return out;
  }

  // ---------------------------------------------------------------------
  // 2) Cabeçalho e rodapé (linhas que se repetem no topo/rodapé das páginas)
  // ---------------------------------------------------------------------
  const RE_PAG = /^((p[aá]g(ina)?|page|p)\.?\s*)?\d{1,4}(\s*(de|of|\/|-)\s*\d{1,4})?$/i;

  function detectarCabecalhoRodape(paginas) {
    const N = paginas.length;
    const zona = (pg, L) => {
      const lim = Math.min(pg.H * 0.13, 110);
      if (L.fundo < lim) return 'cab';
      if (L.topo > pg.H - lim) return 'rod';
      return null;
    };
    const cont = { cab: new Map(), rod: new Map() };
    for (const pg of paginas) {
      const vistos = new Set();
      for (const L of pg.linhas) {
        const z = zona(pg, L); if (!z) continue;
        const k = z + '|' + chave(L.texto);
        if (vistos.has(k)) continue; vistos.add(k);
        cont[z].set(k, (cont[z].get(k) || 0) + 1);
      }
    }
    const minimo = Math.max(2, Math.ceil(N * 0.5));
    for (const pg of paginas) {
      for (const L of pg.linhas) {
        const z = zona(pg, L); if (!z) continue;
        const k = z + '|' + chave(L.texto);
        const repetida = N >= 2 && cont[z].get(k) >= minimo;
        if (repetida || RE_PAG.test(L.texto.trim())) L.hf = z;
      }
    }
  }

  // ---------------------------------------------------------------------
  // 3) Margens do corpo do texto
  // ---------------------------------------------------------------------
  function calcularMargens(paginas) {
    const xs0 = [], xs1 = [], tops = [], bots = [];
    for (const pg of paginas) {
      const corpo = pg.linhas.filter((L) => !L.hf);
      for (const L of corpo) { xs0.push(L.x0); xs1.push(L.x1); }
      if (corpo.length) { tops.push(Math.min(...corpo.map((L) => L.topo))); bots.push(Math.max(...corpo.map((L) => L.fundo))); }
    }
    const W = paginas[0].W, H = paginas[0].H;
    const ML = xs0.length ? percentil(xs0, 0.05) : 72;
    let MR = xs1.length ? percentil(xs1, 0.95) : W - 72;
    if (MR - ML < W * 0.3) MR = W - ML;
    const MT = tops.length ? percentil(tops, 0.1) : 72;
    const MB = bots.length ? percentil(bots, 0.95) : H - 72;
    const cab = [], rod = [];
    for (const pg of paginas) for (const L of pg.linhas) {
      if (L.hf === 'cab') cab.push(L.topo);
      if (L.hf === 'rod') rod.push(pg.H - L.fundo);
    }
    return {
      ML, MR, MT, MB,
      cabDist: cab.length ? Math.max(12, Math.min(...cab) - 2) : Math.min(36, MT / 2),
      rodDist: rod.length ? Math.max(12, Math.min(...rod) - 2) : Math.min(36, (H - MB) / 2),
    };
  }

  // ---------------------------------------------------------------------
  // 4) Classificação das linhas e montagem de parágrafos
  // ---------------------------------------------------------------------
  function classificar(L, M) {
    const segs = L.segs;
    if (!segs.length) { L.tipo = 'vazia'; return; }
    const ult = segs[segs.length - 1];
    if (segs.length >= 2 && RE_NUM_PAG.test(ult.texto) && (ult.lider || ult.x1 > M.MR - 3 * L.tam)) { L.tipo = 'sumario'; return; }
    if (segs.length >= 2) { L.tipo = 'tab'; return; }
    // Lista: marcador sozinho seguido de texto
    const it = L.itens;
    if (it.length >= 2) {
      const prim = it[0].str.trim();
      const prox = it[1];
      const gap = prox.x - (it[0].x + it[0].largura);
      if (RE_MARCADOR.test(prim) || (RE_NUM_LISTA.test(prim) && gap > 0.6 * L.tam)) {
        L.tipo = 'lista'; L.marcador = it[0]; L.textoX = prox.x; return;
      }
    }
    const m = /^([•●○◦▪■►–\-*])\s+/.exec(it[0].str);
    if (m) {
      L.tipo = 'lista';
      const k = it[0].largura / it[0].str.length;
      L.marcador = { ...it[0], str: m[1], largura: k };
      L.textoX = it[0].x + m[0].length * k;
      it[0] = { ...it[0], str: it[0].str.slice(m[0].length), x: L.textoX, largura: it[0].largura - m[0].length * k };
      return;
    }
    L.tipo = 'texto';
  }

  function palavraInicial(L) {
    const it = L.itens.find((i) => !L.marcador || i !== L.marcador) || L.itens[0];
    const w = (it.str.trim().split(/\s+/)[0] || '');
    return (w.length / Math.max(1, it.str.length)) * it.largura;
  }

  function montarParagrafos(pg, M) {
    const blocos = [];
    const corpo = pg.linhas.filter((L) => !L.hf);
    corpo.forEach((L) => classificar(L, M));
    let P = null;
    const fecha = () => { if (P) blocos.push(P); P = null; };
    for (let i = 0; i < corpo.length; i++) {
      const L = corpo[i];
      if (L.tipo === 'vazia') continue;
      if (L.tipo === 'sumario' || L.tipo === 'tab') {
        fecha();
        // Linhas tabuladas seguidas viram tabela
        if (L.tipo === 'tab') {
          const grupo = [L];
          while (i + 1 < corpo.length && corpo[i + 1].tipo === 'tab' && corpo[i + 1].segs.length === L.segs.length
            && corpo[i + 1].base - grupo[grupo.length - 1].base < 3.5 * L.tam) grupo.push(corpo[++i]);
          if (grupo.length >= 2 && colunasAlinhadas(grupo)) { blocos.push({ tipo: 'tabela', linhas: grupo }); continue; }
          i -= grupo.length - 1;
        }
        blocos.push({ tipo: L.tipo, linhas: [L] });
        continue;
      }
      if (P && podeJuntar(P, L, M)) { P.linhas.push(L); continue; }
      fecha();
      P = { tipo: L.tipo === 'lista' ? 'lista' : 'texto', linhas: [L] };
    }
    fecha();
    return blocos;
  }

  function colunasAlinhadas(grupo) {
    const n = grupo[0].segs.length;
    for (let c = 0; c < n; c++) {
      const x0s = grupo.map((L) => L.segs[c].x0), x1s = grupo.map((L) => L.segs[c].x1);
      const esq = Math.max(...x0s) - Math.min(...x0s) < 2 * grupo[0].tam;
      const dir = Math.max(...x1s) - Math.min(...x1s) < 2 * grupo[0].tam;
      if (!esq && !dir) return false;
    }
    return true;
  }

  function podeJuntar(P, L, M) {
    if (L.tipo !== 'texto') return false;
    const ult = P.linhas[P.linhas.length - 1];
    if (Math.abs(L.tam - ult.tam) > 0.6) return false;
    if (L.fonte.negrito !== ult.fonte.negrito && P.linhas.length === 1 && L.itens.every((i) => i.fonte.negrito === L.fonte.negrito)) return false;
    const gap = L.base - ult.base;
    if (gap < 0.9 * L.tam || gap > 1.75 * L.tam) return false;
    if (P.linhas.length >= 2) {
      const passo = P.linhas[1].base - P.linhas[0].base;
      if (Math.abs(gap - passo) > 0.2 * passo) return false;
    }
    // Se a primeira palavra desta linha cabia na anterior, a anterior terminou o parágrafo
    const direita = Math.max(M.MR, ...P.linhas.map((x) => x.x1));
    const cabe = ult.x1 + 0.3 * L.tam + palavraInicial(L) < direita - 2;
    const centrada = (x) => Math.abs((x.x0 + x.x1) / 2 - (M.ML + M.MR) / 2) < 4 && x.x0 > M.ML + 2 * x.tam;
    if (centrada(ult) && centrada(L)) return !cabe || true;
    if (cabe) return false;
    // Alinhamento à esquerda coerente
    const ref = P.tipo === 'lista' ? P.linhas[0].textoX : (P.linhas.length >= 2 ? P.linhas[1].x0 : null);
    if (ref !== null) return Math.abs(L.x0 - ref) < 1.5;
    return L.x0 <= P.linhas[0].x0 + 1.5 && L.x0 >= P.linhas[0].x0 - 6 * L.tam;
  }

  // ---------------------------------------------------------------------
  // 5) Geração do Word
  // ---------------------------------------------------------------------
  function estiloRun(it, L) {
    const tam = Math.round(it.tam * 2) / 2;
    return {
      font: it.fonte.familia, size: Math.max(2, Math.round((it.sobrescrito ? L.tam : tam) * 2)),
      bold: it.fonte.negrito, italics: it.fonte.italico, color: it.cor,
      superScript: it.sobrescrito || undefined,
    };
  }

  // Converte itens em runs, juntando trechos de mesmo estilo
  function runsDe(linhas, opts = {}) {
    const pedacos = []; // {texto, estilo}
    const add = (texto, est) => {
      const ult = pedacos[pedacos.length - 1];
      if (ult && JSON.stringify(ult.est) === JSON.stringify(est)) ult.texto += texto;
      else pedacos.push({ texto, est });
    };
    linhas.forEach((L, li) => {
      const itens = (opts.itens && opts.itens(L)) || L.itens.filter((i) => i !== L.marcador && !i.lider && !(RE_LIDER.test(i.str) && i.str.replace(/\s/g, '').length >= 3));
      let fim = null;
      if (li > 0 && pedacos.length) {
        const u = pedacos[pedacos.length - 1];
        const primeira = (itens[0] && itens[0].str) || '';
        if (/[A-Za-zÀ-ÿ]-$/.test(u.texto) && /^[a-zà-ÿ]/.test(primeira)) u.texto = u.texto.slice(0, -1); // hifenização
        else if (!/\s$/.test(u.texto)) u.texto += ' ';
      }
      for (const it of itens) {
        const est = estiloRun(it, L);
        if (fim !== null && it.x - fim > 0.12 * it.tam && pedacos.length && !/\s$/.test(pedacos[pedacos.length - 1].texto) && !/^\s/.test(it.str)) {
          pedacos[pedacos.length - 1].texto += ' ';
        }
        add(it.str, est);
        fim = it.x + it.largura;
      }
    });
    if (pedacos.length) pedacos[pedacos.length - 1].texto = pedacos[pedacos.length - 1].texto.replace(/\s+$/, '');
    return pedacos.map((p) => new (D().TextRun)({ text: p.texto.replace(/ {2,}/g, ' '), ...p.est }));
  }

  function espacamentoLinhas(P) {
    const L0 = P.linhas[0];
    if (P.linhas.length < 2) return {};
    const passos = [];
    for (let i = 1; i < P.linhas.length; i++) passos.push(P.linhas[i].base - P.linhas[i - 1].base);
    const passo = mediana(passos);
    let line = Math.round(240 * passo / (L0.tam * alturaLinha(L0.fonte.familia)));
    if (Math.abs(line - 240) <= 10) line = 240;
    line = Math.max(180, Math.min(720, line));
    return { line, lineRule: D().LineRuleType.AUTO };
  }

  function passoDe(P) {
    const L0 = P.linhas[0];
    if (P.linhas.length >= 2) return P.linhas[1].base - P.linhas[0].base;
    return L0.tam * alturaLinha(L0.fonte.familia);
  }

  function alinhamento(P, M) {
    const A = D().AlignmentType;
    const ls = P.linhas;
    const centro = (M.ML + M.MR) / 2;
    const centrada = ls.every((L) => Math.abs((L.x0 + L.x1) / 2 - centro) < Math.max(4, 0.02 * (M.MR - M.ML)) && L.x0 > M.ML + 2 * L.tam);
    if (centrada) return { alignment: A.CENTER };
    const aDireita = ls.every((L) => Math.abs(L.x1 - M.MR) < 3 && L.x0 > M.ML + (M.MR - M.ML) * 0.3);
    if (aDireita) return { alignment: A.RIGHT };
    if (ls.length >= 2) {
      const meio = ls.slice(0, -1);
      if (meio.every((L) => Math.abs(L.x1 - M.MR) < 2) && meio.some((L) => L.texto.includes(' '))) return { alignment: A.JUSTIFIED };
    }
    return { alignment: A.LEFT };
  }

  function recuos(P, M) {
    const ls = P.linhas;
    const tw = (v) => Math.round(Math.max(0, v) * TW);
    if (P.tipo === 'lista') {
      const L0 = ls[0];
      return { left: tw(L0.textoX - M.ML), hanging: tw(L0.textoX - L0.marcador.x) };
    }
    const outros = ls.length >= 2 ? Math.min(...ls.slice(1).map((L) => L.x0)) : ls[0].x0;
    const left = outros - M.ML;
    const first = ls[0].x0 - outros;
    const r = {};
    if (left > 1.5) r.left = tw(left);
    if (first > 1.5) r.firstLine = tw(first);
    else if (first < -1.5) r.hanging = tw(-first);
    return r;
  }

  function paragrafoTexto(P, M, antes, extras) {
    const Dx = D();
    const al = alinhamento(P, M);
    const ind = al.alignment === Dx.AlignmentType.CENTER || al.alignment === Dx.AlignmentType.RIGHT ? {} : recuos(P, M);
    const filhos = [];
    if (P.tipo === 'lista') {
      const mk = P.linhas[0].marcador;
      const est = estiloRun(mk, P.linhas[0]);
      filhos.push(new Dx.TextRun({ ...est, children: [mk.str.trim(), new Dx.Tab()] }));
    }
    filhos.push(...runsDe(P.linhas));
    const op = {
      children: [...(extras || []), ...filhos],
      ...al, indent: Object.keys(ind).length ? ind : undefined,
      spacing: { before: Math.round(Math.max(0, antes) * TW), after: 0, ...espacamentoLinhas(P) },
    };
    if (P.nivel) op.heading = [Dx.HeadingLevel.HEADING_1, Dx.HeadingLevel.HEADING_2, Dx.HeadingLevel.HEADING_3][P.nivel - 1];
    if (P.quebra) op.pageBreakBefore = true;
    return new Dx.Paragraph(op);
  }

  function paragrafoTabulado(P, M, antes, extras) {
    const Dx = D();
    const L = P.linhas[0];
    const tabs = [];
    const filhos = [];
    L.segs.forEach((s, i) => {
      if (i > 0) {
        const ultimo = i === L.segs.length - 1;
        const numerico = /^[\d\s.,%R$()+\-–]+$/.test(s.texto);
        if (P.tipo === 'sumario' && ultimo) tabs.push({ type: Dx.TabStopType.RIGHT, position: Math.round((s.x1 - M.ML) * TW), leader: s.lider ? Dx.LeaderType.DOT : undefined });
        else if (numerico) tabs.push({ type: Dx.TabStopType.RIGHT, position: Math.round((s.x1 - M.ML) * TW) });
        else tabs.push({ type: Dx.TabStopType.LEFT, position: Math.round((s.x0 - M.ML) * TW) });
        filhos.push(new Dx.TextRun({ ...estiloRun(s.itens[0], L), children: [new Dx.Tab()] }));
      }
      filhos.push(...runsDe([L], { itens: () => s.itens }));
    });
    const left = L.segs[0].x0 - M.ML;
    return new Dx.Paragraph({
      children: [...(extras || []), ...filhos],
      tabStops: tabs,
      indent: left > 1.5 ? { left: Math.round(left * TW) } : undefined,
      spacing: { before: Math.round(Math.max(0, antes) * TW), after: 0 },
      pageBreakBefore: P.quebra || undefined,
    });
  }

  function tabelaNativa(P, M, graficos) {
    const Dx = D();
    const ls = P.linhas;
    const n = ls[0].segs.length;
    const PAD = 5.4; // margem interna padrão da célula no Word (0,08")
    const inicio = [];
    for (let c = 0; c < n; c++) inicio.push(Math.min(...ls.map((L) => L.segs[c].x0)) - PAD);
    const fimTab = Math.max(...ls.map((L) => L.segs[n - 1].x1)) + PAD;
    const topo = Math.min(...ls.map((L) => L.topo)) - 4, base = Math.max(...ls.map((L) => L.fundo)) + 4;
    // Se há desenho (bordas/fundos) sobre a tabela, usa bordas nativas e tira o desenho
    let x0 = inicio[0], x1 = fimTab, comBorda = false, bordaTopo = null, bordaFundo = null;
    for (let g = graficos.length - 1; g >= 0; g--) {
      const G = graficos[g];
      const sobrep = G.x < x1 && G.x + G.w > x0 && G.y < base && G.y + G.h > topo;
      const dentro = G.x >= x0 - 12 && G.x + G.w <= Math.max(x1, M.MR) + 12 && G.y >= topo - 12 && G.y + G.h <= base + 12;
      if (sobrep && dentro) {
        comBorda = true;
        x0 = Math.min(x0, G.x); x1 = Math.max(x1, G.x + G.w);
        bordaTopo = bordaTopo === null ? G.y : Math.min(bordaTopo, G.y);
        bordaFundo = bordaFundo === null ? G.y + G.h : Math.max(bordaFundo, G.y + G.h);
        graficos.splice(g, 1);
      }
    }
    if (comBorda) { inicio[0] = Math.min(inicio[0], x0); }
    const fim = comBorda ? Math.max(fimTab, x1) : fimTab;
    const larguras = inicio.map((x, c) => ((c + 1 < n ? inicio[c + 1] : fim) - x));
    // Altura das linhas da tabela igual à do PDF
    const passos = [];
    for (let i = 1; i < ls.length; i++) passos.push(ls[i].base - ls[i - 1].base);
    let passo = passos.length ? mediana(passos) : (ls[0].fundo - ls[0].topo) + 6;
    let tTopo = ls[0].topo - (passo - (ls[0].fundo - ls[0].topo)) / 2;
    if (comBorda && bordaTopo !== null) { tTopo = bordaTopo; passo = (bordaFundo - bordaTopo) / ls.length; }
    const linhaBorda = comBorda ? { style: Dx.BorderStyle.SINGLE, size: 4, color: '000000' } : { style: Dx.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const borders = { top: linhaBorda, bottom: linhaBorda, left: linhaBorda, right: linhaBorda, insideHorizontal: linhaBorda, insideVertical: linhaBorda };
    const rows = ls.map((L) => new Dx.TableRow({
      height: { value: Math.round(passo * TW), rule: Dx.HeightRule.ATLEAST },
      children: L.segs.map((s, c) => {
        const fundo = s.itens[0].fundoCor;
        const numerico = /^[\d\s.,%R$()+\-–]+$/.test(s.texto) && c > 0;
        return new Dx.TableCell({
          verticalAlign: Dx.VerticalAlign.CENTER,
          width: { size: Math.round(larguras[c] * TW), type: Dx.WidthType.DXA },
          shading: quaseBranco(fundo) ? undefined : { fill: hex(fundo), type: Dx.ShadingType.CLEAR, color: 'auto' },
          children: [new Dx.Paragraph({
            alignment: numerico && Math.abs(s.x1 - Math.max(...ls.map((x) => x.segs[c].x1))) < 2 && Math.abs(s.x0 - Math.min(...ls.map((x) => x.segs[c].x0))) > 2 ? Dx.AlignmentType.RIGHT : Dx.AlignmentType.LEFT,
            spacing: { before: 0, after: 0 },
            children: runsDe([L], { itens: () => s.itens }),
          })],
        });
      }),
    }));
    const recuo = inicio[0] - M.ML + PAD;
    P.topo = tTopo; P.fundo = tTopo + passo * ls.length;
    return new Dx.Table({
      rows, borders,
      columnWidths: larguras.map((w) => Math.round(w * TW)),
      width: { size: Math.round(larguras.reduce((a, b) => a + b, 0) * TW), type: Dx.WidthType.DXA },
      indent: Math.abs(recuo) > 1 ? { size: Math.round(recuo * TW), type: Dx.WidthType.DXA } : undefined,
      layout: Dx.TableLayoutType.FIXED,
    });
  }

  function imagensFlutuantes(graficos) {
    const Dx = D();
    return graficos.map((G) => new Dx.ImageRun({
      data: G.bytes,
      transformation: { width: G.w * 96 / 72, height: G.h * 96 / 72 },
      floating: {
        horizontalPosition: { relative: Dx.HorizontalPositionRelativeFrom.PAGE, offset: Math.round(G.x * EMU) },
        verticalPosition: { relative: Dx.VerticalPositionRelativeFrom.PAGE, offset: Math.round(G.y * EMU) },
        behindDocument: true, allowOverlap: true, lockAnchor: true,
      },
    }));
  }

  // Cabeçalho/rodapé: monta a partir de uma página representativa
  function montarHF(paginas, z, M, total) {
    const Dx = D();
    const rep = paginas.find((p, i) => i > 0 && p.linhas.some((L) => L.hf === z)) || paginas.find((p) => p.linhas.some((L) => L.hf === z));
    if (!rep) return null;
    const linhas = rep.linhas.filter((L) => L.hf === z);
    const paras = linhas.map((L, idx) => {
      const segs = L.segs.length ? L.segs : [{ itens: L.itens, x0: L.x0, x1: L.x1, texto: L.texto }];
      const filhos = [];
      const tabs = [];
      const centro = (M.ML + M.MR) / 2;
      const estado = { pag: false, total: false };
      // Só troca por campo se o número muda de uma página para outra (não é um ano, por exemplo)
      const rotuloPag = /\d/.test(L.texto) && paginas.filter((p) => p.linhas.some((x) => x.hf === z && chave(x.texto) === chave(L.texto) && x.texto !== L.texto)).length > 0
        || RE_PAG.test(L.texto.trim());
      let al = Dx.AlignmentType.LEFT;
      if (segs.length === 1) {
        const s = segs[0];
        if (Math.abs((s.x0 + s.x1) / 2 - centro) < 6) al = Dx.AlignmentType.CENTER;
        else if (Math.abs(s.x1 - M.MR) < 6 || s.x0 > centro) al = Dx.AlignmentType.RIGHT;
      }
      segs.forEach((s, i) => {
        if (i > 0) {
          const mid = (s.x0 + s.x1) / 2;
          if (Math.abs(mid - centro) < 20) tabs.push({ type: Dx.TabStopType.CENTER, position: Math.round((centro - M.ML) * TW) });
          else if (s.x1 > M.MR - 20) tabs.push({ type: Dx.TabStopType.RIGHT, position: Math.round((M.MR - M.ML) * TW) });
          else tabs.push({ type: Dx.TabStopType.LEFT, position: Math.round((s.x0 - M.ML) * TW) });
          filhos.push(new Dx.TextRun({ children: [new Dx.Tab()] }));
        }
        let fim = null;
        for (const it of s.itens) {
          const est = estiloRun(it, L);
          let str = it.str;
          if (fim !== null && it.x - fim > 0.12 * it.tam && !/^\s/.test(str)) str = ' ' + str;
          fim = it.x + it.largura;
          // Troca o número da página (e o total) pelos campos automáticos do Word
          const children = [];
          str.split(/(\d+)/).forEach((pt) => {
            if (!pt) return;
            if (/^\d+$/.test(pt) && !estado.pag && Number(pt) === rep.num && rotuloPag) { children.push(Dx.PageNumber.CURRENT); estado.pag = true; }
            else if (/^\d+$/.test(pt) && estado.pag && !estado.total && Number(pt) === total) { children.push(Dx.PageNumber.TOTAL_PAGES); estado.total = true; }
            else children.push(pt);
          });
          filhos.push(new Dx.TextRun({ ...est, children }));
        }
      });
      return new Dx.Paragraph({ children: filhos, alignment: al, tabStops: tabs.length ? tabs : undefined, spacing: { before: idx ? 0 : 0, after: 0 } });
    });
    const cls = z === 'cab' ? Dx.Header : Dx.Footer;
    return new cls({ children: paras });
  }

  function nivelTitulos(paginas, corpoTam) {
    const tamanhos = new Set();
    for (const pg of paginas) for (const B of pg.blocos) {
      if (B.tipo !== 'texto' || B.linhas.length > 3) continue;
      const t = B.linhas[0].tam;
      if (t >= corpoTam * 1.15) tamanhos.add(t);
    }
    const ordem = [...tamanhos].sort((a, b) => b - a);
    // O maior tamanho geralmente é o título do documento (aparece uma vez): fica fora dos níveis
    const cont = new Map();
    for (const pg of paginas) for (const B of pg.blocos) if (B.tipo === 'texto' && tamanhos.has(B.linhas[0].tam)) cont.set(B.linhas[0].tam, (cont.get(B.linhas[0].tam) || 0) + 1);
    let niveis = ordem;
    if (ordem.length > 1 && cont.get(ordem[0]) === 1) niveis = ordem.slice(1);
    niveis = niveis.slice(0, 3);
    for (const pg of paginas) for (const B of pg.blocos) {
      if (B.tipo !== 'texto' || B.linhas.length > 3) continue;
      const k = niveis.indexOf(B.linhas[0].tam);
      if (k >= 0) B.nivel = k + 1;
    }
  }

  // ---------------------------------------------------------------------
  // Conversão completa
  // ---------------------------------------------------------------------
  async function converter(pdf, { editavel, progresso, previa, titulo }) {
    const Dx = D();
    const paginas = [];
    for (let p = 1; p <= pdf.numPages; p++) {
      progresso(0.85 * (p - 1) / pdf.numPages, `Lendo página ${p} de ${pdf.numPages}…`);
      await pausa();
      const page = await pdf.getPage(p);
      if (!editavel) {
        paginas.push(await paginaImagem(page, p));
      } else {
        const pg = await lerPagina(page, p);
        paginas.push(pg);
      }
      if (p <= 3 && previa) previa(paginas[p - 1].previa, p === 1);
      paginas[p - 1].previa = null;
    }

    progresso(0.9, 'Reconstruindo parágrafos, títulos e tabelas…');
    await pausa();
    const semTexto = editavel && paginas.every((pg) => !pg.temTexto);

    if (semTexto) {
      for (let p = 1; p <= pdf.numPages; p++) paginas[p - 1] = await paginaImagem(await pdf.getPage(p), p);
    }
    if (!editavel || semTexto) {
      return { blob: await Dx.Packer.toBlob(new Dx.Document({ title: titulo, sections: paginas.map(secaoImagem) })), semTexto };
    }

    detectarCabecalhoRodape(paginas);
    const M = calcularMargens(paginas);
    for (const pg of paginas) pg.blocos = montarParagrafos(pg, M);

    // Tamanho e fonte do corpo do texto
    const porTam = new Map(), porFonte = new Map();
    for (const pg of paginas) for (const B of pg.blocos) for (const L of B.linhas) {
      porTam.set(L.tam, (porTam.get(L.tam) || 0) + L.texto.length);
      porFonte.set(L.fonte.familia, (porFonte.get(L.fonte.familia) || 0) + L.texto.length);
    }
    const corpoTam = porTam.size ? [...porTam.entries()].sort((a, b) => b[1] - a[1])[0][0] : 11;
    const corpoFonte = porFonte.size ? [...porFonte.entries()].sort((a, b) => b[1] - a[1])[0][0] : 'Calibri';
    nivelTitulos(paginas, corpoTam);

    const cab = montarHF(paginas, 'cab', M, paginas.length);
    const rod = montarHF(paginas, 'rod', M, paginas.length);
    const primeiraSemCab = paginas.length > 1 && (
      (cab && !paginas[0].linhas.some((L) => L.hf === 'cab')) || (rod && !paginas[0].linhas.some((L) => L.hf === 'rod')));

    // Seções: páginas seguidas com o mesmo tamanho
    const secoes = [];
    let atual = null;
    paginas.forEach((pg, idx) => {
      const filhos = [];
      let ultimaBase = null, ultimoPasso = 0;
      pg.blocos.forEach((B, bi) => {
        const L0 = B.linhas[0];
        const passo = B.tipo === 'tabela' ? 0 : passoDe(B);
        let antes;
        if (ultimaBase === null) antes = L0.topo - M.MT;
        else antes = (L0.base - ultimaBase) - passo;
        B.quebra = idx > 0 && bi === 0 && !(atual && atual.tamanho !== `${pg.W}x${pg.H}`);
        const extras = null;
        if (B.tipo === 'tabela') {
          const tabela = tabelaNativa(B, M, pg.graficos);
          // fim do bloco anterior ≈ linha de base + parte de baixo da linha
          const fimAnterior = ultimaBase === null ? M.MT : ultimaBase + 0.25 * ultimoPasso;
          const vazio = B.topo - fimAnterior;
          if (vazio > 1 || B.quebra) filhos.push(new Dx.Paragraph({ spacing: { before: 0, after: 0, line: Math.round(Math.max(1, vazio) * TW), lineRule: Dx.LineRuleType.EXACT }, pageBreakBefore: B.quebra || undefined, children: [] }));
          filhos.push(tabela);
        } else if (B.tipo === 'sumario' || B.tipo === 'tab') filhos.push(paragrafoTabulado(B, M, antes, extras));
        else filhos.push(paragrafoTexto(B, M, antes, extras));
        const Lu = B.linhas[B.linhas.length - 1];
        // depois de uma tabela, a "linha de base" equivalente fica de modo que o próximo cálculo de espaço dê certo
        ultimaBase = B.tipo === 'tabela' ? B.fundo - 0.25 * corpoTam * 1.2 : Lu.base;
        ultimoPasso = B.tipo === 'tabela' ? corpoTam * 1.2 : passo;
      });
      // Imagens/desenhos da página: ancorados no primeiro parágrafo dela
      if (pg.graficos.length) {
        const imgs = imagensFlutuantes(pg.graficos);
        const idxPar = filhos.findIndex((f) => f instanceof Dx.Paragraph);
        if (idxPar >= 0) filhos[idxPar].addChildElement ? imgs.forEach((im) => filhos[idxPar].addChildElement(im)) : null;
        else filhos.push(new Dx.Paragraph({ children: imgs, pageBreakBefore: idx > 0 || undefined }));
      }
      if (!filhos.length) filhos.push(new Dx.Paragraph({ children: [], pageBreakBefore: idx > 0 || undefined }));

      const tamanho = `${pg.W}x${pg.H}`;
      if (!atual || atual.tamanho !== tamanho) {
        atual = { tamanho, W: pg.W, H: pg.H, filhos: [] };
        secoes.push(atual);
      }
      atual.filhos.push(...filhos);
    });

    const folga = corpoTam * 1.5; // evita que uma diferença mínima de fonte empurre a última linha para outra página
    const doc = new Dx.Document({
      title: titulo,
      styles: {
        default: {
          document: { run: { font: corpoFonte, size: Math.round(corpoTam * 2) }, paragraph: { spacing: { before: 0, after: 0 } } },
          heading1: { run: { font: corpoFonte }, paragraph: { spacing: { before: 0, after: 0 } } },
          heading2: { run: { font: corpoFonte }, paragraph: { spacing: { before: 0, after: 0 } } },
          heading3: { run: { font: corpoFonte }, paragraph: { spacing: { before: 0, after: 0 } } },
        },
      },
      sections: secoes.map((s) => ({
        properties: {
          titlePage: primeiraSemCab || undefined,
          page: {
            size: { width: Math.round(s.W * TW), height: Math.round(s.H * TW), orientation: s.W > s.H ? Dx.PageOrientation.LANDSCAPE : Dx.PageOrientation.PORTRAIT },
            margin: {
              top: Math.round(Math.max(14, M.MT) * TW),
              bottom: Math.round(Math.max(14, s.H - M.MB - folga) * TW),
              left: Math.round(Math.max(0, M.ML) * TW),
              right: Math.round(Math.max(0, s.W - M.MR - 1) * TW),
              header: Math.round(M.cabDist * TW), footer: Math.round(M.rodDist * TW), gutter: 0,
            },
          },
        },
        headers: cab ? { default: cab, first: primeiraSemCab ? new Dx.Header({ children: [new Dx.Paragraph({ children: [] })] }) : undefined } : undefined,
        footers: rod ? { default: rod, first: primeiraSemCab ? new Dx.Footer({ children: [new Dx.Paragraph({ children: [] })] }) : undefined } : undefined,
        children: s.filhos,
      })),
    });
    return { blob: await Dx.Packer.toBlob(doc), semTexto: false };
  }

  // Modo "cópia idêntica": cada página vira uma imagem
  async function paginaImagem(page, num) {
    const vp = page.getViewport({ scale: 1 });
    const W = vp.width, H = vp.height;
    const S = Math.min(200 / 72, 5000 / W, 5000 / H);
    const canvas = document.createElement('canvas');
    canvas.width = Math.round(W * S); canvas.height = Math.round(H * S);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport: page.getViewport({ scale: S }) }).promise;
    const blob = await new Promise((r) => canvas.toBlob(r, 'image/jpeg', 0.92));
    page.cleanup();
    return { num, W, H, previa: canvas, bytes: new Uint8Array(await blob.arrayBuffer()), temTexto: true };
  }

  function secaoImagem(pg) {
    const Dx = D();
    const bytes = pg.bytes;
    return {
      properties: { page: {
        size: { width: Math.round(pg.W * TW), height: Math.round(pg.H * TW), orientation: pg.W > pg.H ? Dx.PageOrientation.LANDSCAPE : Dx.PageOrientation.PORTRAIT },
        margin: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0, gutter: 0 },
      } },
      children: [new Dx.Paragraph({
        spacing: { before: 0, after: 0, line: 20, lineRule: Dx.LineRuleType.EXACT },
        children: [new Dx.ImageRun({
          data: bytes || pg.imagem,
          transformation: { width: pg.W * 96 / 72, height: pg.H * 96 / 72 },
          floating: {
            horizontalPosition: { relative: Dx.HorizontalPositionRelativeFrom.PAGE, offset: 0 },
            verticalPosition: { relative: Dx.VerticalPositionRelativeFrom.PAGE, offset: 0 },
            behindDocument: true, allowOverlap: true, lockAnchor: true,
          },
        })],
      })],
    };
  }

  window.PdfWord = { converter };
})();
