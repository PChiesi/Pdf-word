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
  // 1) Leitura de uma página: texto, cores, desenhos, tabelas e figuras
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
      if (base.x > W || base.x + largura < 0 || base.fundo < 0 || base.topo > H) continue; // fora da página
      itens.push(...dividirLider(base));
    }

    const orig = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const pad = 1.5;
    for (const it of itens) {
      const a = analisarCaixa(orig.data, canvas.width, canvas.height, it.x * S, it.topo * S - pad, (it.x + it.largura) * S, it.fundo * S + pad);
      it.cor = a ? hex(a.cor) : '000000';
      it.fundoCor = a ? a.fundo : [255, 255, 255];
    }

    // Desenhos vetoriais e imagens, lidos dos comandos do PDF
    const ops = await page.getOperatorList();
    const G = analisarOperadores(ops, vp.transform, W, H);
    marcarSublinhados(G, itens);
    const tabelas = detectarGrades(G, itens);
    const { figuras, enfeites } = agruparGraficos(G, itens, tabelas, W, H);

    // Distribui o texto: figura (fica na imagem), tabela (vai para a célula) ou corpo
    const dentro = (it, b, tol = 1.5) => {
      const cx = it.x + it.largura / 2, cy = (it.topo + it.fundo) / 2;
      return cx >= b.x0 - tol && cx <= b.x1 + tol && cy >= b.y0 - tol && cy <= b.y1 + tol;
    };
    const corpo = [];
    for (const it of itens) {
      const fig = figuras.find((f) => dentro(it, f, 2));
      if (fig) { it.naFigura = true; continue; }
      const tab = tabelas.find((t) => dentro(it, t));
      if (tab) {
        const cel = tab.celulas.find((c) => dentro(it, { x0: tab.x[c.c], x1: tab.x[c.c + c.cs], y0: tab.y[c.r], y1: tab.y[c.r + c.rs] }, 0.5));
        if (cel) { cel.itens.push(it); continue; }
      }
      corpo.push(it);
    }

    // Cópia sem o texto (para os enfeites e o fundo das células)
    const limpa = document.createElement('canvas');
    limpa.width = canvas.width; limpa.height = canvas.height;
    const lc = limpa.getContext('2d', { willReadFrequently: true });
    lc.drawImage(canvas, 0, 0);
    for (const it of itens) {
      if (it.naFigura) continue;
      lc.fillStyle = `rgb(${it.fundoCor.map(Math.round).join(',')})`;
      lc.fillRect(it.x * S - pad, it.topo * S - pad, it.largura * S + 2 * pad, (it.fundo - it.topo) * S + 2 * pad);
    }
    const limpos = lc.getImageData(0, 0, limpa.width, limpa.height);

    for (const T of tabelas) prepararGrade(T, limpos, orig, S, canvas.width, canvas.height);
    for (const f of figuras) f.bytes = await recortar(canvas, orig, S, f, false);
    for (const e of enfeites) e.bytes = await recortar(limpa, limpos, S, e, true);

    const linhas = agruparLinhas(corpo);
    page.cleanup();
    return {
      num, W, H, linhas, previa: canvas, temTexto: itens.length > 0,
      tabelas: tabelas.filter((t) => t.celulas.some((c) => c.itens.length) || t.celulas.length > 1),
      figuras: figuras.filter((f) => f.bytes), enfeites: enfeites.filter((e) => e.bytes),
    };
  }

  // Lê os comandos de desenho: linhas retas, retângulos, curvas e imagens
  function analisarOperadores(lista, vpT, W, H) {
    const OPS = pdfjsLib.OPS;
    const fn = lista.fnArray, args = lista.argsArray;
    let ctm = vpT.slice(), lw = 1;
    const pilha = [];
    const G = { hs: [], vs: [], rets: [], desenhos: [], imagens: [], fundoPagina: false };
    let cam = null;
    const T = (x, y) => [ctm[0] * x + ctm[2] * y + ctm[4], ctm[1] * x + ctm[3] * y + ctm[5]];
    const novoCam = () => ({ rets: [], segs: [], curva: false, pts: [], ini: null, atual: null });
    const caixa = (pts) => {
      let x0 = 1e9, y0 = 1e9, x1 = -1e9, y1 = -1e9;
      for (const [x, y] of pts) { x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y); }
      return { x0, y0, x1, y1 };
    };
    const naPagina = (b) => b.x1 > 0 && b.y1 > 0 && b.x0 < W && b.y0 < H;
    const linhaH = (y, x0, x1, e) => G.hs.push({ y, x0: Math.min(x0, x1), x1: Math.max(x0, x1), e });
    const linhaV = (x, y0, y1, e) => G.vs.push({ x, y0: Math.min(y0, y1), y1: Math.max(y0, y1), e });

    function pintar(preenche, contorna) {
      if (!cam) return;
      const escala = Math.sqrt(Math.abs(ctm[0] * ctm[3] - ctm[1] * ctm[2])) || 1;
      const e = Math.max(0.3, lw * escala);
      for (const r of cam.rets) {
        if (!naPagina(r)) continue;
        const w = r.x1 - r.x0, h = r.y1 - r.y0;
        if (preenche) {
          if (h <= 3 && w > 3 * h) linhaH((r.y0 + r.y1) / 2, r.x0, r.x1, Math.max(h, 0.3));
          else if (w <= 3 && h > 3 * w) linhaV((r.x0 + r.x1) / 2, r.y0, r.y1, Math.max(w, 0.3));
          else if (w > 0.5 && h > 0.5) {
            if (w * h > W * H * 0.85) G.fundoPagina = true;
            else G.rets.push({ ...r });
          }
        }
        if (contorna) {
          if (h < 0.5) linhaH(r.y0, r.x0, r.x1, e);
          else if (w < 0.5) linhaV(r.x0, r.y0, r.y1, e);
          else { linhaH(r.y0, r.x0, r.x1, e); linhaH(r.y1, r.x0, r.x1, e); linhaV(r.x0, r.y0, r.y1, e); linhaV(r.x1, r.y0, r.y1, e); }
        }
      }
      if (cam.segs.length || cam.curva) {
        const todos = cam.pts;
        const b = caixa(todos);
        if (naPagina(b)) {
          const retos = cam.segs.every((s) => Math.abs(s[1] - s[3]) < 0.8 || Math.abs(s[0] - s[2]) < 0.8);
          if (cam.curva || !retos) G.desenhos.push(b);
          else if (contorna) {
            for (const s of cam.segs) {
              if (Math.abs(s[1] - s[3]) < 0.8 && Math.abs(s[0] - s[2]) >= 0.8) linhaH((s[1] + s[3]) / 2, s[0], s[2], e);
              else if (Math.abs(s[0] - s[2]) < 0.8 && Math.abs(s[1] - s[3]) >= 0.8) linhaV((s[0] + s[2]) / 2, s[1], s[3], e);
            }
          }
          if (!cam.curva && retos && preenche) {
            const w = b.x1 - b.x0, h = b.y1 - b.y0;
            if (h <= 3 && w > 3 * h) linhaH((b.y0 + b.y1) / 2, b.x0, b.x1, Math.max(h, 0.3));
            else if (w <= 3 && h > 3 * w) linhaV((b.x0 + b.x1) / 2, b.y0, b.y1, Math.max(w, 0.3));
            else if (w > 0.5 && h > 0.5) { if (w * h > W * H * 0.85) G.fundoPagina = true; else G.rets.push(b); }
          }
        }
      }
      cam = null;
    }

    for (let i = 0; i < fn.length; i++) {
      const f = fn[i], a = args[i];
      switch (f) {
        case OPS.save: pilha.push([ctm, lw]); break;
        case OPS.restore: if (pilha.length) [ctm, lw] = pilha.pop(); break;
        case OPS.transform: ctm = pdfjsLib.Util.transform(ctm, a); break;
        case OPS.setLineWidth: lw = a[0]; break;
        case OPS.paintFormXObjectBegin: pilha.push([ctm, lw]); if (a && a[0]) ctm = pdfjsLib.Util.transform(ctm, a[0]); break;
        case OPS.paintFormXObjectEnd: if (pilha.length) [ctm, lw] = pilha.pop(); break;
        case OPS.constructPath: {
          if (!cam) cam = novoCam();
          const ops = a[0], c = a[1];
          let j = 0;
          for (let k = 0; k < ops.length; k++) {
            switch (ops[k] | 0) {
              case OPS.rectangle: {
                const x = c[j++], y = c[j++], w = c[j++], h = c[j++];
                const p = [T(x, y), T(x + w, y), T(x + w, y + h), T(x, y + h)];
                cam.pts.push(...p);
                const reto = Math.abs(ctm[1]) < 1e-6 && Math.abs(ctm[2]) < 1e-6 || Math.abs(ctm[0]) < 1e-6 && Math.abs(ctm[3]) < 1e-6;
                if (reto) cam.rets.push(caixa(p));
                else { cam.segs.push([...p[0], ...p[1]], [...p[1], ...p[2]], [...p[2], ...p[3]], [...p[3], ...p[0]]); }
                break;
              }
              case OPS.moveTo: { const p = T(c[j++], c[j++]); cam.ini = p; cam.atual = p; cam.pts.push(p); break; }
              case OPS.lineTo: {
                const p = T(c[j++], c[j++]);
                if (cam.atual) cam.segs.push([...cam.atual, ...p]);
                cam.atual = p; cam.pts.push(p); break;
              }
              case OPS.curveTo: { cam.pts.push(T(c[j], c[j + 1]), T(c[j + 2], c[j + 3])); const p = T(c[j + 4], c[j + 5]); j += 6; cam.curva = true; cam.atual = p; cam.pts.push(p); break; }
              case OPS.curveTo2: case OPS.curveTo3: { cam.pts.push(T(c[j], c[j + 1])); const p = T(c[j + 2], c[j + 3]); j += 4; cam.curva = true; cam.atual = p; cam.pts.push(p); break; }
              case OPS.closePath: if (cam.atual && cam.ini) cam.segs.push([...cam.atual, ...cam.ini]); cam.atual = cam.ini; break;
              default: break;
            }
          }
          break;
        }
        case OPS.fill: case OPS.eoFill: pintar(true, false); break;
        case OPS.stroke: case OPS.closeStroke: pintar(false, true); break;
        case OPS.fillStroke: case OPS.eoFillStroke: case OPS.closeFillStroke: case OPS.closeEOFillStroke: pintar(true, true); break;
        case OPS.endPath: cam = null; break;
        case OPS.paintImageXObject: case OPS.paintInlineImageXObject: case OPS.paintImageMaskXObject: case OPS.paintSolidColorImageMask: {
          const b = caixa([T(0, 0), T(1, 0), T(1, 1), T(0, 1)]);
          if (naPagina(b) && (b.x1 - b.x0) > 1 && (b.y1 - b.y0) > 1) {
            if ((b.x1 - b.x0) * (b.y1 - b.y0) > W * H * 0.85) G.fundoPagina = true;
            G.imagens.push({ x0: Math.max(0, b.x0), y0: Math.max(0, b.y0), x1: Math.min(W, b.x1), y1: Math.min(H, b.y1) });
          }
          break;
        }
        default: break;
      }
    }
    G.hs = juntarLinhas(G.hs, 'y', 'x0', 'x1');
    G.vs = juntarLinhas(G.vs, 'x', 'y0', 'y1');
    return G;
  }

  // Junta pedaços de linha alinhados e encostados (bordas desenhadas célula a célula)
  function juntarLinhas(ls, eixo, a, b) {
    ls = ls.filter((l) => l[b] - l[a] > 1).sort((p, q) => p[eixo] - q[eixo] || p[a] - q[a]);
    const out = [];
    for (const l of ls) {
      const m = out.find((o) => Math.abs(o[eixo] - l[eixo]) <= 1.2 && l[a] <= o[b] + 1.5 && l[b] >= o[a] - 1.5);
      if (m) { m[a] = Math.min(m[a], l[a]); m[b] = Math.max(m[b], l[b]); m.e = Math.max(m.e, l.e); }
      else out.push({ ...l });
    }
    return out;
  }

  // Linha fina logo abaixo de um texto = sublinhado (vira formatação, não desenho)
  function marcarSublinhados(G, itens) {
    G.hs = G.hs.filter((h) => {
      if (h.e > 1.6) return true;
      const alvo = itens.filter((it) => h.y >= it.base - 0.5 && h.y <= it.base + 0.3 * it.tam
        && it.x >= h.x0 - 1.5 && it.x + it.largura <= h.x1 + 1.5);
      if (!alvo.length) return true;
      const cob = alvo.reduce((s, it) => s + it.largura, 0);
      if (cob < (h.x1 - h.x0) * 0.6) return true;
      alvo.forEach((it) => { it.sublinhado = true; });
      return false;
    });
  }

  const agrupar = (vals, tol) => {
    const s = [...vals].sort((a, b) => a - b), out = [];
    for (const v of s) { const u = out[out.length - 1]; if (u && v - u.max <= tol) { u.max = v; u.soma += v; u.n++; } else out.push({ max: v, soma: v, n: 1 }); }
    return out.map((g) => g.soma / g.n);
  };

  // Tabelas com linhas de grade
  function detectarGrades(G, itens) {
    const hs = G.hs, vs = G.vs;
    const n = hs.length + vs.length;
    if (hs.length < 2 || vs.length < 2) return [];
    const pai = [...Array(n).keys()];
    const raiz = (i) => (pai[i] === i ? i : (pai[i] = raiz(pai[i])));
    for (let i = 0; i < hs.length; i++) for (let j = 0; j < vs.length; j++) {
      const h = hs[i], v = vs[j];
      if (v.x >= h.x0 - 2.5 && v.x <= h.x1 + 2.5 && h.y >= v.y0 - 2.5 && h.y <= v.y1 + 2.5) pai[raiz(i)] = raiz(hs.length + j);
    }
    const grupos = new Map();
    for (let i = 0; i < n; i++) { const r = raiz(i); if (!grupos.has(r)) grupos.set(r, { hs: [], vs: [] }); (i < hs.length ? grupos.get(r).hs : grupos.get(r).vs).push(i < hs.length ? hs[i] : vs[i - hs.length]); }
    const tabelas = [];
    for (const g of grupos.values()) {
      if (g.hs.length < 2 || g.vs.length < 2) continue;
      const x0 = Math.min(...g.hs.map((h) => h.x0)), x1 = Math.max(...g.hs.map((h) => h.x1));
      const y0 = Math.min(...g.vs.map((v) => v.y0)), y1 = Math.max(...g.vs.map((v) => v.y1));
      const xs = agrupar([x0, x1, ...g.vs.map((v) => v.x)], 3);
      const ys = agrupar([y0, y1, ...g.hs.map((h) => h.y)], 3);
      if (xs.length < 2 || ys.length < 2) continue;
      const cobre = (lista, eixo, pos, a, b, ini, fim) => {
        let tot = 0;
        for (const l of lista) if (Math.abs(l[eixo] - pos) <= 2.5) tot += Math.max(0, Math.min(fim, l[b]) - Math.max(ini, l[a]));
        return tot >= (fim - ini) * 0.6;
      };
      const temV = (x, ya, yb) => cobre(g.vs, 'x', x, 'y0', 'y1', ya, yb);
      const temH = (y, xa, xb) => cobre(g.hs, 'y', y, 'x0', 'x1', xa, xb);
      const nr = ys.length - 1, nc = xs.length - 1;
      const dono = Array.from({ length: nr }, () => Array(nc).fill(null));
      const celulas = [];
      for (let r = 0; r < nr; r++) {
        for (let c = 0; c < nc; c++) {
          if (dono[r][c]) continue;
          let cs = 1;
          while (c + cs < nc && !dono[r][c + cs] && !temV(xs[c + cs], ys[r], ys[r + 1])) cs++;
          let rs = 1;
          while (r + rs < nr && !temH(ys[r + rs], xs[c], xs[c + cs]) && dono[r + rs].slice(c, c + cs).every((d) => !d)) rs++;
          const cel = { r, c, cs, rs, itens: [] };
          for (let i = r; i < r + rs; i++) for (let k = c; k < c + cs; k++) dono[i][k] = cel;
          celulas.push(cel);
        }
      }
      if (celulas.length < 2) continue; // um retângulo só é moldura, não tabela
      const bordaDe = (lista, eixo, pos, a, b, ini, fim) => {
        const l = lista.find((x) => Math.abs(x[eixo] - pos) <= 2.5 && Math.min(fim, x[b]) - Math.max(ini, x[a]) >= (fim - ini) * 0.6);
        return l ? { e: l.e, pos: l[eixo], meio: (Math.max(ini, l[a]) + Math.min(fim, l[b])) / 2 } : null;
      };
      for (const c of celulas) {
        const xa = xs[c.c], xb = xs[c.c + c.cs], ya = ys[c.r], yb = ys[c.r + c.rs];
        c.bordas = {
          top: bordaDe(g.hs, 'y', ya, 'x0', 'x1', xa, xb), bottom: bordaDe(g.hs, 'y', yb, 'x0', 'x1', xa, xb),
          left: bordaDe(g.vs, 'x', xa, 'y0', 'y1', ya, yb), right: bordaDe(g.vs, 'x', xb, 'y0', 'y1', ya, yb),
        };
      }
      g.hs.forEach((h) => { h.usada = true; }); g.vs.forEach((v) => { v.usada = true; });
      tabelas.push({ tipo: 'grade', x: xs, y: ys, celulas, x0: xs[0], x1: xs[nc], y0: ys[0], y1: ys[nr], comBorda: true });
    }
    // Retângulos de fundo dentro da tabela viram cor das células
    for (const t of tabelas) for (const r of G.rets) {
      if (r.x0 >= t.x0 - 2 && r.x1 <= t.x1 + 2 && r.y0 >= t.y0 - 2 && r.y1 <= t.y1 + 2) r.usado = true;
    }
    return tabelas;
  }

  // Agrupa os desenhos que sobraram: figuras (imagem, gráfico) ou enfeites (linhas, caixas, faixas)
  function agruparGraficos(G, itens, tabelas, W, H) {
    const els = [];
    for (const b of G.imagens) els.push({ ...b, img: true });
    for (const b of G.desenhos) els.push({ ...b, des: true });
    for (const r of G.rets) if (!r.usado) els.push({ ...r, ret: true });
    for (const h of G.hs) if (!h.usada) els.push({ x0: h.x0, x1: h.x1, y0: h.y - h.e / 2 - 0.5, y1: h.y + h.e / 2 + 0.5, lin: true });
    for (const v of G.vs) if (!v.usada) els.push({ x0: v.x - v.e / 2 - 0.5, x1: v.x + v.e / 2 + 0.5, y0: v.y0, y1: v.y1, lin: true });
    // Imagens/desenhos dentro de tabelas ficam soltos (não juntam com a tabela)
    const perto = (a, b, d) => a.x0 <= b.x1 + d && b.x0 <= a.x1 + d && a.y0 <= b.y1 + d && b.y0 <= a.y1 + d;
    let grupos = els.map((e) => ({ x0: e.x0, y0: e.y0, x1: e.x1, y1: e.y1, img: !!e.img, des: !!e.des, rets: e.ret ? 1 : 0, area: (e.x1 - e.x0) * (e.y1 - e.y0) }));
    let mudou = true;
    while (mudou) {
      mudou = false;
      for (let a = 0; a < grupos.length && !mudou; a++) for (let b = a + 1; b < grupos.length; b++) {
        const A = grupos[a], B = grupos[b];
        // só junta peças de um mesmo desenho: encostadas, ou uma dentro da outra
        if (!perto(A, B, 4)) continue;
        A.x0 = Math.min(A.x0, B.x0); A.y0 = Math.min(A.y0, B.y0); A.x1 = Math.max(A.x1, B.x1); A.y1 = Math.max(A.y1, B.y1);
        A.img = A.img || B.img; A.des = A.des || B.des; A.rets += B.rets; A.area += B.area;
        grupos.splice(b, 1); mudou = true; break;
      }
    }
    const figuras = [], enfeites = [];
    for (const g of grupos) {
      const w = g.x1 - g.x0, h = g.y1 - g.y0;
      if (w < 1.5 && h < 1.5) continue;
      const areaTexto = itens.reduce((s, it) => {
        const ix = Math.max(0, Math.min(g.x1, it.x + it.largura) - Math.max(g.x0, it.x));
        const iy = Math.max(0, Math.min(g.y1, it.fundo) - Math.max(g.y0, it.topo));
        return s + ix * iy;
      }, 0);
      const densidade = areaTexto / Math.max(1, w * h);
      const grande = w * h > W * H * 0.6;
      const ehFigura = !grande && (g.img || g.des || g.rets >= 3) && densidade < 0.3 && w > 6 && h > 6;
      (ehFigura ? figuras : enfeites).push({ x0: g.x0, y0: g.y0, x1: g.x1, y1: g.y1, grafico: !g.img && (g.des || g.rets >= 3) });
    }
    // Rótulos de gráfico colados no desenho (eixos, categorias) entram na figura
    for (let volta = 0; volta < 2; volta++) for (const f of figuras) {
      if (!f.grafico) continue;
      for (const it of itens) {
        const dx = Math.max(f.x0 - (it.x + it.largura), it.x - f.x1, 0);
        const dy = Math.max(f.y0 - it.fundo, it.topo - f.y1, 0);
        if ((dx > 0 || dy > 0) && dx < 0.8 * it.tam && dy < 0.8 * it.tam && it.str.trim().length <= 24 && it.largura < (f.x1 - f.x0) * 0.5) {
          f.x0 = Math.min(f.x0, it.x - 1); f.x1 = Math.max(f.x1, it.x + it.largura + 1);
          f.y0 = Math.min(f.y0, it.topo - 1); f.y1 = Math.max(f.y1, it.fundo + 1);
        }
      }
    }
    return { figuras, enfeites };
  }

  // Fundo e bordas das células (cores lidas da página)
  function prepararGrade(T, limpos, orig, S, cw, ch) {
    const px = (d, x, y) => { const X = Math.max(0, Math.min(cw - 1, Math.round(x * S))), Y = Math.max(0, Math.min(ch - 1, Math.round(y * S))); const i = (Y * cw + X) * 4; return [d.data[i], d.data[i + 1], d.data[i + 2]]; };
    for (const c of T.celulas) {
      const xa = T.x[c.c], xb = T.x[c.c + c.cs], ya = T.y[c.r], yb = T.y[c.r + c.rs];
      const r = [], g = [], b = [];
      for (let k = 1; k <= 5; k++) for (let m = 1; m <= 5; m++) {
        const p = px(limpos, xa + 2 + (xb - xa - 4) * k / 6, ya + 2 + (yb - ya - 4) * m / 6);
        r.push(p[0]); g.push(p[1]); b.push(p[2]);
      }
      const fundo = [mediana(r), mediana(g), mediana(b)];
      c.fundo = quaseBranco(fundo) ? null : hex(fundo);
      for (const lado of ['top', 'bottom', 'left', 'right']) {
        const bd = c.bordas && c.bordas[lado];
        if (!bd) continue;
        // cor: o pixel mais escuro sobre a linha
        let melhor = null;
        for (let d = -1.5; d <= 1.5; d += 0.5) {
          const p = lado === 'top' || lado === 'bottom' ? px(orig, bd.meio, bd.pos + d) : px(orig, bd.pos + d, bd.meio);
          if (!melhor || p[0] + p[1] + p[2] < melhor[0] + melhor[1] + melhor[2]) melhor = p;
        }
        bd.cor = hex(melhor);
      }
    }
  }

  // Recorta uma área da página; descarta se estiver vazia (só branco)
  async function recortar(canvas, dados, S, b, ehEnfeite) {
    const x0 = Math.max(0, Math.floor((b.x0 - 1) * S)), y0 = Math.max(0, Math.floor((b.y0 - 1) * S));
    const x1 = Math.min(canvas.width, Math.ceil((b.x1 + 1) * S)), y1 = Math.min(canvas.height, Math.ceil((b.y1 + 1) * S));
    const w = x1 - x0, h = y1 - y0;
    if (w < 2 || h < 2) return null;
    let tinta = 0, total = 0;
    const passo = Math.max(1, Math.floor(Math.sqrt(w * h / 20000)));
    for (let y = y0; y < y1; y += passo) for (let x = x0; x < x1; x += passo) {
      const i = (y * canvas.width + x) * 4; total++;
      if (765 - dados.data[i] - dados.data[i + 1] - dados.data[i + 2] > 30) tinta++;
    }
    if (!tinta || tinta / total < (ehEnfeite ? 0.002 : 0.001)) return null;
    b.x0 = x0 / S; b.y0 = y0 / S; b.x1 = x1 / S; b.y1 = y1 / S;
    const rc = document.createElement('canvas');
    rc.width = w; rc.height = h;
    rc.getContext('2d').drawImage(canvas, x0, y0, w, h, 0, 0, w, h);
    const tipo = w * h > 400000 && !ehEnfeite ? 'image/jpeg' : 'image/png';
    const blob = await new Promise((r) => rc.toBlob(r, tipo, 0.92));
    return new Uint8Array(await blob.arrayBuffer());
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
    L.tamMax = Math.max(...L.itens.filter((i) => !i.sobrescrito).map((i) => i.tam), L.tam);
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
  function classificar(L, M, celula) {
    const segs = L.segs;
    if (!segs.length) { L.tipo = 'vazia'; return; }
    if (celula && segs.length >= 2) { L.tipo = 'texto'; return; }
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

  // Monta a sequência de blocos (parágrafos, sumário, tabelas, figuras) na ordem da página
  function montarBlocos(linhas, M, extras, celula) {
    const blocos = [];
    const corpo = linhas.filter((L) => !L.hf);
    corpo.forEach((L) => classificar(L, M, celula));
    const seq = [
      ...corpo.map((L) => ({ k: L.base - 0.7 * L.tam, L })),
      ...extras.map((B) => ({ k: B.y0, B })),
    ].sort((a, b) => a.k - b.k);
    let P = null;
    const fecha = () => { if (P) blocos.push(P); P = null; };
    for (let i = 0; i < seq.length; i++) {
      const { L, B } = seq[i];
      if (B) { fecha(); blocos.push(B); continue; }
      if (L.tipo === 'vazia') continue;
      if (L.tipo === 'sumario') { fecha(); blocos.push({ tipo: 'sumario', linhas: [L] }); continue; }
      if (L.tipo === 'tab') {
        fecha();
        // Linhas em colunas seguidas viram tabela (inclui linhas que continuam uma célula)
        const grupo = [L];
        let j = i + 1;
        const xMin = () => Math.min(...grupo.map((g) => g.x0)), xMax = () => Math.max(...grupo.map((g) => g.x1));
        while (j < seq.length && seq[j].L) {
          const C = seq[j].L, ant = grupo[grupo.length - 1];
          const gap = C.base - ant.base;
          if (gap > 2.8 * Math.max(C.tam, ant.tam) || Math.abs(C.tam - ant.tam) > 3) break;
          if (C.tipo === 'tab') { grupo.push(C); j++; continue; }
          if ((C.tipo === 'texto' || C.tipo === 'lista') && C.x0 >= xMin() - 2 && C.x1 <= xMax() + 2 && C.x1 - C.x0 < (xMax() - xMin()) * 0.6
            && j + 1 < seq.length && seq[j + 1].L && seq[j + 1].L.tipo === 'tab' || (gap < 1.5 * C.tam && C.x0 > xMin() + 4 && C.x1 <= xMax() + 2)) { grupo.push(C); j++; continue; }
          break;
        }
        const T = grupo.length >= 2 ? tabelaDeTexto(grupo) : null;
        if (T) { blocos.push(T); i = j - 1; continue; }
        blocos.push({ tipo: 'tab', linhas: [L] });
        continue;
      }
      if (P && podeJuntar(P, L, M)) { P.linhas.push(L); continue; }
      fecha();
      P = { tipo: L.tipo === 'lista' ? 'lista' : 'texto', linhas: [L] };
    }
    fecha();
    return blocos;
  }

  // Tabela sem bordas: colunas = faixas horizontais ocupadas pelo texto
  function tabelaDeTexto(grupo) {
    const segsDe = (L) => (L.segs.length ? L.segs : [{ itens: L.itens, x0: L.x0, x1: L.x1 }]);
    const ints = grupo.filter((L) => L.segs.length >= 2).flatMap((L) => L.segs.map((s) => [s.x0, s.x1])).sort((a, b) => a[0] - b[0]);
    const cols = [];
    for (const [a, b] of ints) { const u = cols[cols.length - 1]; if (u && a <= u[1] + 1) u[1] = Math.max(u[1], b); else cols.push([a, b]); }
    if (cols.length < 2) return null;
    const colDe = (s) => { const cx = (s.x0 + s.x1) / 2; return cols.findIndex((c) => cx >= c[0] - 1 && cx <= c[1] + 1); };
    const rows = [];
    for (const L of grupo) {
      const segs = segsDe(L);
      const idx = segs.map(colDe);
      if (idx.some((k) => k < 0) || new Set(idx).size !== idx.length) return null;
      const ant = rows[rows.length - 1];
      const continua = ant && !idx.includes(0) && L.base - ant.ultBase < 1.5 * L.tam;
      const row = continua ? ant : { cel: cols.map(() => []), topo: L.topo, fundo: L.fundo, ultBase: L.base };
      if (!continua) rows.push(row);
      segs.forEach((s, k) => row.cel[idx[k]].push(...s.itens));
      row.topo = Math.min(row.topo, L.topo); row.fundo = Math.max(row.fundo, L.fundo); row.ultBase = L.base;
    }
    if (rows.length < 2) return null;
    const PAD = 5.4;
    const x = cols.map((c) => c[0] - PAD);
    x.push(cols[cols.length - 1][1] + PAD);
    const y = [rows[0].topo - 2];
    for (let r = 1; r < rows.length; r++) y.push((rows[r - 1].fundo + rows[r].topo) / 2);
    y.push(rows[rows.length - 1].fundo + 2);
    const celulas = [];
    rows.forEach((R, r) => R.cel.forEach((its, c) => celulas.push({
      r, c, cs: 1, rs: 1, itens: its, bordas: {},
      fundo: its.length && !quaseBranco(its[0].fundoCor) ? hex(its[0].fundoCor) : null,
    })));
    return { tipo: 'grade', x, y, celulas, x0: x[0], x1: x[x.length - 1], y0: y[0], y1: y[y.length - 1], comBorda: false, margemFixa: PAD };
  }

  // Gera a tabela do Word (bordas e cores por célula, células mescladas)
  function emitirTabela(T, M, corpoTam) {
    const Dx = D();
    const nr = T.y.length - 1;
    const esq = [], topos = [];
    for (const c of T.celulas) {
      if (!c.itens.length) continue;
      esq.push(Math.min(...c.itens.map((i) => i.x)) - T.x[c.c]);
      topos.push(Math.min(...c.itens.map((i) => i.topo)) - T.y[c.r]);
    }
    const mL = T.margemFixa != null ? T.margemFixa : Math.max(0, Math.min(10, esq.length ? percentil(esq, 0.2) : 5.4));
    const mT = Math.max(0, Math.min(8, topos.length ? percentil(topos, 0.2) - 0.15 * corpoTam : 1));
    const NADA = { style: Dx.BorderStyle.NONE, size: 0, color: 'auto' };
    const borda = (bd) => (bd ? { style: Dx.BorderStyle.SINGLE, size: Math.max(2, Math.min(96, Math.round(bd.e * 8))), color: bd.cor || '000000' } : NADA);

    const celula = (c) => {
      const xa = T.x[c.c], xb = T.x[c.c + c.cs], ya = T.y[c.r], yb = T.y[c.r + c.rs];
      const Mc = { ML: xa + mL, MR: xb - mL };
      const linhas = c.itens.length ? agruparLinhas(c.itens.slice()) : [];
      const blocos = montarBlocos(linhas, Mc, [], true);
      const filhos = [];
      let ultimaBase = null;
      for (const B of blocos) {
        const L0 = B.linhas[0];
        const antes = ultimaBase === null ? 0 : (L0.base - ultimaBase) - passoDe(B);
        filhos.push(paragrafoTexto(B, Mc, antes));
        ultimaBase = B.linhas[B.linhas.length - 1].base;
      }
      if (!filhos.length) filhos.push(new Dx.Paragraph({ children: [] }));
      let va = Dx.VerticalAlign.TOP;
      if (linhas.length) {
        const gTop = Math.min(...linhas.map((l) => l.topo)) - ya, gBot = yb - Math.max(...linhas.map((l) => l.fundo));
        if (gTop - gBot > 4 && gBot < gTop * 0.5) va = Dx.VerticalAlign.BOTTOM;
        else if (Math.abs(gTop - gBot) < Math.max(2, 0.3 * (gTop + gBot)) && gTop > mT + 3) va = Dx.VerticalAlign.CENTER;
      }
      const b = c.bordas || {};
      return new Dx.TableCell({
        columnSpan: c.cs > 1 ? c.cs : undefined,
        rowSpan: c.rs > 1 ? c.rs : undefined,
        width: { size: Math.round((xb - xa) * TW), type: Dx.WidthType.DXA },
        verticalAlign: va,
        shading: c.fundo ? { fill: c.fundo, type: Dx.ShadingType.CLEAR, color: 'auto' } : undefined,
        borders: { top: borda(b.top), bottom: borda(b.bottom), left: borda(b.left), right: borda(b.right) },
        children: filhos,
      });
    };

    const rows = [];
    for (let r = 0; r < nr; r++) {
      const cels = T.celulas.filter((c) => c.r === r).sort((a, b) => a.c - b.c);
      if (!cels.length) continue;
      rows.push(new Dx.TableRow({
        height: { value: Math.round((T.y[r + 1] - T.y[r]) * TW), rule: Dx.HeightRule.ATLEAST },
        children: cels.map(celula),
      }));
    }
    const larguras = [];
    for (let k = 0; k + 1 < T.x.length; k++) larguras.push(Math.round((T.x[k + 1] - T.x[k]) * TW));
    return new Dx.Table({
      rows,
      columnWidths: larguras,
      width: { size: larguras.reduce((a, b) => a + b, 0), type: Dx.WidthType.DXA },
      layout: Dx.TableLayoutType.FIXED,
      indent: { size: Math.round((T.x[0] - M.ML) * TW), type: Dx.WidthType.DXA },
      borders: { top: NADA, bottom: NADA, left: NADA, right: NADA, insideHorizontal: NADA, insideVertical: NADA },
      margins: { left: Math.round(mL * TW), right: Math.round(mL * TW), top: Math.round(mT * TW), bottom: 0 },
    });
  }

  // Figura no meio do texto: parágrafo próprio com a imagem
  function emitirFigura(F, M, antes, quebra) {
    const Dx = D();
    const centro = (M.ML + M.MR) / 2, cx = (F.x0 + F.x1) / 2;
    let alignment = Dx.AlignmentType.LEFT, indent;
    if (Math.abs(cx - centro) < 10) alignment = Dx.AlignmentType.CENTER;
    else if (Math.abs(F.x1 - M.MR) < 6) alignment = Dx.AlignmentType.RIGHT;
    else if (F.x0 - M.ML > 1.5) indent = { left: Math.round((F.x0 - M.ML) * TW) };
    return new Dx.Paragraph({
      alignment, indent, pageBreakBefore: quebra || undefined,
      spacing: { before: Math.round(Math.max(0, antes) * TW), after: 0, line: 240, lineRule: Dx.LineRuleType.AUTO },
      children: [new Dx.ImageRun({ data: F.bytes, transformation: { width: (F.x1 - F.x0) * 96 / 72, height: (F.y1 - F.y0) * 96 / 72 } })],
    });
  }

  // Imagens soltas na página (enfeites atrás do texto; figuras com texto ao lado)
  function imagensFlutuantes(lista) {
    const Dx = D();
    return lista.map((G) => new Dx.ImageRun({
      data: G.bytes,
      transformation: { width: (G.x1 - G.x0) * 96 / 72, height: (G.y1 - G.y0) * 96 / 72 },
      floating: {
        horizontalPosition: { relative: Dx.HorizontalPositionRelativeFrom.PAGE, offset: Math.round(G.x0 * EMU) },
        verticalPosition: { relative: Dx.VerticalPositionRelativeFrom.PAGE, offset: Math.round(G.y0 * EMU) },
        behindDocument: !G.contornar, allowOverlap: true, lockAnchor: true,
        wrap: G.contornar ? { type: Dx.TextWrappingType.SQUARE, side: Dx.TextWrappingSide.BOTH_SIDES } : undefined,
        margins: G.contornar ? { left: 57150, right: 57150 } : undefined,
      },
    }));
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
      underline: it.sublinhado ? {} : undefined,
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
    return (L0.tamMax || L0.tam) * alturaLinha(L0.fonte.familia);
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

  // Logotipos e faixas que se repetem no alto/pé das páginas vão para o cabeçalho/rodapé
  function graficosRepetidos(paginas) {
    const res = { cab: [], rod: [] };
    const N = paginas.length;
    if (N < 2) return res;
    const zonaDe = (pg, g) => (g.y1 < Math.min(pg.H * 0.15, 120) ? 'cab' : g.y0 > pg.H - Math.min(pg.H * 0.15, 120) ? 'rod' : null);
    const mapa = new Map();
    paginas.forEach((pg) => {
      for (const lista of ['figuras', 'enfeites']) for (const g of pg[lista]) {
        const z = zonaDe(pg, g); if (!z) continue;
        const k = z + '|' + [g.x0, g.y0, g.x1, g.y1].map((v) => Math.round(v / 4)).join(',');
        if (!mapa.has(k)) mapa.set(k, []);
        mapa.get(k).push({ pg, lista, g });
      }
    });
    for (const [k, arr] of mapa) {
      if (new Set(arr.map((a) => a.pg)).size < Math.max(2, Math.ceil(N * 0.5))) continue;
      res[k.split('|')[0]].push(arr[0].g);
      for (const a of arr) a.pg[a.lista] = a.pg[a.lista].filter((x) => x !== a.g);
    }
    return res;
  }

  // Cabeçalho/rodapé: monta a partir de uma página representativa
  function montarHF(paginas, z, M, total, imagens) {
    const Dx = D();
    const cls = z === 'cab' ? Dx.Header : Dx.Footer;
    const rep = paginas.find((p, i) => i > 0 && p.linhas.some((L) => L.hf === z)) || paginas.find((p) => p.linhas.some((L) => L.hf === z));
    const flut = imagens && imagens.length ? imagensFlutuantes(imagens) : [];
    if (!rep) return flut.length ? new cls({ children: [new Dx.Paragraph({ children: flut })] }) : null;
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
    if (flut.length && paras.length) flut.forEach((im) => paras[0].addChildElement(im));
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
    const repetidos = graficosRepetidos(paginas);
    for (const pg of paginas) {
      const corpoL = pg.linhas.filter((L) => !L.hf);
      pg.flutuantes = [...pg.enfeites];
      const noTexto = [];
      for (const f of pg.figuras) {
        // texto ao lado da figura → imagem com o texto contornando; senão, imagem no meio do texto
        const lado = corpoL.some((L) => Math.min(L.fundo, f.y1) - Math.max(L.topo, f.y0) > 0.5 * (L.fundo - L.topo) && (L.x1 < f.x0 || L.x0 > f.x1));
        if (lado) { f.contornar = true; pg.flutuantes.push(f); } else noTexto.push({ ...f, tipo: 'figura', linhas: [] });
      }
      pg.tabelas.forEach((t) => { t.linhas = []; });
      pg.blocos = montarBlocos(pg.linhas, M, [...noTexto, ...pg.tabelas], false);
    }

    // Tamanho e fonte do corpo do texto
    const porTam = new Map(), porFonte = new Map();
    for (const pg of paginas) for (const B of pg.blocos) for (const L of (B.linhas || [])) {
      porTam.set(L.tam, (porTam.get(L.tam) || 0) + L.texto.length);
      porFonte.set(L.fonte.familia, (porFonte.get(L.fonte.familia) || 0) + L.texto.length);
    }
    const corpoTam = porTam.size ? [...porTam.entries()].sort((a, b) => b[1] - a[1])[0][0] : 11;
    const corpoFonte = porFonte.size ? [...porFonte.entries()].sort((a, b) => b[1] - a[1])[0][0] : 'Calibri';
    nivelTitulos(paginas, corpoTam);

    const cab = montarHF(paginas, 'cab', M, paginas.length, repetidos.cab);
    const rod = montarHF(paginas, 'rod', M, paginas.length, repetidos.rod);
    const primeiraSemCab = paginas.length > 1 && (
      (cab && !paginas[0].linhas.some((L) => L.hf === 'cab')) || (rod && !paginas[0].linhas.some((L) => L.hf === 'rod')));

    // Seções: páginas seguidas com o mesmo tamanho
    const secoes = [];
    let atual = null;
    paginas.forEach((pg, idx) => {
      const filhos = [];
      let ultimaBase = null, ultimoPasso = corpoTam * 1.2, anteriorTabela = false;
      // Só força quebra se a página anterior do PDF terminou antes do fim (quebra intencional);
      // se estava cheia, deixa o texto correr naturalmente (evita página quase vazia)
      const ant = idx > 0 ? paginas[idx - 1] : null;
      const fimAnt = ant ? Math.max(0, ...ant.blocos.map((b) => (b.linhas && b.linhas.length ? b.linhas[b.linhas.length - 1].fundo : b.y1 || 0))) : 0;
      const forcar = ant && fimAnt < M.MB - 3 * corpoTam * 1.2;
      const continua = ant && !forcar;
      const quebraAqui = (bi) => forcar && bi === 0 && !(atual && atual.tamanho !== `${pg.W}x${pg.H}`);
      pg.blocos.forEach((B, bi) => {
        B.quebra = quebraAqui(bi);
        if (B.tipo === 'grade' || B.tipo === 'figura') {
          const fimAnterior = ultimaBase === null ? M.MT : ultimaBase + 0.25 * ultimoPasso;
          const vazio = ultimaBase === null && continua ? 0 : B.y0 - fimAnterior;
          if (B.tipo === 'grade') {
            if (vazio > 1 || B.quebra || anteriorTabela) {
              filhos.push(new Dx.Paragraph({ spacing: { before: 0, after: 0, line: Math.round(Math.max(1, vazio) * TW), lineRule: Dx.LineRuleType.EXACT }, pageBreakBefore: B.quebra || undefined, children: [] }));
            }
            filhos.push(emitirTabela(B, M, corpoTam));
          } else filhos.push(emitirFigura(B, M, vazio, B.quebra));
          ultimaBase = B.y1 - 0.25 * corpoTam * 1.2;
          ultimoPasso = corpoTam * 1.2;
          anteriorTabela = B.tipo === 'grade';
          return;
        }
        anteriorTabela = false;
        const L0 = B.linhas[0];
        const passo = passoDe(B);
        const antes = ultimaBase === null ? (continua ? 0 : L0.topo - M.MT) : (L0.base - ultimaBase) - passo;
        if (B.tipo === 'sumario' || B.tipo === 'tab') filhos.push(paragrafoTabulado(B, M, antes));
        else filhos.push(paragrafoTexto(B, M, antes));
        ultimaBase = B.linhas[B.linhas.length - 1].base;
        ultimoPasso = passo;
      });
      // Imagens soltas da página: ancoradas no primeiro parágrafo dela
      if (pg.flutuantes.length) {
        const imgs = imagensFlutuantes(pg.flutuantes);
        const idxPar = filhos.findIndex((f) => f instanceof Dx.Paragraph);
        if (idxPar >= 0) imgs.forEach((im) => filhos[idxPar].addChildElement(im));
        else filhos.unshift(new Dx.Paragraph({ children: imgs, pageBreakBefore: idx > 0 || undefined }));
      }
      if (!filhos.length) filhos.push(new Dx.Paragraph({ children: [], pageBreakBefore: idx > 0 || undefined }));

      const tamanho = `${pg.W}x${pg.H}`;
      if (!atual || atual.tamanho !== tamanho) {
        atual = { tamanho, W: pg.W, H: pg.H, filhos: [] };
        secoes.push(atual);
      }
      atual.filhos.push(...filhos);
    });

    const folga = corpoTam * 0.5; // pequena tolerância para diferenças de medida das fontes
    const doc = new Dx.Document({
      title: titulo,
      compatabilityModeVersion: 15,
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
