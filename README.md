# Conversor PDF ⇄ Word

Página web que converte **PDF em Word (.docx)** e **Word (.docx) em PDF**, direto no navegador.
Os arquivos **não são enviados para nenhum servidor**: tudo acontece no seu computador.

**Usar online:** https://pchiesi.github.io/Pdf-word/

## Como usar

1. Abra a página e escolha **PDF para Word** ou **Word para PDF**.
2. Arraste o arquivo para a área indicada (ou clique para escolher).
3. O arquivo convertido é baixado automaticamente.

### PDF → Word
- **Word editável** (padrão): o texto vira parágrafos normais do Word (sem caixas de texto), com a mesma fonte, tamanho, cor, alinhamento, recuos e espaçamentos. O conversor reconhece:
  - **cabeçalho e rodapé**: linhas que se repetem no topo/pé das páginas vão para o cabeçalho/rodapé do Word, com o número da página como campo automático ("Página 1 de 5");
  - **sumário/índice**: título, linha pontilhada e número da página viram tabulação com pontilhado, alinhada à direita;
  - **títulos** (Título 1, 2 e 3 do Word, aparecem no painel de navegação) e **listas** com marcador;
  - **tabelas** nativas do Word, lidas das linhas de grade do PDF: células mescladas, células vazias, cor de fundo e bordas (espessura e cor) de cada célula. Tabelas sem bordas são reconhecidas pelo alinhamento das colunas;
  - **imagens e gráficos** são localizados pelos comandos de desenho do PDF e entram como imagem no meio do texto, na posição certa (com os rótulos do gráfico juntos). Se há texto ao lado, a imagem fica com o texto contornando;
  - logotipos que se repetem no topo/rodapé vão para o cabeçalho/rodapé; textos sublinhados viram sublinhado do Word.
- **Cópia idêntica**: cada página vira uma imagem de alta resolução no Word. Fica 100% igual, mas não dá para editar o texto. Use para PDFs escaneados.

### Word → PDF
- **PDF com texto selecionável** (padrão): abre a janela de impressão; escolha *Salvar como PDF*.
- **Baixar PDF direto**: baixa na hora, com as páginas como imagem.

## Limites
- Arquivos `.doc` antigos precisam ser salvos como `.docx` antes.
- Para o Word → PDF ficar idêntico, o computador precisa ter as fontes usadas no documento (no Windows com Office, já tem).
- PDF → Word: textos em duas colunas, girados ou dentro de desenhos complexos podem não ser reconstruídos perfeitamente (texto girado fica como imagem).

## Publicação (GitHub Pages)
Settings → Pages → *Deploy from a branch* → escolha o branch com estes arquivos e a pasta `/ (root)`.

## Bibliotecas usadas (na pasta `vendor/`)
pdf.js, docx, docx-preview, JSZip, html2canvas e jsPDF — todas de código aberto.
