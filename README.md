# Conversor PDF ⇄ Word

Página web que converte **PDF em Word (.docx)** e **Word (.docx) em PDF**, direto no navegador.
Os arquivos **não são enviados para nenhum servidor**: tudo acontece no seu computador.

**Usar online:** https://pchiesi.github.io/Pdf-word/

## Como usar

1. Abra a página e escolha **PDF para Word** ou **Word para PDF**.
2. Arraste o arquivo para a área indicada (ou clique para escolher).
3. O arquivo convertido é baixado automaticamente.

### PDF → Word
- **Editável, com o mesmo visual** (padrão): cada linha de texto vira texto editável, na mesma posição, fonte, tamanho e cor. Imagens, tabelas, linhas e fundos coloridos ficam como fundo da página, então o visual é o mesmo do PDF.
- **Cópia idêntica**: cada página vira uma imagem de alta resolução no Word. Fica 100% igual, mas não dá para editar o texto. Use para PDFs escaneados.

### Word → PDF
- **PDF com texto selecionável** (padrão): abre a janela de impressão; escolha *Salvar como PDF*.
- **Baixar PDF direto**: baixa na hora, com as páginas como imagem.

## Limites
- Arquivos `.doc` antigos precisam ser salvos como `.docx` antes.
- Para o Word → PDF ficar idêntico, o computador precisa ter as fontes usadas no documento (no Windows com Office, já tem).
- No PDF → Word editável, o texto é posicionado linha a linha (caixas de texto), não em parágrafos corridos. É o que garante o mesmo visual.

## Publicação (GitHub Pages)
Settings → Pages → *Deploy from a branch* → escolha o branch com estes arquivos e a pasta `/ (root)`.

## Bibliotecas usadas (na pasta `vendor/`)
pdf.js, docx, docx-preview, JSZip, html2canvas e jsPDF — todas de código aberto.
