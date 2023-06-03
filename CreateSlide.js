const fs = require("fs");
const pptxgen = require("pptxgenjs");

// Caminho para o arquivo de texto contendo o texto a ser adicionado
const caminho_arquivo_txt = "C:/Users/Rapha/Desktop/CreateSlideAutoJS/Letra.txt";

// Caminho para a imagem de fundo
const caminho_imagem_fundo = "C:/Users/Rapha/Desktop/CreateSlideAutoJS/image_background/fundo.jpg";

// Caminho para a imagem no canto inferior esquerdo
const caminho_imagem_canto = "C:/Users/Rapha/Desktop/CreateSlideAutoJS/image_logo/Logo.png";

// Caminho para o arquivo PowerPoint de saída
const caminho_arquivo_pptx_saida = "C:/Users/Rapha/Desktop/CreateSlideAutoJS/New_Slide.pptx";

// Lê o conteúdo do arquivo de texto
const texto_slide = fs.readFileSync(caminho_arquivo_txt, "utf-8");

// Cria um novo arquivo PowerPoint
const pptx = new pptxgen();

// Cria um novo slide
const slide = pptx.addSlide();

// Adiciona a imagem de fundo ao slide
slide.background = { path: caminho_imagem_fundo };

// Define o estilo do texto
const estiloTexto = {
  color: "ffffff", // Cor do texto (branco)
  fontSize: 54, // Tamanho da fonte em pontos
  bold: true, // Negrito
  align: "center", // Alinhamento centralizado
  valign: "middle", // Centraliza verticalmente o texto
};

// Quebra o texto em linhas
const linhas = texto_slide.split("\n");

// Variável para contar as linhas adicionadas
let contadorLinhas = 0;

// Percorre cada linha do texto
for (let i = 0; i < linhas.length; i++) {
  const linha = linhas[i].trim();

  // Verifica se a linha não está vazia
  if (linha.length > 0) {
    // Adiciona o texto ao slide com o estilo definido
    const textbox = slide.addText(linha, estiloTexto);

    // Define as propriedades de posicionamento do textbox
    textbox.x = "c"; // Posição horizontal centralizada
    textbox.y = "c"; // Posição vertical centralizada

    // Incrementa o contador de linhas
    contadorLinhas++;

    // Verifica se atingiu o limite de 6 linhas
    if (contadorLinhas === 6) {
      // Cria um novo slide
      const novoSlide = pptx.addSlide();
      novoSlide.background = { path: caminho_imagem_fundo };

      // Adiciona a imagem no canto inferior esquerdo no novo slide
      novoSlide.addImage({
        path: caminho_imagem_canto,
        x: 0, // Coordenada X do canto inferior esquerdo
        y: novoSlide.height - 2, // Coordenada Y do canto inferior esquerdo
        w: 2, // Largura da imagem
        h: 2, // Altura da imagem
      });

      // Reinicia o contador de linhas
      contadorLinhas = 0;
    }
  }
}

// Salva o arquivo PowerPoint no caminho especificado
pptx.writeFile(caminho_arquivo_pptx_saida);
