# Aulas com Slides

Aplicação local para cadastrar aulas com um PDF de slides e um áudio `.m4a`, ouvir o áudio e marcar em que tempo cada slide começa.

## Rodar

```bash
npm install
npm run dev -- --port 5173
```

Abra `http://localhost:5173/`.

## Rodar como aplicativo

```bash
npm run app
```

## Gerar executável

```bash
npm run dist
```

O arquivo gerado fica em `dist-electron/Aulas com Slides-1.0.0.AppImage`.

## Fluxo

1. Cadastre uma aula informando nome, PDF e áudio.
2. Abra a aula na lista inicial.
3. Use o botão de menu no topo para abrir ou esconder a navegação por slides.
4. Durante o áudio, deixe na tela o slide que está valendo naquele momento e use `Marcar slide atual`.
5. O app salva o início desse slide; o fim é calculado pelo início do próximo slide marcado.
6. Se o professor voltar para um slide já visto, marque esse slide de novo: o app cria uma nova ocorrência na linha do tempo sem apagar a anterior.
7. Ative ou desative `Auto` para ligar ou desligar a troca automática.
8. Ajuste a velocidade no controle entre `0.75x` e `2.00x`.

Os arquivos e marcações ficam salvos no IndexedDB do próprio navegador.

## Editor de comentários

Na aula aberta, clique no botão de lápis no topo para abrir o editor em outra janela.

- Cada slide tem uma página de comentário própria.
- O botão de relógio insere o momento atual do áudio e muda o editor para o slide que está tocando.
- A barra do editor tem estilos de texto, tamanhos, negrito, itálico, lista e lista numerada.
- `DOCX` exporta um arquivo `.docx` separado por slide.
- `PDF` exporta um arquivo `.pdf` separado por slide.
- O editor e as exportações usam padrão acadêmico minimalista: Inter/Aptos/Arial, margens compactas, texto 12 pt, títulos em 3 níveis e blocos de observação com borda esquerda.

## Atalhos

- `Espaço`: marca o slide atual no tempo atual.
- `Seta esquerda`: volta 5 segundos no áudio.
- `Seta direita`: avança 5 segundos no áudio.
- `A`: slide anterior.
- `D`: próximo slide.
