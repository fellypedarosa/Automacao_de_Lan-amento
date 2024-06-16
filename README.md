# Automatização de Entrada de Dados de Vendas

Este projeto consiste em um script Python que automatiza a entrada de dados de vendas em uma aplicação a partir de um arquivo Excel. Ele utiliza as bibliotecas Openpyxl e PyAutoGUI para realizar a automação.

## Funcionalidades

- Lê os dados de vendas de uma planilha Excel.
- Insere os dados em campos específicos de uma aplicação, clicando em coordenadas pré-definidas na tela.

## Bibliotecas Utilizadas

- [Openpyxl] pip install openpyxl: Para trabalhar com arquivos Excel.
- [PyAutoGUI] pip install pyautogui: Para automação de mouse e teclado.

## Funcionamento
1. Usamos o projeto de Cadastro_de_Produtos para aplicar a nossa automação https://github.com/fellypedarosa/Cadastro_de_Produtos
2. Com as coordenadas de clique aplicadas a ferramenta, ela irá escrever os itens da nossa Tabela .xlsx e adicionar nos respectivos espaços
   
![Captura de tela 2024-06-16 003832](https://github.com/fellypedarosa/Automacao_de_Lan-amento/assets/171340743/e7ed8bc0-a72b-4938-a650-98a7a15597d7) ![Captura de tela 2024-06-16 012638](https://github.com/fellypedarosa/Automacao_de_Lan-amento/assets/171340743/3be71306-2a90-4962-8759-8928cbab2736)

3. Com isso, ganhamos muito tempo! Gerando até 1 arquivo por segundo
   
![Captura de tela 2024-06-16 002935](https://github.com/fellypedarosa/Automacao_de_Lan-amento/assets/171340743/b49bdd03-5c6a-4adf-b412-d2d277676676)

![Screenshot_8](https://github.com/fellypedarosa/Automacao_de_Lan-amento/assets/171340743/6397ef18-f0d0-4d2a-934a-27a4fc7dc7d6)

4. ATENÇÂO: Vale lembrar que a nossa biblioteca pyautogui, não escreve acentuação. Então considere uma tabela .xlsx pré formatada sem acentos.

## Como Usar

1. Instale as dependências: openpyxl, pyautogui e mouseinfo
2. Para utilizar o mouse info, instale a biblioteca: pip install mouseinfo
3. Abra o mouseinfo com o python: from mouseinfo import mouseInfo , e inicie o mouse info: mouseInfo()
   
![Captura de tela 2024-06-15 233807](https://github.com/fellypedarosa/Automacao_de_Lan-amento/assets/171340743/92dead8e-78eb-48db-bed8-eb7a8e71f376)

4. Com mouseinfo aberto, ele mostrará a posição em tempo real do seu mouse. E apertando F6 poderá gravar a posição...
5. XY position é a posição: (935,1389)
   
   ![Captura de tela 2024-06-15 233829](https://github.com/fellypedarosa/Automacao_de_Lan-amento/assets/171340743/6a3302fc-0eb1-41ab-807c-7937c85f0abf)

6. Certifique-se de que o arquivo Excel contendo os dados de vendas está no mesmo diretório do script ou forneça o caminho correto, ou se estiver na mesma pasta o seu nome.

7. Execute o script Python.

## Observações

- As coordenadas de clique estão ajustadas conforme necessário para o ambiente específico. Certifique-se de que estão corretas para o seu sistema.
- Ele só executa uma pagina de cada vez, então certifique-se do nome da pagina do arquivo Excel

