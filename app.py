import openpyxl #le, edita e escreve em arquivos excel
import pyautogui #automação de mouse e teclado

tabela = openpyxl.load_workbook('clientes_produtos.xlsx') #Openpyxl le o nosso arquivo, que como está na mesma pasta só precisei citala, se não colocaria seu caminho. E na sequencia coloquei numa variavel chamada tabela
pagina_de_vendas = tabela['pagina de vendas'] #Especifique qual a pagina do arquivo excel vou trabalhar, e guardei a informação na variavel pagina_de_vendas

for linha in pagina_de_vendas.iter_rows(min_row=2): #Aqui crio um laço especificando que iremos analisar a partir da linha 2, com o min_row=2
# ["Cliente1"	"Bola de Futebol"	"5"	"Esportes"] --> Basicamente, em cada linha temos itens dentro de uma lista, lista linha...
    #Então especificamos nossos 4 itens, lendo do 0 ao 3:
    #(linha[0].value)
    #(linha[1].value)
    #(linha[2].value)
    #(linha[3].value)
    
    ###Cliente
    pyautogui.click(249,195,duration=0.1) #Usamos o pyautogui para clicar na cordenada que especificamos em  249,195 e o tempo que isso levará: duration=1
    pyautogui.write(str(linha[0].value))#Usamos o pyautogui para escrever usando o item da nossa linha que especificamos: 0
    ###Produto
    pyautogui.click(247,223,duration=0.1)
    pyautogui.write(str(linha[1].value))
    ###Quantidade
    pyautogui.click(246,252,duration=0.1)
    pyautogui.write(str(linha[2].value)) #Como o pyautogui não consegue escrever numeros, convertemos nosso numero em string. Fazemos, isso em todos para evitar erros
    ###Categoria
    pyautogui.click(248,288,duration=0.1)
    pyautogui.write(str(linha[3].value))
    ###Salvar
    pyautogui.click(267,324,duration=0.1)
    ###OK
    pyautogui.click(888,505,duration=0.1)
    
    
















