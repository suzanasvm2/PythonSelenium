from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import openpyxl

# Função para ler os dados do arquivo Excel
def ler_dados_excel(nome_arquivo):
    wb = openpyxl.load_workbook(nome_arquivo)
    planilha = wb.active
    dados = []
    for linha in planilha.iter_rows(values_only=True):
        dados.append(linha)
    # Ignora a primeira linha que contém os títulos das colunas
    return dados[1:] 

# Caminho para o arquivo Excel com os dados
arquivo_excel = 'data/dados_formulario.xlsx'

# URL do formulário web
url_formulario = 'https://docs.google.com/forms/d/e/1FAIpQLSde0eTjVIV-YU6S0UvGSaWeNpV6PMu4-QB49rlfoGC9Y1FRMg/viewform'

# Ler os dados do arquivo Excel
dados = ler_dados_excel(arquivo_excel)
MAX_DADOS = 3
# Loop sobre os dados do Excel
for dado in dados:
    while(MAX_DADOS!=0):
        # Inicializar o navegador Chrome
        driver = webdriver.Chrome()
        
        # Navegar até a página do formulário
        driver.get(url_formulario)
        time.sleep(2)  # Aguardar 2 segundos para o formulário carregar completamente

        # Encontrar todos os elementos input com a classe especificada
        campos_input = driver.find_elements(By.CSS_SELECTOR, 'input.whsOnd.zHQkBf')

        # Preencher os campos de entrada com os dados fornecidos
        for campo, valor in zip(campos_input, dado):
            campo.send_keys(valor)
            time.sleep(1)  # Aguardar 1 segundo após digitar cada informação

        # Selecionar um item da lista
        item_selecionado = dado[3]  # Considerando que o índice 3 é usado para a coluna com os itens da lista
        itens_lista = driver.find_elements(By.CLASS_NAME, 'eBFwI')
        for item in itens_lista:
            texto_elemento = item.find_element(By.CLASS_NAME, 'aDTYNe').text
            if texto_elemento == item_selecionado:
                item.click()
                break

        # Localizar todos os elementos da lista de restrições alimentares
        restricoes_alimentares = driver.find_elements(By.CLASS_NAME, 'nWQGrd')

        # Iterar sobre os elementos da lista
        restricao_texto_input = dado[4]
        for restricao in restricoes_alimentares:
            # Encontrar o texto do item da lista
            texto_restricao = restricao.find_element(By.CLASS_NAME, 'aDTYNe').text
            if texto_restricao == restricao_texto_input:
                # Clicar no item da lista se for o item desejado
                restricao.click()
                break

        # Localizar o botão de enviar pelo atributo class
        botao_enviar = driver.find_element(By.CLASS_NAME, 'uArJ5e')

        # Clicar no botão de enviar
        botao_enviar.click()
        
        # Fechar o navegador ao final da execução
        driver.quit()
        MAX_DADOS -= 1

