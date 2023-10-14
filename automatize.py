from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from tqdm import tqdm
import pandas as pd
import time, datetime, pytz
import xlwings as xw

# Especifique o local da planilha e a URL do formulário
planilha_chamados = 'Planilha_de_Chamado.xlsx'  # Local da planilha
planilha_chamados_para_enviar = 'Chamados_para_enviar'  # Planilha dos chamados não enviados.
planilha_modelo = 'Planilha_Modelo' # Planilha modelo para ser copiada e renomeada.
url_do_forms = 'https://forms.office.com/pages/responsepage.aspx?id=8kveREH2aU6_zT_IQS-tZ83OoFNdpiBBiTEr_WrEUX5URUhTVFI3RURLM1RaSEVWTFpVT1Y3MDdRRy4u'

# Mapeamento dos nomes dos meses em inglês para português
meses_em_portugues = {
    'January': 'Janeiro',
    'February': 'Fevereiro',
    'March': 'Março',
    'April': 'Abril',
    'May': 'Maio',
    'June': 'Junho',
    'July': 'Julho',
    'August': 'Agosto',
    'September': 'Setembro',
    'October': 'Outubro',
    'November': 'Novembro',
    'December': 'Dezembro'
}

def formatar_nome_mes(data):
    # Formate o nome do mês em português usando o mapeamento
    nome_mes = meses_em_portugues[data.strftime('%B')]
    return f"{nome_mes}_{data.strftime('%Y')}"

def atualizar_planilha_de_chamados(dataframe):
    # Defina o fuso horário desejado (por exemplo, America/Sao_Paulo para São Paulo, Brasil)
    timezone = pytz.timezone('America/Sao_Paulo')

    # Obtenha a data atual no fuso horário especificado
    data_atual = datetime.datetime.now(timezone)

    # Formate o nome do mês/ano atual usando a função formatar_nome_mes
    nome_mes_ano = formatar_nome_mes(data_atual)

    # Abra o arquivo principal com xlwings
    wb = xw.Book(planilha_chamados)

    proxima_linha = 6  # Inicialize com um valor padrão, assumindo que comece na linha 5
    
    # Verifique se a planilha do mês/ano atual existe, senão, crie uma cópia da planilha modelo e renomeie-a
    if nome_mes_ano not in [sheet.name for sheet in wb.sheets]:
        modelo_sheet = wb.sheets[planilha_modelo]
        planilha_atual = modelo_sheet.copy(name=nome_mes_ano)
        wb.save()
    else:
        # Leia os dados existentes da planilha do mês/ano atual em um DataFrame do pandas
        planilha_atual = wb.sheets[nome_mes_ano]

        # Encontre a próxima célula vazia na coluna A a partir de A5
        coluna_A = planilha_atual.range('A5').expand('down')
        proxima_linha = coluna_A.last_cell.row + 1

        # Verifique se o chamado já existe na planilha para evitar duplicação
        chamados_na_planilha = planilha_atual.range(f'A5:A{proxima_linha - 1}').value
        chamados_existentes = set([chamado[0] for chamado in chamados_na_planilha])
        novos_chamados = [row['CHAMADO'] for _, row in dataframe.iterrows() if row['CHAMADO'] not in chamados_existentes]

        if not novos_chamados:
            return  # Não há novos chamados para adicionar, saia da função

        # Filtra o DataFrame para incluir apenas os novos chamados
        dataframe = dataframe[dataframe['CHAMADO'].isin(novos_chamados)]

    # Adicione a coluna "Cliente" ao DataFrame
    dataframe['CLIENTE'] = dataframe['CHAMADO'].apply(mapear_cliente_por_chamado)

    # Reorganize as colunas para colocar "Cliente" como a primeira coluna
    colunas = dataframe.columns.tolist()
    colunas = ['CLIENTE'] + [coluna for coluna in colunas if coluna != 'CLIENTE']
    dataframe = dataframe[colunas]

    # Adicione os dados mesclados à planilha do mês/ano atual a partir da próxima linha vazia na coluna A
    planilha_atual.range(f'A{proxima_linha}').options(index=False, header=False).value = dataframe

    # Salve as alterações
    wb.save()
    wb.close()


def mapear_cliente_por_chamado(chamado):
    if str(chamado).startswith('55'):
        return 'MAPA'
    elif str(chamado).startswith('22'):
        return 'PCDF'
    elif str(chamado).startswith('96'):
        return 'AGU'
    elif str(chamado).startswith('83'):
        return 'ANA'
    elif str(chamado).startswith('40'):
        return 'ANM'
    elif str(chamado).startswith('78'):
        return 'MINFRA'
    elif str(chamado).startswith('15'):
        return 'MMA'
    elif str(chamado).startswith('81'):
        return 'MMA'

def mapear_observacao(observacao):
    observacoes_mapping = {
        'Chamado Cancelado': 'Chamado Cancelado (Autorizado Pelo Ticket Manager)',
        'Falha no Acesso Remoto': 'Falha no Acesso Remoto (Tentativa Pela Estação e Pelo IP)',
        'Não Existe Procedimento na WIKI': 'Não Existe Procedimento na WIKI (Solicitar Apoio Equipe Avançada Antes de Encaminhar)',
        'Não Foi Possível Contato com o Usuário': 'Não Foi Possível Contato com o Usuário (Realizado Pelo Menos 2 Tentativas de Contato No Prazo de 10 Minutos)',
        'Procedimento Equipe Local': 'Procedimento Equipe Local (WIKI Direciona a Encaminhar Para Equipe Local)',
        'Não Foi Possível Solução na Central (Com Acesso Remoto)': 'Realizado Procedimentos, Não Foi Possível Solução na Central (Com Acesso Remoto)',
        'Usuário Não Conseguiu Seguir os Procedimentos Orientados': 'Usuário Não Conseguiu Seguir os Procedimentos Orientados (Orientações Por Telefone)'
    }
    return observacoes_mapping.get(observacao, observacao)


# Leia a planilha em um DataFrame usando o pandas
dataframe = pd.read_excel(planilha_chamados, sheet_name=planilha_chamados_para_enviar)

# Inicialize o driver do Selenium para o navegador Chrome
navegador = webdriver.Chrome()

# Inicialize uma lista vazia para armazenar os chamados enviados
chamados_enviados = []

try:
    # Loop através das linhas do DataFrame
    for _, row in dataframe.iterrows():
        # Abra a URL do formulário no navegador
        navegador.get(url_do_forms)

        # Use WebDriverWait para esperar até que o elemento de número de chamado seja visível
        elemento_texto_numeroChamado = WebDriverWait(navegador, 100).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="question-list"]/div[1]/div[2]/div/span/input'))
        )

        # Preencha o número de chamado com os dados da planilha
        elemento_texto_numeroChamado.send_keys(row['CHAMADO'])

        # Aguarde até que o elemento de assunto do chamado seja visível
        elemento_texto_assunto_do_chamado = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="question-list"]/div[2]/div[2]/div/span/input'))
        )

        # Preencha o assunto do chamado com os dados da planilha
        elemento_texto_assunto_do_chamado.send_keys(row["ASSUNTO"])

        # Preencha o cliente do chamado com base no número de chamado
        cliente = mapear_cliente_por_chamado(row['CHAMADO'])
        elemento_cliente_chamado = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@aria-label="' + cliente + '"]'))
        )
        elemento_cliente_chamado.click()

        # Mapeie a observação
        observacao_mapeada = mapear_observacao(row['Observações de Resolução'])

        # Aguarde até que o elemento de Observações de Resolução seja visível
        elemento_observacao = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@aria-label="' + observacao_mapeada + '"]'))
        )
        elemento_observacao.click()

        # Converte o valor de APOIO para string
        apoio_str = str(row['APOIO'])

        # Verifica se o valor de APOIO está vazio ou é NaN
        if pd.isna(row['APOIO']) or not apoio_str.strip():
            apoio_str = 'Não Se Aplica'

        # Preencha o campo de apoio
        elemento_apoio = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@aria-label="' + apoio_str + '"]'))
        )
        elemento_apoio.click()

        # Aguarde até que o elemento de Elegível N1 seja visível
        elemento_elegivel_N1 = WebDriverWait(navegador, 20).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="QuestionId_r0672c1f872ce44a19b3dbcb8a56a90d7"]/div/span/span[2]'))
        )

        # Mapeie os valores de "TIPO" para os respectivos XPaths
        xpath_mapping = {
            'Elegível': '//*[@aria-label="Elegível (Existe o Procedimento na Base de Conhecimento)"]',
            'Não Elegível': '//*[@aria-label="Não Elegível  (Na Base de Conhecimento Grupo Resolvedor Outras Equipes N2, N3, NUSOP e ETC)"]',
            'Resolvido sem WIKI': '//*[@aria-label="Resolvido Sem Procedimento na Base de Conhecimento"]'
        }

        # Selecione o elemento com base no valor de "TIPO" usando o mapeamento
        if row['TIPO'] in xpath_mapping:
            elemento_elegivel_N1 = WebDriverWait(navegador, 20).until(
                EC.presence_of_element_located((By.XPATH, xpath_mapping[row['TIPO']]))
            )
            elemento_elegivel_N1.click()
        else:
            print(f"Tipo desconhecido: {row['TIPO']}")

        # Encontre o botão Enviar e clique nele
        elemento_button_enviar = navegador.find_element(By.XPATH, '//*[@data-automation-id="submitButton"]')
        # elemento_button_enviar.click()

        # Registre o chamado enviado
        chamados_enviados.append(f"Chamado {row['CHAMADO']} enviado com sucesso!")
        # Após enviar o chamado para o forms, chame a função para atualizar a planilha
        time.sleep(1)

finally:
    # Certifique-se de que o navegador seja fechado, independentemente de exceções
    navegador.quit()

    #Atualiza a planilha de Chamados
    atualizar_planilha_de_chamados(dataframe)

    # Obtenha a data e hora atual no formato AAAA-MM-DD HH:MM:SS
    data_atual = datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')


    # Crie o nome do arquivo com a data atual
    nome_arquivo = f'historico_chamados_enviados_{data_atual}.txt'

    # Salve o histórico dos chamados enviados em um arquivo .txt
    with open(nome_arquivo, 'w') as arquivo:
        for chamado in chamados_enviados:
            arquivo.write(chamado + '\n')
