import pandas as pd
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
from selenium.common.exceptions import NoSuchElementException
import matplotlib.pyplot as plt


# 1 Abrir o browser no site do whatsapp

navegador = webdriver.Chrome()
navegador.get('https://web.whatsapp.com/')

# Aguardar até que seja carregado
while len(navegador.find_elements(By.ID, "side")) < 1:
    time.sleep(1)
time.sleep(3)

# 2 Ler a planilha que será usada para envio de mensagens

df = pd.read_excel("Primeira parcela.xlsx")

data = (df['Vencimento'].unique())
# Converter a data para dia, mês e ano
df_date = pd.DataFrame({'Vencimento': data})
df_date['Vencimento'] = pd.to_datetime(df_date['Vencimento'])
df_date['Vencimento_formatado'] = df_date['Vencimento'].dt.strftime('%d-%m-%Y')
df['Vencimento'] = df['Vencimento'].map(
    df_date.set_index('Vencimento')['Vencimento_formatado'])

# substituir linhas vazias por zero
df.fillna(0, inplace=True)

print(df.info())


def alt_telefone(numero):
    if isinstance(numero, str):
        return numero.replace('-', '').replace(' ', '').replace('(', '').replace(')', '')
    return numero


df['Telefone'] = df['Telefone'].apply(alt_telefone)

for linha in df.itertuples():
    nome = linha[1]
    vcto = linha[2]
    pgto = linha[3]
    seguradora = linha[4]
    telefone = linha[5]
    telefone = int(telefone)
    consultor = linha[6]
    print(telefone)
    try:
        # percorrer colunas da planilha e utilizar os dados para o envio das mensagens
        msg = f"Olá {nome}, venho através desta mensagem lembrar que a parcela de seu seguro {seguradora} vencerá em {vcto} através da forma de pagamento: {
            pgto}, caso tenha alguma dúvida entrar em contato com a(o) especialista de seguros que lhe atendeu, {consultor} ."
        link_mensagem_whatsapp = f"https://web.whatsapp.com/send?phone=55{
            telefone}&text={quote(msg)}"
        time.sleep(1)
# abrir o navegador, digitar numero informado e escrever a mensagem
        navegador.get(link_mensagem_whatsapp)

        while len(navegador.find_elements(By.ID, "side")) < 1:
            time.sleep(1)
        time.sleep(14)

        navegador.find_element(
            By.XPATH, '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        time.sleep(3)
    # tratar os erros e exibir em aquivo txt todos os clientes que não enviou mensagem
    except ValueError as e:
        print(f'Erro ao enviar mensagem para {nome}:{e}')
        with open('erros.txt', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'Erro ao enviar mensagem para {nome}\n')

    except Exception as e:
        print(f'Erro ao enviar mensagem para {nome}')
        with open('erros.txt', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'Erro ao enviar mensagem para {nome}\n')
navegador.quit()
# exibir grafico com a forma de pagamento mais utilizada pelos clientes
forma_pgto = df['Forma de pagamento'].value_counts()
print(forma_pgto)

plt.figure(figsize=(8, 6))  # Ajusta o tamanho da figura
plt.pie(forma_pgto, labels=forma_pgto.index, autopct='%1.0f%%', startangle=75)
plt.title('Distribuição de Formas de Pagamento')
plt.axis('equal')
plt.show()
