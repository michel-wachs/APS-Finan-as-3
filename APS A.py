#Integrantes: 
#Bernardo Silva Ferreira 
#Fernando Camillo 
#Gabriel Erlichman 
#Gustavo Iochua Mejlachowicz 
#Katarina Dantas Camarotti 
#Luísa Monteiro Bragaia 
#Michel Wachslicht 


import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


# Importando os dados do excel para o Python # -------------------------------------------------------------------------
dados = pd.read_excel('Input_APS.xlsx', sheet_name = "Operacoes")
df = pd.DataFrame(dados)

# Criando todos os valores possíveis para St com precisão 0.01 # -------------------------------------------------------
st = np.arange(0, 2 * max(df.Strike), 0.01)  # Fim dos valores possíveis de St vai ser 200% de St

# Fazendo o cálculo para o payoff de cada ativo # ----------------------------------------------------------------------
payoff = np.zeros((len(st), len(df.Strike)))

for i in range(len(df.Strike)):
    for j in range(len(st)):
        if df.Tipo[i] == "Call":
            if st[j] > df.Strike[i]:
                payoff[j][i] = df.Qtde[i] * (st[j] - df.Strike[i])

        if df.Tipo[i] == "Put":
            if st[j] < df.Strike[i]:
                payoff[j][i] = df.Qtde[i] * (df.Strike[i] - st[j])

        if df.Tipo[i] == "Ação":
            payoff[j][i] = df.Qtde[i] * st[j]

        if df.Tipo[i] == "Ativo Rf":
            payoff[j][i] = df.Qtde[i] * df.Strike[i]

# Somando os payoffs dos diferentes ativos para ter o payoff da estratégia # -------------------------------------------
payoff_estrategia = np.zeros(len(st))
for i in range(len(st)):
    payoff_estrategia[i] = sum(payoff[:][i])

# Calculando o fluxo inicial (negativo = investimento, positivo = financiamento) # -------------------------------------
fluxo_inicial = sum(df.Qtde * df.Valor)

# Calculando o retorno da estratégia baseado em seu payoff e seu custo # -----------------------------------------------
retorno_estrategia = np.zeros(len(st))
for j in range(len(st)):
    retorno_estrategia[j] = payoff_estrategia[j] - fluxo_inicial

# Calculando o retorno sobre investimento em % # -----------------------------------------------------------------------
retorno_porcentagem = []
for i in range(len(retorno_estrategia)):
    if fluxo_inicial == 0:
        retorno_porcentagem.append(retorno_estrategia[i])
    else:
        retorno_porcentagem.append((retorno_estrategia[i] / fluxo_inicial)*100)

# Criando um Data Frame com os resultados da estratégia # --------------------------------------------------------------

# Criando as variáveis simplificadas com menos linhas para facilitar a visualização no Data Frame

simplificador = 1000  # variável que vai difinir o tamanho do data frame

st_simplificado = []
payoff_estrategia_simplificado = []
retorno_estrategia_simplificado = []
retorno_porcentagem_simplificado = []


for i in range(int(len(st) / simplificador)):
    # simplificando st
    st_simplificado.append(st[simplificador * i])
    # simplificando payoff
    payoff_estrategia_simplificado.append(payoff_estrategia[simplificador * i])
    # simplificando retorno
    retorno_estrategia_simplificado.append(retorno_estrategia[simplificador * i])
    # simplificando retorno em porcentagem
    retorno_porcentagem_simplificado.append(round(retorno_porcentagem[simplificador * i], 2))


# Criando o Data Frame
df_estrategia = pd.DataFrame({'St': st_simplificado, 'Payoff': payoff_estrategia_simplificado,
                              'Retorno': retorno_estrategia_simplificado,
                              'Retorno (%)': retorno_porcentagem_simplificado})
pd.set_option('display.max_rows', None)
print("#---------------------------------------------------------------------------------------------#")
print("Tabela de valores para a estratégia montada:")
print(df_estrategia)

# Printando informações da estratégia # --------------------------------------------------------------------------------
print("#---------------------------------------------------------------------------------------------#")
print(" -> Você realizou uma estratégia com", len(df.Strike), "operações")
print(' ')
print(" -> O fluxo inicial é de R$", - fluxo_inicial,
      "(obs.: se positivo -> financiamento; se negativo -> investimento)")
print("#---------------------------------------------------------------------------------------------#")

# Plotando o gráfico da estratégia # -----------------------------------------------------------------------------------
linha_zero = np.zeros(len(st))  # criando uma linha no p = 0 para ficar mais fácil vizualizar
plt.plot(st, payoff_estrategia, label="Payoff da Estratégia", color='royalblue', lw=1.8)
plt.plot(st, retorno_estrategia, label="Retorno da Estratégia", color='darkblue', lw=1.8)
plt.plot(st, linha_zero, color='k', lw=1.5,  linestyle='--')
plt.fill_between(st, retorno_estrategia, where=(retorno_estrategia <= 0), color='r', interpolate=True, alpha=0.30)
plt.fill_between(st, retorno_estrategia, where=(retorno_estrategia >= 0), color='lightgreen', interpolate=True, alpha=0.30)
plt.grid()
plt.legend()
plt.title("Gráfico da estratégia montada")
plt.xlabel("St")
plt.ylabel("P")
plt.show()