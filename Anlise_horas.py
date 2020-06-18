import pandas as pd
import datetime

data = pd.read_csv("Enviados.csv", ";")  # Dados extraidos da caixa de correio
df_results = pd.DataFrame()  # Cria a tabela df_results

df_results._reindex_columns = ["Data", "Semana", "Trabalhadas", "Extra"]
# cria colunas da tabela results a ser construida com os dados iniciais
data["Data"] = data.Data.str.split(".", expand=True)
data[["Data", "Hora"]] = data.Data.str.split(" ", expand=True)  # Separa a data da hora
data["Data"] = pd.to_datetime(data["Data"], format="%Y-%m-%d")
# Converte string em data
data["Hora"] = pd.to_timedelta(data["Hora"])
# converte string em hora

inicio_dia = datetime.timedelta(hours=8)
filt_hora = data["Hora"] >= inicio_dia
data = data[filt_hora]
print(data)
# elimina todas as horas entre as 24 e 8 da manhã

df_results["Data"] = data["Data"].drop_duplicates()
# Cria coluna na df_results de cada dia que existe na "data"


df_results["Semana"] = df_results["Data"].dt.dayofweek
# cria coluna com o dia da semana correspondete à data

df_results = df_results.reset_index(drop=True)


inicial = []
final = []
numero_mails = []
for date in df_results["Data"]:
    hora = data.loc[data["Data"] == date, "Hora"]
    inicial.append(hora.min())
    final.append(hora.max())
    numero_mails.append(len(hora))
# gera os valores para as colunas

jornada = datetime.timedelta(hours=9)  # 8 horas de trabalho +1 de Almoço
jornada_fds = datetime.timedelta(hours=4)

df_results.insert(2, column="Hora_Final", value=final, allow_duplicates=False)
# Insere as colunes com a hora do ultimo e-mail do dia
df_results.insert(2, column="Hora_Inicial", value=inicial, allow_duplicates=False)
# Insere as colunes com a hora do primeiro e-mail do dia
df_results["Trabalhadas"] = df_results["Hora_Final"] - df_results["Hora_Inicial"]
# insere a coluna que faz a diferença entre a hora do primeiro email do dia e a hora do ultimo email.
df_results.insert(5, column="Num_Mails", value=numero_mails, allow_duplicates=False)
# Insere a coluna com o numero de e-maisl enviados no dia

df_results.to_excel("Teste.xlsx")
filt_sab = df_results["Semana"] == 5
filt_dom = df_results["Semana"] == 6

df_results_fds = df_results.loc[filt_sab | filt_dom]
# df_results_fds = df_results_fds.reset_index

df_results = df_results.loc[~filt_sab | ~filt_dom]

filt = df_results["Trabalhadas"] > jornada
df_results = df_results.loc[filt]
# elimina todos os dias com menos de 8 horas
df_results["Extra"] = df_results["Trabalhadas"] - jornada
# Contabiliza as horas para além das 8 normais do dia


filt_fds = df_results_fds["Trabalhadas"] > jornada_fds
df_results_fds = df_results_fds[filt_fds]
# elimina todos os dias com menos de 4 horas
df_results_fds["Extra"] = df_results_fds["Trabalhadas"] - jornada_fds
# Contabiliza as horas para além das 4 normais da meia folga

df_final = df_results.append(df_results_fds)
print(df_results)
print(df_results_fds)
df_final.to_excel("Resultados.xlsx")
