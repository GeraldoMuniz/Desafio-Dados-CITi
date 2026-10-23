import pandas

#bronze
arquivo = "pedidos_metalurgica_mezzalira.xlsx"
xls = pandas.ExcelFile (arquivo)

bronze = pandas.read_excel (arquivo, sheet_name=xls.sheet_names[0])

#prata

prata = bronze.copy()
prata = prata.dropna (subset=['Pedido']).copy()         #linhas vazias
prata = prata.drop_duplicates().copy()                  #remover duplicados

def padronizarCidade (cidade):
    if pandas.isna (cidade):
        return cidade
    c = str (cidade).strip().title()                    #padronizar cidades
    if c.upper() == 'REC' or c.upper() == 'RECIFE':
        return 'Recife'
    return c

prata ['Cidade'] = prata ['Cidade'].apply (padronizarCidade)

def converterValor (valor):
    if pandas.isna (valor):
        return 0.0
    if isinstance (valor, (int, float)):
        return float (valor)
    modificador = str (valor).replace ("R$", "").strip()                #converter valores para número
    modificador = modificador.replace (".", "").replace (",", ".")
    return float (modificador)

prata ['Valor_Numérico'] = prata ['Valor (R$)'].apply (converterValor)


meses = {'jan': '01', 'fev': '02', 'mar': '03', 'abr': '04', 'mai': '05', 'jun': '06',
         'jul': '07', 'ago': '08', 'set': '09', 'out': '10', 'nov': '11', 'dez': '12'}

def converterData (data):
    if pandas.isna (data):
        return pandas.NaT
    modificador = str (data).strip ()
    if ' 00:00:00' in modificador:
        modificador = modificador.replace (' 00:00:00', '')             #padronizar datas
    partes = modificador.split ("-")
    if len (partes) == 3 and partes [1].lower() in meses:
        dia = partes[0]
        mes = meses [partes[1].lower()]
        ano = partes [2]
        return pandas.to_datetime (f"{ano}-{mes}-{dia}")
    return pandas.to_datetime (modificador, errors = 'coerce')

prata ['Data_Nova'] = prata ['Data'].apply (converterData)
prata ['Mes'] = prata ['Data_Nova'].dt.strftime ('%m-%Y')

#log

log = [
    {
        "O que encontrei": "Linhas em branco",
        "O que fiz": "Exclusão das linhas vazias",
        "Linhas afetadas": "Linhas 14, 33, 51",
        "Justificativa": "Espaços vazios sem dados operacionais"
    },
    {
        "O que encontrei": "Pedidos duplicados",
        "O que fiz": "Remoção dos registros duplicados",
        "Linhas afetadas": "Todas as linhas duplicadas",
        "Justificativa": "Evitar dupla contagem no faturamento"
    },
    {
        "O que encontrei": "Variações na escrita de Recife",
        "O que fiz": "Padronização para 'Recife'",
        "Linhas afetadas": "Todas as linhas que não havia Recife nomeado por 'Recife'",
        "Justificativa": "Unificar escritas diferentes ('recife', 'REC')"
    },
    {
        "O que encontrei": "Valores monetários em texto (ex: R$ ...)",
        "O que fiz": "Conversão para float",
        "Linhas afetadas": "Todas as linhas para float",
        "Justificativa": "Permitir somas e criação de agregações"
    },
    {
        "O que encontrei": "Formatos de data variados",
        "O que fiz": "Padronização para formato de data único",
        "Linhas afetadas": "Linhas com valores em formato de texto",
        "Justificativa": "Permitir agrupamento correto por mês"
    }
]

log = pandas.DataFrame(log)

#ouro

ouro = prata.groupby(['Cidade', 'Mes'])['Valor_Numérico'].sum().reset_index()
ouro = ouro.rename(columns={'Valor_Numérico': 'Faturamento_Total'}) #organizar
ouro = ouro.sort_values(by='Faturamento_Total', ascending=False)

print(ouro.to_string(index=False))

#excel

arquivo_saida = "arquivos_tratados.xlsx"

with pandas.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
    bronze.to_excel(writer, sheet_name='Bronze', index=False)
    prata.to_excel(writer, sheet_name='Prata', index=False)
    log.to_excel(writer, sheet_name='Log', index=False)
    ouro.to_excel(writer, sheet_name='Ouro', index=False)
