import os
import pandas as pd
import tabula
#import pytesseract
#import pdf2image
import PyPDF2
import re
import chardet
import openpyxl
from simple_chalk import chalk
from pprint import pprint
from unidecode import unidecode

empresa_name = "CRESCER"

arr_textos = []

if empresa_name == "ARTE BRILHO":
    coluna = "Líquido"
    coluna_nome = "Nome Conta"
elif empresa_name == "CRESCER":
    coluna = "VALOR"
    coluna_nome = "NOME CPF"
"""
empresa_name = "CRESCER"
coluna = "VALOR"
coluna_nome = "NOME CPF"
"""
script_dir = os.path.dirname(os.path.abspath(__file__))
docs_dir = os.path.join(script_dir,"..", "docs",empresa_name)

#arr_textos.append("Extraindo dados do Comprovante de pagamento")
#log_builder(arr_textos=arr_textos)

# Define o nome do arquivo PDF
pdf_filename = "COMPROVANTE_DE_PAGAMENTO.pdf"

# Constrói o caminho completo para o arquivo PDF
pdf_path = os.path.join(docs_dir, pdf_filename)
options = {"pages": "all", "output_format": "csv"}

# Define o nome do arquivo CSV de saída
csv_filename = "output"

dfs = []

def detect_csv_encoding(file_path):
    with open(file_path, 'rb') as file:
        raw_data = file.read()
        result = chardet.detect(raw_data)
        #print(f"A codificação do CSV do comprovante de pagamento é: {result['encoding']}")
        return result['encoding']

if empresa_name == "CRESCER":

    with open(pdf_path, "rb") as arquivo_pdf:
        leitor_pdf = PyPDF2.PdfReader(arquivo_pdf)
        num_paginas = len(leitor_pdf.pages)

    for i in range(2):
        if i == 0:
            # Constrói o caminho completo para o arquivo CSV incial
            csv_path = os.path.join(docs_dir, csv_filename + str(i) + "ori.csv")
            tabula.convert_into(pdf_path, csv_path, output_format="csv", pages="1", area=[322.36,40,720,570])
            # Pega qual a codificação do csv original e passa para utf-8
            input_encoding = detect_csv_encoding(csv_path)
            df = pd.read_csv(csv_path, encoding=input_encoding)
            new_csv_path = os.path.join(docs_dir, csv_filename + str(i) + ".csv")
            df.to_csv(new_csv_path, encoding='utf-8', index=False)
            dfs.append(df)
        else:
            csv_path = os.path.join(docs_dir, csv_filename + str(i) + "ori.csv")
            tabula.convert_into(pdf_path, csv_path, output_format="csv", pages=f'2-{num_paginas}', area=[20,40,720,570])
            # Pega qual a codificação do csv original e passa para utf-8
            input_encoding = detect_csv_encoding(csv_path)
            df = pd.read_csv(csv_path, encoding=input_encoding)
            new_csv_path = os.path.join(docs_dir, csv_filename + str(i) + ".csv")
            df.to_csv(new_csv_path, encoding='utf-8', index=False)
            dfs.append(df)
    # Lê o arquivo CSV e cria uma tabela Excel
    df = pd.concat(dfs, ignore_index=True)
elif empresa_name == "ARTE BRILHO":
    csv_path = os.path.join(docs_dir, csv_filename + "0.csv")
    # Converte o PDF para um arquivo CSV
    tabula.convert_into(pdf_path, csv_path, output_format="csv", pages="all")
    df = pd.read_csv(csv_path, encoding='utf-8')

# Verifica se há valores vazios na coluna "Líquido"
empty_rows_serie = df[coluna].isna()
valor_rows_serie = df[coluna].isin([coluna])

empty_rows = []

for i in range(len(empty_rows_serie)):
    bol = empty_rows_serie[i]
    if bol == True:
        empty_rows.append(i)

for i in range(len(valor_rows_serie)):
    bol = valor_rows_serie[i]
    if bol and i > 0:
        empty_rows.append(i)

df = df.drop(empty_rows)

# Seleciona somente as colunas de nome e valor (coluna)
df_comp_pag = df[[coluna_nome, coluna]]

df_comp_pag.loc[:,coluna_nome] = df_comp_pag[coluna_nome].apply(lambda x: x.upper())

#arr_textos.append("Extração de dados do comprovante de pagamento concluída com sucesso!")
#log_builder(arr_textos=arr_textos)

# *********************************************************************************************

""" 
script FGTS
"""

#arr_textos.append("Extraindo dados do FGTS (GFIP / SEFIP)")
#log_builder(arr_textos=arr_textos)

if empresa_name == "ARTE BRILHO":

    # Adicione o caminho para o executável do Poppler no sistema PATH
    os.environ['PATH'] += os.pathsep + r'C:\Program Files (x86)\poppler-0.68.0\bin'

    # Caminho para o arquivo PDF
    pdf_path = r'G:\Meu Drive\Dados DIEP\10. GGCAT\Terceirizados Automação\docs\FGTS.pdf'

    # Use pdf2image para converter as páginas do PDF em imagens PIL
    images = pdf2image.convert_from_path(pdf_path)

    # Loop pelas imagens e extrai o texto com o pytesseract
    texts = []
    for image in images:
        text = pytesseract.image_to_string(image, lang='por').split('\n')
        for quebra in text:
            texts.append(quebra)
        # texts.append(text)

    # Cria um dataframe com o texto extraído
    df = pd.DataFrame({'Text': texts})

elif empresa_name == "CRESCER":

    script_dir = os.path.dirname(os.path.abspath(__file__))
    docs_dir = os.path.join(script_dir,"..", "docs",empresa_name)

    # Define o nome do arquivo PDF
    pdf_filename = "FGTS.pdf"

    # Constrói o caminho completo para o arquivo PDF
    pdf_path = os.path.join(docs_dir, pdf_filename)
    options = {"pages": "all", "output_format": "csv"}

    # Define o nome do arquivo CSV de saída
    csv_filename = "output"

    # Abrir o arquivo PDF em modo de leitura binária
    with open(pdf_path, "rb") as arquivo_pdf:
        leitor_pdf = PyPDF2.PdfReader(arquivo_pdf)
        num_paginas = len(leitor_pdf.pages)

    def handle_bad_line(bad_line):
        
        #print("bad", bad_line)
        
        new_line = []
        
        for i, campo in enumerate(bad_line):
            if i!=1:
                new_line.append(campo)
    
        #print("new", new_line)
        # Retorne a linha corrigida
        return new_line

    # Cria o arquivo output0.csv, passa os dados do pdf para ele e cria um DataFrame baseado nesse csv
    csv_path = os.path.join(docs_dir, csv_filename + "0.csv")
    tabula.convert_into(pdf_path, csv_path, output_format="csv", pages=f'1-{num_paginas-2}', area=[230,34,530.5,774.35])
    df_fgts = pd.read_csv(csv_path,encoding='utf-8', on_bad_lines=handle_bad_line, engine='python', header=None)
    """
    print(chalk.yellow("DF FGTS"))
    print(df_fgts)
    """
    # Cria um arquivo Excel a partir do DataFrame

#arr_textos.append("Extração de dados do FGTS (GFIP / SEFIP) concluída com sucesso!")
#log_builder(arr_textos=arr_textos)

# *************************************************************************************

#arr_textos.append("Extraindo dados da folha analítica")
#log_builder(arr_textos=arr_textos)

if empresa_name == "CRESCER":

    def caminhoArquivo(pasta_nome,arquivo_nome):
        arquivo = os.path.join(pasta_nome,arquivo_nome)
        return arquivo

    # Obtém os diretórios necessários
    script_dir = os.path.dirname(os.path.abspath(__file__))
    docs_dir = os.path.join(script_dir,"..", "docs",empresa_name)

    # Define o nome do arquivo PDF
    pdf_filename = "FOLHA ANALITICA.pdf"

    pdf_file = open(caminhoArquivo(docs_dir,pdf_filename), 'rb')
    pdf_reader = PyPDF2.PdfReader(pdf_file)

    nomes = []
    salarios = []
    cargos = []
    data_adm = []

    # Loop
    for page in pdf_reader.pages:
        page_text = page.extract_text()
        lines = page_text.split('\n')
        
        nomes_enviados = 0
        salarios_enviados = 0
        cargos_enviados = 0
        datas_enviados = 0

        n_profs = page_text.count('Dep IR')

        #print(f'Profissionais na página: {n_profs}')

        for i, line in enumerate(lines):
            if (salarios_enviados % n_profs != 0 or salarios_enviados == 0) or (cargos_enviados % n_profs != 0 or cargos_enviados == 0):
                if line == "05118764000108 CNPJ/CEI: " or "Base IRRF Folha" in line:
                    nome = lines[i+1]
                    # Expressão regular para encontrar a cadeia de caracteres desejada
                    padrao = r'[A-Z\s]+'
                    # Extrair a cadeia de caracteres
                    resultado = re.search(padrao, nome).group().strip()
                    if len(resultado) > 0:
                        nomes_enviados += 1
                        nomes.append(resultado)
                    
                    esta_data_adm_ln = lines[i+2]
                    esta_data_adm = esta_data_adm_ln[0:10]
                    if not re.search(padrao, esta_data_adm):
                        datas_enviados += 1
                        data_adm.append(esta_data_adm)
                elif line == '*************** ':
                    salario = lines[i+1]
                    # Expressão regular para encontrar o padrão "X.XXX,XX"
                    padrao = r'\d\.\d{3},\d{2}'
                    # Extrair a cadeia de caracteres
                    #resultado = re.findall(padrao, salario)[-1]
                    resultado = salario.split()[-1]
                    salarios_enviados += 1
                    salarios.append(resultado)
                elif 'Dep IR : Dep SF : ' in line:
                    # String de exemplo
                    string = line
                    # Separar a string utilizando o sinal de dois pontos como delimitador
                    split_string = string.split(':')
                    # Recuperar a última cadeia de caracteres
                    cargo = split_string[-1].strip()
                    cargos_enviados += 1
                    cargos.append(cargo)

    df_fol_an = pd.DataFrame({'Nome': nomes, 'Salário Liquido': salarios,'Cargo': cargos, 'Data_Admissão': data_adm})
    """ 
    print(chalk.yellow("DF FOLHA ANALÍTICA"))
    print(df_fol_an)
    """
elif empresa_name == "ARTE BRILHO":
    """
    Script planilha Folha Análitica
    """
    import PyPDF2
    import pandas as pd
    import os

    def caminhoArquivo(pasta_nome,arquivo_nome):
        arquivo = os.path.join(pasta_nome,arquivo_nome)
        return arquivo

    # Obtém os diretórios necessários
    script_dir = os.path.dirname(os.path.abspath(__file__))
    docs_dir = os.path.join(script_dir,"..", "docs",empresa_name)
    resultado_dir = os.path.join(script_dir,"..", "resultado")

    # Define o nome do arquivo PDF
    pdf_filename = "FOLHA ANALITICA.pdf"

    pdf_file = open(caminhoArquivo(docs_dir,pdf_filename), 'rb')
    pdf_reader = PyPDF2.PdfReader(pdf_file)

    nomes = []
    salarios = []
    cargos = []
    data_adm = []

    # Loop
    for page in pdf_reader.pages:
        page_text = page.extract_text()
        lines = page_text.split('\n')

        for i, line in enumerate(lines):
            if line == "Colaborador:":
                nomes.append(lines[i+1])
            elif line == "Líquido:":
                salarios.append(lines[i+1])
            elif line == "C.Custo:":
                cargos.append(lines[i+1])
            elif line == "Dep. IR:":
                data_adm.append(lines[i+1])

    # Define o nome do xlsx de saída
    xlsx_filename = "folha_analitica.xlsx"

    df = pd.DataFrame({'Nome': nomes, 'Salário Liquido': salarios,'Cargo': cargos, 'Data_Admissão': data_adm})
    df.to_excel(caminhoArquivo(docs_dir,xlsx_filename), index=False)

#arr_textos.append("Extração de dados da folha analítica concluída com sucesso")
#log_builder(arr_textos=arr_textos)

# ******************************************************************************************************

plan_conf_sal = os.path.join(docs_dir,"Conf_salarios.xlsx")

nomes = []
cpfs = []
pis_paseps = []
datas_adm = []
cargos = []
salarios = []


aba_mdo = openpyxl.load_workbook(plan_conf_sal)['MAO DE OBRA ']
for i, line in enumerate(aba_mdo['D']):
    if i >= 13:
        if line.value == None:
            break
        nomes.append(unidecode(line.value))
        data = aba_mdo[f'E{i+1}']
        datas_adm.append(data.value)
        cargo = aba_mdo[f'G{i+1}']
        cargos.append(cargo.value)
        salario = aba_mdo[f'T{i+1}']
        salarios.append(salario.value)
        cpf = aba_mdo[f'B{i+1}']
        cpfs.append(cpf.value)
        pis_pasep = aba_mdo[f'C{i+1}']
        pis_paseps.append(pis_pasep.value)

df_conf_sal = pd.DataFrame({
    'Nome':nomes, 
    'Data de admissão':datas_adm, 
    'Cargo':cargos, 
    'Salário do mês':salarios,
    'CPF': cpfs,
    'PIS/PASEP': pis_paseps
    })

# Exclui todos os arquivos que tem no nome o texto output
arquivos = os.listdir(docs_dir)
for arquivo in arquivos:
    if "output" in arquivo:
        os.remove(os.path.join(docs_dir, arquivo))

# printar todos os df's gerados


print(chalk.bold("df_conf_sal"))
pprint(df_conf_sal)
print('-' * 120)

print(chalk.bold("df_comp_pag"))
pprint(df_comp_pag)
print('-' * 120)

print(chalk.bold("df_fgts"))
pprint(df_fgts)
print('-' * 120)

print(chalk.bold("df_fol_an"))
pprint(df_fol_an)


# Comparações entre os DataFrames

lista_comparacao = [
    {'df':df_comp_pag,'titulo':"Comprovante de pagamento",'coluna_nome':"NOME CPF",'coluna_valor':'VALOR'},
    {'df':df_fgts,'titulo':"FGTS",'coluna_nome':0},
    {'df':df_fol_an,'titulo':"Folha analítica",'coluna_nome':"Nome",'coluna_valor':'Salário Liquido'}
]


df_relatorio_erros = pd.DataFrame(columns=['conferencia','nome','erro'])

doc = lista_comparacao[2]

for linha in df_conf_sal.values:
    conf_sal_nome = linha[0]
    conf_sal_data_adm = linha[1]
    conf_sal_cargo = linha[2]
    conf_sal_sal = linha[3]

    nome_compos = conf_sal_nome.split()
    n_nomes = len(nome_compos)

    for i in range(len(nome_compos), 1, -1):
        if i > 1:
            nome_procurado = " ".join(nome_compos[:i])
            resultado = doc['df'].loc[doc['df'][doc['coluna_nome']].str.contains(nome_procurado),doc['coluna_nome']]
            if not resultado.empty:
                break
            #else:
                #print(f'Nome {nome_procurado} não encontrado.')
        else:
            break

    if resultado.empty:
        df_erro = pd.DataFrame([['Planilha -> Folha analítica',conf_sal_nome,'Nome não encontrado']],columns=['conferencia','nome','erro'])
        df_relatorio_erros = pd.concat([df_relatorio_erros,df_erro],ignore_index=True)
    else:
        if len(resultado) == 1:
            cargo = doc['df'].loc[doc['df'][doc['coluna_nome']].str.contains(nome_procurado),'Cargo'].values[0]
            data_adm = doc['df'].loc[doc['df'][doc['coluna_nome']].str.contains(nome_procurado),'Data_Admissão'].values[0]
            
            conf_sal_data_adm_dia = "0" + str(conf_sal_data_adm.day) if conf_sal_data_adm.day < 10 else conf_sal_data_adm.day
            conf_sal_data_adm_mes = "0" + str(conf_sal_data_adm.month) if conf_sal_data_adm.month < 10 else conf_sal_data_adm.month

            conf_sal_data_adm = f'{conf_sal_data_adm_dia}/{conf_sal_data_adm_mes}/{conf_sal_data_adm.year}'

            if(conf_sal_cargo != cargo):
                df_erro = pd.DataFrame([['Planilha -> Folha analítica',conf_sal_nome,'Erro com o cargo']],columns=['conferencia','nome','erro'])
                df_relatorio_erros = pd.concat([df_relatorio_erros,df_erro],ignore_index=True)
            if(conf_sal_data_adm != data_adm):
                df_erro = pd.DataFrame([['Planilha -> Folha analítica',conf_sal_nome,'Erro com a data de admissão']],columns=['conferencia','nome','erro'])
                df_relatorio_erros = pd.concat([df_relatorio_erros,df_erro],ignore_index=True)
        else:
            df_erro = pd.DataFrame([['Planilha -> Folha analítica',conf_sal_nome,'Possível homônimo']],columns=['conferencia','nome','erro'])
            df_relatorio_erros = pd.concat([df_relatorio_erros,df_erro],ignore_index=True)

doc = lista_comparacao[0]

for i, nome in enumerate(df_fol_an['Nome']):
    
    nome_compos = nome.split()
    n_nomes = len(nome_compos)

    for i_nomes in range(len(nome_compos), 1, -1):
        if i_nomes > 1:
            nome_procurado = " ".join(nome_compos[:i_nomes])
            resultado = doc['df'].loc[doc['df'][doc['coluna_nome']].str.contains(nome_procurado),doc['coluna_nome']]
            if not resultado.empty:
                break
            #else:
                #print(f'Nome {nome_procurado} não encontrado.')
        else:
            break

    if resultado.empty: 
        df_erro = pd.DataFrame([['Folha analítica -> Comprovante de pagamento',nome,'Nome não encontrado']],columns=['conferencia','nome','erro'])
        df_relatorio_erros = pd.concat([df_relatorio_erros,df_erro],ignore_index=True)
    else:
        if len(resultado) == 1:
            valor = doc['df'].loc[doc['df'][doc['coluna_nome']].str.contains(nome_procurado),doc['coluna_valor']].values[0]

            if isinstance(valor,str):
                if valor[-5:] == " 0,00":
                    valor = valor[:-5]
            valor = valor.replace("R$","")
            valor = valor.replace(".","")
            valor = valor.replace(",",".")
            salario_documento = round(float(valor),2)

            salario_fol_an = df_fol_an['Salário Liquido'][i]
            if isinstance(valor,str):
                if salario_fol_an[-5:] == " 0,00":
                    salario_fol_an = salario_fol_an[:-5]
            salario_fol_an = salario_fol_an.replace("R$","")
            salario_fol_an = salario_fol_an.replace(".","")
            salario_fol_an = salario_fol_an.replace(",",".")
            salario_fol_an = round(float(salario_fol_an),2)

            if salario_fol_an != salario_documento:
                df_erro = pd.DataFrame([['Folha analítica -> Comprovante de pagamento',nome,'Erro com o salário']],columns=['conferencia','nome','erro'])
                df_relatorio_erros = pd.concat([df_relatorio_erros,df_erro],ignore_index=True)
        else:
            df_erro = pd.DataFrame([['Folha analítica -> Comprovante de pagamento',nome,'Possível homônimo']],columns=['conferencia','nome','erro'])
            df_relatorio_erros = pd.concat([df_relatorio_erros,df_erro],ignore_index=True)

doc = lista_comparacao[1]

for linha in df_conf_sal.values:
    
    conf_sal_nome = linha[0]
    conf_sal_data_adm = linha[1]
    conf_sal_cargo = linha[2]
    conf_sal_sal = linha[3]
    
    nome_compos = conf_sal_nome.split()
    n_nomes = len(nome_compos)

    for i in range(len(nome_compos), 1, -1):
        if i > 1:
            nome_procurado = " ".join(nome_compos[:i])
            resultado = doc['df'].loc[doc['df'][doc['coluna_nome']].str.contains(nome_procurado),doc['coluna_nome']]
            if not resultado.empty:
                break
            #else:
                #print(f'Nome {nome_procurado} não encontrado.')
        else:
            break

    if resultado.empty: 
        df_erro = pd.DataFrame([['Planilha -> GFIP - SEFIP',conf_sal_nome,'Nome não encontrado']],columns=['conferencia','nome','erro'])
        df_relatorio_erros = pd.concat([df_relatorio_erros,df_erro],ignore_index=True)
    else:
        if len(resultado) == 1:
            data_adm = doc['df'].loc[doc['df'][doc['coluna_nome']].str.contains(conf_sal_nome),2].values[0]

            conf_sal_data_adm_dia = "0" + str(conf_sal_data_adm.day) if conf_sal_data_adm.day < 10 else conf_sal_data_adm.day
            conf_sal_data_adm_mes = "0" + str(conf_sal_data_adm.month) if conf_sal_data_adm.month < 10 else conf_sal_data_adm.month

            conf_sal_data_adm = f'{conf_sal_data_adm_dia}/{conf_sal_data_adm_mes}/{conf_sal_data_adm.year}'

            if(conf_sal_data_adm != data_adm):
                df_erro = pd.DataFrame([['Planilha -> GFIP - SEFIP',conf_sal_nome,'Erro com a data de admissão']],columns=['conferencia','nome','erro'])
                df_relatorio_erros = pd.concat([df_relatorio_erros,df_erro],ignore_index=True)
        else:
            df_erro = pd.DataFrame([['Folha analítica -> Comprovante de pagamento',nome,'Possível homônimo']],columns=['conferencia','nome','erro'])
            df_relatorio_erros = pd.concat([df_relatorio_erros,df_erro],ignore_index=True)

xlsx_filename = "Relatório.xlsx"
df_relatorio_erros.to_excel(caminhoArquivo(docs_dir,xlsx_filename), index=False)
