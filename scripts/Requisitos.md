# Automatização de análise - NUT

## Definições

- VT
  - O extrato de carga mostra qual o valor que foi carregado no vale transporte do cidadão
  - a planilha mostra os cálculos de quanto ele deveria receber (colunas R e V)
    - caso a empresa tenha depositado a mais para o cidadão em relação ao que ele deveria receber, está tudo bem
    - caso a empresa tenha depositado a menos temos que verificar com a empresa qual o motivo de ele ter recebido menos

## Orientações gerais

- Os nomes dos arquivos devem ser:
  - Gerais
    - Relação de admitidos, demitidos e férias: relacao_adm_dem_ferias.xlsx
    - Planilha de conferencia de salários:  Conf_salarios.xlsx
  - Mensal
    - Comprovante de pagamento: COMPROVANTE_DE_PAGAMENTO.pdf
    - Folha analítica: FOLHA ANALITICA.pdf
    - GFIP/SEFIP: FGTS.pdf
  - Vale Transporte
    - Planilha relatorio vt
    - Extrato de carga (VT)

- Todos os arquivos devem ser inseridos na pasta da empresa a ser analisada (docs/empresa)
- As planilhas devem ser salvas sem filtros ou congelamentos e **apenas valores (VERIFICAR)**
- Os formatos devem ser apenas .pdf e .xlsx

## Verificações

### Primeira fase - Entrega oficial em 04/10/2023

- Mensal
  - Planilha -> Folha analítica > Profissional | Cargo | Data de admissão
  - Folha analítica -> Comprovante de pagamento > Valores de salário
  - Planilha -> FGTS (GFIP / SEFIP) > Profissional | Data de admissão

### Segunda fase - EM DESENVOLVIMENTO

- Melhoria da Mensal
  - NOVO: Folha analítica -> Planilha > Profissional
  - NOVO: Comprovante de pagamento -> Folha analítica > Profissional
  - NOVO: FGTS (GFIP / SEFIP) -> Planilha > Profissional
  - NOVO: Planilha -> Comprovante de pagamento > Profissional
  - NOVO: Comprovante de pagamento -> Planilha > Profissional

### Terceira fase

- planilha de custo (matriz de tudo)
  - o que é e como ela se relaciona?

- VT (reunião dia 29/09/2023)
  - Planilha Vt -> Planilha de custo > Custo Liquido (planilha vt) ... Insumo (planilha custo) (DE ACORDO COM O CARGO)
  - Planilha vt -> vt (extrato de carga) > Profissional | Valor de carga
    - vê no mensal se a pessoa teve atestado ou falta e na "relação de admitidos, demitidos e férias" se a pessoa teve férias e conta quantos dias essas intercorrências aconteceram para calcular os dias úteis
    - Valor da carga = dias_úteis \* n_vales_colaborador_dia \* 4,5
    - compara o valor do extrato com da planilha
- vale alimentação
- rescisão
- férias
- paf
- hora extra
- reajuste (atípico)

## Gargalos

- Pedir para as empresas a padronização dos nomes dos documentos e planilhas

## Requisitos do sistema

- Desenvolvido para ambiente Windows
- Ter instalado:
  - JAVA
  - python
    - pandas
    - tabula
    - tabula-py
    - inquirer
    - pytesseract
    - pdf2image
    - PyPDF2
    - re // nativo?
    - chardet
    - openpyxl
    - simple_chalk
    - pprint // nativo?
    - unidecode
    - PyQt5

## Informações de desenvolvimento

- folha analítica = ['Matrícula'] com 6 dígitos
- fgts = pis/pasep apenas números
- comprovante de pagamento = cpf com traco e ponto
- planilha conferencia = cpf com traco e ponto, pis/pasep apenas números, matricula desformatada
