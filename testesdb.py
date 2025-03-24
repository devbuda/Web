import streamlit as st
import pandas as pd
import pyodbc
from io import BytesIO

st.set_page_config(
    page_title="Amazoncopy | Consulta de movimentação",
    page_icon="Amazoncopy Logo menor.png",
    layout="centered"
)

st.title("📑 Consulta de equipamentos")

with st.form("consulta_form", clear_on_submit=False):
    codigo = st.text_input("Código:")
    num_serie = st.text_input("Número de série:")

    submitted = st.form_submit_button("Buscar")

def get_connection():
    return pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};'
        'SERVER=192.168.0.22\\MSSQLSERVER2017;'
        'DATABASE=DADOSADV;'
        'UID=amcopy;'
        'PWD=Admin@123'
    )

def buscar_dados(codigo=None, num_serie=None):
    query = """
SELECT 
  DB_PRODUTO AS 'CÓDIGO', 
  B1_DESC AS 'DESCRIÇÃO', 
  DB_NUMSERI AS 'N° SÉRIE',  
  DB_DOC AS 'NOTA',
  FORMAT( CAST(DB_DATA AS DATE), 'dd/MM/yyyy','en-US') EMISSÃO,
  
  CASE
   WHEN DB_ORIGEM = 'SD1' AND DB_TIPONF = 'N' THEN 'ENTRADA'
   WHEN DB_ORIGEM = 'SD1' AND DB_TIPONF = 'B' THEN 'RETORNO'
   WHEN DB_ORIGEM = 'SD1' AND DB_TIPONF = 'D' THEN 'RETORNO'
   WHEN DB_ORIGEM = 'SC6' THEN 'SAIDA'      
   ELSE 'INTERNA'
    END AS "TIPO MOV",  

  CASE
   WHEN DB_ORIGEM = 'SD1' AND DB_TIPONF = 'N' THEN A2_NOME 
   WHEN DB_ORIGEM = 'SD1' AND DB_TIPONF = 'B' THEN A1_NOME
   WHEN DB_ORIGEM = 'SD1' AND DB_TIPONF = 'D' THEN A1_NOME
   WHEN DB_ORIGEM = 'SC6' THEN A1_NOME
   ELSE 'MV_INTER'
    END AS CLIENTE_FORNECEDOR,
  
  CASE  
   WHEN DB_ESTORNO = 'S' THEN 'CANCELADO' ELSE 'OK' 
    END AS ESTORNO,

  DB_LOCAL AS 'ARMAZÉM', 
  DB_LOCALIZ AS 'ENDEREÇO'
FROM SDB010 DB (NOLOCK)
INNER JOIN SB1010 B1 (NOLOCK) ON DB.DB_PRODUTO = B1.B1_COD AND DB.DB_FILIAL = B1.B1_FILIAL AND B1.D_E_L_E_T_ = ''
LEFT JOIN SA2010 A2 (NOLOCK) ON A2.A2_COD = DB.DB_CLIFOR AND A2.A2_LOJA = DB.DB_LOJA AND A2.D_E_L_E_T_ = ''
LEFT JOIN SA1010 A1 (NOLOCK) ON A1.A1_COD = DB.DB_CLIFOR AND A1.A1_LOJA = DB.DB_LOJA AND A1.D_E_L_E_T_ = ''
WHERE DB.D_E_L_E_T_ = ''
  AND NOT (DB_ORIGEM NOT IN ('SD1', 'SC6') OR (DB_ORIGEM = 'SD1' AND DB_TIPONF NOT IN ('N','B','D')))
    """

    if codigo:
        query += f" AND DB_PRODUTO LIKE '%{codigo}%'"
    if num_serie:
        query += f" AND DB_NUMSERI LIKE '%{num_serie}%'"

    query += " ORDER BY DB_NUMSERI, DB_DATA"

    with get_connection() as conn:
        return pd.read_sql(query, conn)

if submitted:
    if (not codigo or len(codigo) < 5) and (not num_serie or len(num_serie) < 5):
        st.warning("Digite pelo menos 5 caracteres no Código ou Número de Série para realizar a busca.")
    else:
        try:
            df = buscar_dados(codigo, num_serie)
            if df.empty:
                st.info("Nenhum resultado encontrado.")
            else:
                st.success(f"{len(df)} resultado(s) encontrado(s):")
                st.dataframe(df, use_container_width=True)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Movimentações')
                output.seek(0)

                st.download_button(
                    label="Baixar excel",
                    data=output,
                    file_name="movimentacoes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Erro ao buscar dados: {e}")