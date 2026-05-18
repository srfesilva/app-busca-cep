import streamlit as st
import pandas as pd
import requests
import time
import io

# Configuração inicial da página do Streamlit
st.set_page_config(page_title="Buscador de CEP - Escolas", page_icon="🏫", layout="centered")

st.title("🏫 Validador e Buscador de CEP")
st.markdown("""
Esta aplicação processa a sua lista de escolas, valida os CEPs existentes e 
busca automaticamente os CEPs em falta utilizando a API do ViaCEP.
""")

# Componente para selecionar o Estado (UF)
uf_selecionada = st.selectbox(
    "Selecione o Estado (UF) das escolas que vai processar:",
    ["SP", "RJ", "MG", "ES", "PR", "SC", "RS", "BA", "PE", "CE", "DF", "GO", "MA", "MT", "MS", "PA", "PB", "PI", "RN", "RO", "AM", "AL", "SE", "TO", "AC", "AP", "RR"],
    index=0
)

# Componente de Upload do Arquivo Excel
uploaded_file = st.file_uploader("👉 Carregue o seu arquivo Excel (.xlsx ou .xls):", type=["xlsx", "xls"])

def limpar_texto(texto):
    if pd.isna(texto):
        return ""
    return str(texto).strip()

def limpar_cep(cep):
    if pd.isna(cep):
        return ""
    return "".join(filter(str.isdigit, str(cep)))

def buscar_cep_por_endereco(uf, cidade, logradouro):
    cidade = limpar_texto(cidade)
    logradouro = limpar_texto(logradouro).split(",")[0].split("-")[0].strip()

    if len(logradouro) < 3 or not cidade:
        return None, "Endereço insuficiente"

    url = f"https://viacep.com.br/ws/{uf}/{cidade}/{logradouro}/json/"

    try:
        response = requests.get(url, timeout=5)
        if response.status_code == 200:
            dados = response.json()
            if isinstance(dados, list) and len(dados) > 0:
                return dados[0]["cep"], "Encontrado por Endereço"
            else:
                return None, "Não encontrado no Busca CEP"
        return None, f"Erro API ({response.status_code})"
    except:
        return None, "Erro de Conexão"

# Se o usuário carregou um arquivo
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ Arquivo carregado com sucesso! Total de registos: {len(df)} linhas.")
        
        # Exibe uma pré-visualização das primeiras linhas
        st.write("📊 Pré-visualização dos dados carregados:")
        st.dataframe(df.head(3))
        
        # Botão para iniciar o processamento
        if st.button("🚀 Iniciar Validação e Busca de CEP"):
            ceps_validados = []
            status_resultados = []

            # Cria elementos visuais de progresso nativos do Streamlit
            barra_progresso = st.progress(0)
            texto_status = st.empty()
            total_linhas = len(df)

            for idx, linha in df.iterrows():
                cep_original = limpar_cep(linha.get("NumeroCEP", ""))

                if len(cep_original) == 8:
                    ceps_validados.append(f"{cep_original[:5]}-{cep_original[5:]}")
                    status_resultados.append("Mantido (CEP Original Válido)")
                else:
                    cidade = linha.get("MunicipioEscola", linha.get("NM_MUNICIPIO", ""))
                    logradouro = linha.get("EnderecoEscola", "")

                    cep_encontrado, status = buscar_cep_por_endereco(uf_selecionada, cidade, logradouro)

                    if cep_encontrado:
                        ceps_validados.append(cep_encontrado)
                        status_resultados.append(status)
                    else:
                        ceps_validados.append(linha.get("NumeroCEP", ""))
                        status_resultados.append(status)

                # Atualiza a barra de progresso do Streamlit
                percentagem = (idx + 1) / total_linhas
                barra_progresso.progress(percentagem)
                texto_status.text(f"A processar linha {idx + 1} de {total_linhas}...")

                # Pausa de segurança para a API
                time.sleep(0.1)

            df["CEP_VALIDADO"] = ceps_validados
            df["STATUS_BUSCA_CEP"] = status_resultados

            st.success("🎉 Processamento concluído!")

            # Transforma o DataFrame atualizado de volta para bytes de Excel em memória
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            dados_excel = output.getvalue()

            # Botão de Download real do Streamlit
            st.download_button(
                label="⬇️ Descarregar Resultado Atualizado (Excel)",
                data=dados_excel,
                file_name="resultado_escolas_cep.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Erro ao ler o arquivo: {e}")
