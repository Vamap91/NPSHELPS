import streamlit as st
import pandas as pd
import io
from openai import OpenAI

st.set_page_config(page_title="Helps - Curadoria de Comentários", layout="wide")

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

def classificar_comentario(descricao, comentario):
    if pd.isna(comentario) or str(comentario).strip() == "":
        return "SEM_COMENTARIO"
    descricao_texto = "" if pd.isna(descricao) else str(descricao)
    comentario_texto = str(comentario)
    prompt = (
        "Você irá atuar como uma camada de curadoria da voz do cliente para a empresa Helps.\n"
        "Receberá a descrição do atendimento e o comentário do cliente.\n\n"
        "Seu objetivo é identificar apenas os comentários que exigem revisão prévia, evitando ruídos desnecessários.\n"
        "Classifique o comentário em UMA das categorias abaixo, pensando na gravidade, tom e relevância atual:\n\n"
        "1) CRITICO: tom muito negativo ou agressivo, reclamação forte, possível risco de imagem ou de relacionamento, problema aparentemente não resolvido.\n"
        "2) ATENCAO: comentário negativo ou de insatisfação, mas em tom mais controlado ou com problema encaminhado/mitigado.\n"
        "3) NEUTRO_POSITIVO: comentário neutro, elogio, sugestão leve, sem crítica relevante.\n"
        "4) IRRELEVANTE_DESCONTEXTUALIZADO: comentário fora de contexto, pouco claro, já resolvido ou que não contribui para a leitura atual da experiência.\n"
        "5) SEM_COMENTARIO: quando não houver texto relevante para avaliar.\n\n"
        "Use também a descrição do atendimento apenas como apoio de contexto.\n\n"
        f"Descrição do atendimento: {descricao_texto}\n"
        f"Comentário do cliente: {comentario_texto}\n\n"
        "Responda apenas com UMA das opções, exatamente como escrito:\n"
        "CRITICO, ATENCAO, NEUTRO_POSITIVO, IRRELEVANTE_DESCONTEXTUALIZADO ou SEM_COMENTARIO."
    )
    response = client.responses.create(
        model="gpt-4o-mini",
        instructions="Você é um curador de feedback de clientes, focado em identificar quais comentários precisam de análise prévia antes de serem encaminhados para a Helps.",
        input=prompt,
        max_output_tokens=10
    )
    classificacao = response.output_text.strip().upper()
    opcoes_validas = {
        "CRITICO",
        "ATENCAO",
        "NEUTRO_POSITIVO",
        "IRRELEVANTE_DESCONTEXTUALIZADO",
        "SEM_COMENTARIO"
    }
    if classificacao not in opcoes_validas:
        return "ATENCAO"
    return classificacao

st.title("Helps - Curadoria de Comentários (Classificação de Sentimento)")

st.write(
    "Faça upload da base NPS e gere uma nova coluna com a classificação dos comentários, "
    "para facilitar o filtro dos casos mais críticos antes de enviar para a Helps."
)

arquivo = st.file_uploader("Envie o arquivo Excel (.xlsx)", type=["xlsx", "xls"])

if arquivo is not None:
    try:
        df = pd.read_excel(arquivo, header=1)
    except Exception:
        arquivo.seek(0)
        df = pd.read_excel(arquivo)

    st.subheader("Pré-visualização da planilha")
    st.dataframe(df.head())

    colunas = list(df.columns)

    col_descricao = st.selectbox(
        "Coluna de descrição/contexto (opcional, ajuda no entendimento do comentário)",
        ["(nenhuma)"] + colunas,
        index=(colunas.index("Motivo_escolha_Nota") + 1) if "Motivo_escolha_Nota" in colunas else 0
    )

    col_comentario = st.selectbox(
        "Coluna de comentários do cliente (texto a ser classificado)",
        colunas,
        index=colunas.index("Comentario") if "Comentario" in colunas else 0
    )

    nome_col_classificacao = st.text_input(
        "Nome da nova coluna de classificação",
        value="classificacao_comentario"
    )

    if st.button("Gerar coluna de classificação"):
        if nome_col_classificacao in df.columns:
            st.error("O nome da nova coluna já existe na planilha. Escolha outro nome.")
            st.stop()

        progresso = st.progress(0)
        status_text = st.empty()
        classificacoes = []
        total_linhas = len(df)

        for i, linha in df.iterrows():
            descricao_val = None if col_descricao == "(nenhuma)" else linha[col_descricao]
            comentario_val = linha[col_comentario]
            status_text.text(f"Processando linha {i + 1} de {total_linhas}...")
            try:
                classificacao = classificar_comentario(descricao_val, comentario_val)
            except Exception:
                classificacao = "ATENCAO"
            classificacoes.append(classificacao)
            progresso.progress(int(((i + 1) / total_linhas) * 100))

        df_saida = df.copy()
        df_saida[nome_col_classificacao] = classificacoes

        st.success("Coluna de classificação gerada com sucesso.")
        st.subheader("Pré-visualização da planilha com classificação")
        st.dataframe(df_saida.head())

        buffer = io.BytesIO()
        df_saida.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            label="Baixar planilha com curadoria (.xlsx)",
            data=buffer,
            file_name="base_helps_curadoria.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
