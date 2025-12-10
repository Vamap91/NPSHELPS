import streamlit as st
import pandas as pd
import io
import json
from openai import OpenAI

st.set_page_config(page_title="Helps - Curadoria de Risco e Sentimento", layout="wide")

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

def analisar_risco_sentimento(descricao, comentario):
    if pd.isna(comentario) or str(comentario).strip() == "":
        return "Baixo", "Sem comentário relevante para análise."
    descricao_texto = "" if pd.isna(descricao) else str(descricao)
    comentario_texto = str(comentario)
    prompt = (
        "Você será um curador da voz do cliente para a empresa Helps.\n"
        "Seu objetivo é analisar o comentário do cliente e retornar:\n"
        "1) Um GRAU DE RISCO entre: Muito Alto, Alto, Médio, Baixo.\n"
        "2) Uma breve explicação do sentimento do pedido, em português do Brasil, de forma objetiva e profissional.\n\n"
        "Definições gerais:\n"
        "- Muito Alto: tom muito agressivo, alto risco de insatisfação, menção a cancelamento, reclamação grave, possível impacto de imagem.\n"
        "- Alto: reclamação clara, problema aparentemente não resolvido ou parcialmente resolvido, insatisfação relevante.\n"
        "- Médio: incômodo moderado, pequena frustração, pontos de melhoria sem grande risco imediato.\n"
        "- Baixo: elogio, comentário neutro, sugestão leve ou feedback positivo.\n\n"
        "Use a descrição do atendimento apenas como contexto complementar.\n\n"
        f"Descrição do atendimento: {descricao_texto}\n"
        f"Comentário do cliente: {comentario_texto}\n\n"
        "Responda apenas com um JSON válido no seguinte formato:\n"
        "{ \"grau_risco\": \"Muito Alto|Alto|Médio|Baixo\", \"explicacao\": \"texto curto explicando o sentimento\" }"
    )
    response = client.responses.create(
        model="gpt-4o-mini",
        instructions="Você é um especialista em experiência do cliente, classificando risco e explicando o sentimento de forma clara e concisa.",
        input=prompt,
        max_output_tokens=200
    )
    texto = response.output_text.strip()
    try:
        dados = json.loads(texto)
        grau = str(dados.get("grau_risco", "")).strip()
        explicacao = str(dados.get("explicacao", "")).strip()
    except Exception:
        grau = ""
        explicacao = ""
    if grau not in ["Muito Alto", "Alto", "Médio", "Baixo"]:
        grau = "Médio"
    if explicacao == "":
        explicacao = "Comentário indica insatisfação moderada, recomendável acompanhamento."
    return grau, explicacao

st.markdown("<h1 style='text-align: center;'>Helps - Curadoria de Risco e Sentimento</h1>", unsafe_allow_html=True)
st.markdown(
    "<p style='text-align: center;'>Faça o upload da base NPS, selecione as colunas e gere duas novas colunas: "
    "<b>Grau de Risco</b> e <b>Explicação do Sentimento</b>, sem alterar o texto original do cliente.</p>",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Passos")
    st.markdown("1. Faça upload do Excel (.xlsx)")
    st.markdown("2. Confirme se os títulos estão corretos (linha 3 da planilha)")
    st.markdown("3. Selecione as colunas de descrição e comentário")
    st.markdown("4. Clique em **Gerar análise** para criar as duas novas colunas")
    st.markdown("---")
    st.markdown("Os dados do cliente não são alterados, apenas enriquecidos com a curadoria.")

arquivo = st.file_uploader("Envie o arquivo Excel (.xlsx)", type=["xlsx", "xls"])

if arquivo is not None:
    try:
        df = pd.read_excel(arquivo, header=2)
    except Exception:
        arquivo.seek(0)
        df = pd.read_excel(arquivo)

    st.subheader("Pré-visualização da planilha (após considerar títulos na linha 3)")
    st.dataframe(df.head())

    colunas = list(df.columns)

    col_descricao = st.selectbox(
        "Coluna de descrição/contexto (opcional, ajuda a entender o pedido)",
        ["(nenhuma)"] + colunas,
        index=0
    )

    col_comentario = st.selectbox(
        "Coluna de comentários do cliente (texto a ser analisado)",
        colunas,
    )

    nome_col_risco = st.text_input(
        "Nome da coluna 1 (Grau de Risco)",
        value="Grau de Risco"
    )

    nome_col_explicacao = st.text_input(
        "Nome da coluna 2 (Explicação do Sentimento)",
        value="Explicação do Sentimento"
    )

    col1, col2 = st.columns(2)
    with col1:
        processar = st.button("Gerar análise de risco e sentimento", use_container_width=True)
    with col2:
        st.write("")

    if processar:
        if nome_col_risco in df.columns or nome_col_explicacao in df.columns:
            st.error("O nome de uma das novas colunas já existe na planilha. Altere os nomes e tente novamente.")
            st.stop()

        progresso = st.progress(0)
        status_text = st.empty()
        riscos = []
        explicacoes = []
        total_linhas = len(df)

        for i, linha in df.iterrows():
            descricao_val = None if col_descricao == "(nenhuma)" else linha[col_descricao]
            comentario_val = linha[col_comentario]
            status_text.text(f"Processando linha {i + 1} de {total_linhas}...")
            try:
                grau, explicacao = analisar_risco_sentimento(descricao_val, comentario_val)
            except Exception:
                grau = "Médio"
                explicacao = "Não foi possível analisar automaticamente. Recomenda-se revisão manual."
            riscos.append(grau)
            explicacoes.append(explicacao)
            progresso.progress(int(((i + 1) / total_linhas) * 100))

        df_saida = df.copy()
        df_saida[nome_col_risco] = riscos
        df_saida[nome_col_explicacao] = explicacoes

        st.success("Análise concluída com sucesso. As duas colunas foram adicionadas à base.")

        st.subheader("Distribuição de Grau de Risco")
        dist_risco = pd.Series(riscos).value_counts().reindex(["Muito Alto", "Alto", "Médio", "Baixo"]).fillna(0).astype(int)
        st.bar_chart(dist_risco)

        st.subheader("Pré-visualização da planilha com curadoria")
        st.dataframe(df_saida.head())

        buffer = io.BytesIO()
        df_saida.to_excel(buffer, index=False)
        buffer.seek(0)

        st.download_button(
            label="Baixar planilha com Grau de Risco e Explicação (.xlsx)",
            data=buffer,
            file_name="base_helps_curadoria_risco_sentimento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
