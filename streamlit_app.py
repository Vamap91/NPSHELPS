import streamlit as st
import pandas as pd
import io
import json
import re
from openai import OpenAI

st.set_page_config(page_title="Helps - Curadoria de Risco e Sentimento", layout="wide")

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

def heuristica_risco_explicacao(descricao, comentario):
    if comentario is None:
        return "Baixo", "Sem comentário relevante para análise."
    texto = str(comentario).strip()
    if texto == "":
        return "Baixo", "Sem comentário relevante para análise."
    t = texto.lower()
    palavras_muito_negativas = [
        "péssimo", "pessimo", "horrível", "horrivel", "terrível", "terrivel",
        "nunca mais", "absurdo", "ridículo", "ridiculo", "engan", "procon",
        "reclame aqui", "processo", "processar", "cancelar serviço", "cancelar o serviço"
    ]
    palavras_negativas = [
        "ruim", "demorado", "demora", "atraso", "demorou", "problema",
        "insatisfeito", "insatisfação", "nao gostei", "não gostei",
        "falta de", "demorando", "espera", "fila", "reclamação", "reclamacao"
    ]
    palavras_positivas = [
        "ótimo", "otimo", "bom", "boa", "excelente", "maravilhoso",
        "rápido", "rapido", "ágil", "agil", "cordial", "educado",
        "perfeito", "muito bom", "muito boa", "ador", "satisfeito", "satisfatória", "satisfatoria"
    ]
    muito_negativo = any(p in t for p in palavras_muito_negativas)
    negativo = any(p in t for p in palavras_negativas)
    positivo = any(p in t for p in palavras_positivas)
    if muito_negativo:
        grau = "Muito Alto"
    elif negativo and not positivo:
        grau = "Alto"
    elif positivo and not (negativo or muito_negativo):
        grau = "Baixo"
    else:
        grau = "Médio"
    trecho = texto
    if len(trecho) > 160:
        trecho = trecho[:157] + "..."
    if grau == "Baixo":
        sentimento = "claramente positivo ou elogioso"
        recomendacao = "reforçar os pontos fortes e manter o padrão de atendimento."
    elif grau == "Médio":
        sentimento = "misto ou neutro, com possíveis pontos de melhoria"
        recomendacao = "acompanhar, mas sem urgência imediata."
    elif grau == "Alto":
        sentimento = "negativo, com insatisfação relevante"
        recomendacao = "analisar e tratar em curto prazo para evitar desgaste."
    else:
        sentimento = "muito negativo e crítico"
        recomendacao = "priorizar o tratamento imediato para reduzir risco de reclamações mais graves."
    explicacao = f"O comentário é {sentimento}. Exemplo do texto do cliente: \"{trecho}\". Recomenda-se {recomendacao}"
    return grau, explicacao

def analisar_risco_sentimento(descricao, comentario):
    if comentario is None or str(comentario).strip() == "":
        return "Baixo", "Sem comentário relevante para análise."
    descricao_texto = "" if descricao is None else str(descricao)
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
        "Se o comentário for claramente elogioso (por exemplo, 'agendamento rápido e cordialidade no atendimento'), use GRAU DE RISCO = Baixo.\n\n"
        "Use a descrição do atendimento apenas como contexto complementar.\n\n"
        f"Descrição do atendimento: {descricao_texto}\n"
        f"Comentário do cliente: {comentario_texto}\n\n"
        "Responda apenas com um JSON puro e válido, sem texto antes ou depois, exatamente no formato:\n"
        "{ \"grau_risco\": \"Muito Alto|Alto|Médio|Baixo\", \"explicacao\": \"texto curto explicando o sentimento\" }"
    )
    try:
        response = client.responses.create(
            model="gpt-4o-mini",
            instructions="Você é um especialista em experiência do cliente, classificando risco e explicando o sentimento de forma clara e concisa.",
            input=prompt,
            max_output_tokens=200
        )
        content = response.output[0].content[0].text
        texto = content.value if hasattr(content, "value") else str(content)
    except Exception:
        return heuristica_risco_explicacao(descricao, comentario)
    m = re.search(r"\{.*\}", texto, re.DOTALL)
    if not m:
        return heuristica_risco_explicacao(descricao, comentario)
    try:
        dados = json.loads(m.group(0))
        grau = str(dados.get("grau_risco", "")).strip()
        explicacao = str(dados.get("explicacao", "")).strip()
    except Exception:
        return heuristica_risco_explicacao(descricao, comentario)
    if grau not in ["Muito Alto", "Alto", "Médio", "Baixo"]:
        return heuristica_risco_explicacao(descricao, comentario)
    if explicacao == "":
        return heuristica_risco_explicacao(descricao, comentario)
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
    st.markdown("2. Títulos da planilha começam na linha 3")
    st.markdown("3. Confirme as colunas: Descricao e Comentario")
    st.markdown("4. Clique em **Gerar análise** para criar as duas novas colunas")
    st.markdown("---")
    st.markdown("Os dados do cliente não são alterados, apenas enriquecidos com a curadoria.")

arquivo = st.file_uploader("Envie o arquivo Excel (.xlsx)", type=["xlsx", "xls"])

if arquivo is not None:
    try:
        df = pd.read_excel(arquivo, header=2)
    except Exception:
        arquivo.seek(0)
        df = pd.read_excel(arquivo, header=2)

    st.subheader("Pré-visualização da planilha (títulos na linha 3)")
    st.dataframe(df.head())

    colunas = list(df.columns)

    if "Descricao" in colunas:
        idx_desc = colunas.index("Descricao") + 1
    else:
        idx_desc = 0

    col_descricao = st.selectbox(
        "Coluna de descrição/contexto (opcional, ajuda a entender o pedido)",
        ["(nenhuma)"] + colunas,
        index=idx_desc if idx_desc < len(colunas) + 1 else 0
    )

    if "Comentario" in colunas:
        idx_coment = colunas.index("Comentario")
    else:
        idx_coment = 0

    col_comentario = st.selectbox(
        "Coluna de comentários do cliente (texto a ser analisado)",
        colunas,
        index=idx_coment if idx_coment < len(colunas) else 0
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
                grau, explicacao = heuristica_risco_explicacao(descricao_val, comentario_val)
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
