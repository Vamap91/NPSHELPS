"""
Helps - Curadoria de Risco e Sentimento NPS
Versão 2.0 - Análise Aprimorada

Melhorias implementadas:
1. Análise contextual combinando Descrição + Comentário
2. Detecção de intensificadores e negações
3. Análise de padrões de escrita (caps, pontuação)
4. Score ponderado para classificação mais precisa
5. Prompt engineering otimizado para GPT
6. Fallback heurístico robusto com múltiplas camadas
7. Detecção de ironia/sarcasmo básica
8. Tratamento especial para casos ambíguos
"""

import streamlit as st
import pandas as pd
import io
import json
import re
from openai import OpenAI
from typing import Tuple, Optional, Dict, List
import unicodedata

# ============================================================================
# CONFIGURAÇÃO
# ============================================================================

st.set_page_config(
    page_title="Helps - Curadoria de Risco e Sentimento NPS",
    page_icon="📊",
    layout="wide"
)

client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

# ============================================================================
# DICIONÁRIOS DE ANÁLISE SEMÂNTICA
# ============================================================================

# Palavras com peso de risco (quanto maior, mais grave)
PALAVRAS_RISCO = {
    # Nível Crítico (peso 10) - Ameaças legais e extremo descontentamento
    "procon": 10, "processo": 10, "processar": 10, "advogado": 10, "justiça": 10,
    "reclame aqui": 10, "reclameaqui": 10, "consumidor.gov": 10, "juizado": 10,
    "indenização": 10, "indenizar": 10, "danos morais": 10, "nunca mais": 9,
    "vergonha": 9, "vergonhoso": 9, "absurdo": 9, "inadmissível": 9,
    "inaceitável": 9, "revoltado": 9, "revoltante": 9, "indignado": 9,
    "indignação": 9, "escândalo": 9, "escandaloso": 9, "criminoso": 10,
    "crime": 10, "fraude": 10, "golpe": 10, "enganado": 9, "enganação": 9,
    "mentira": 8, "mentiroso": 9, "calote": 10, "roubo": 10, "roubado": 10,
    
    # Nível Alto (peso 7-8) - Forte insatisfação
    "péssimo": 8, "pessimo": 8, "horrível": 8, "horrivel": 8, "terrível": 8,
    "terrivel": 8, "lixo": 8, "nojo": 8, "nojento": 8, "incompetente": 8,
    "incompetência": 8, "descaso": 8, "abandono": 7, "abandonado": 7,
    "desrespeito": 8, "desrespeitado": 8, "desrespeitoso": 8, "falta de respeito": 8,
    "humilhado": 8, "humilhação": 8, "deboche": 8, "debochado": 8,
    "irresponsável": 8, "irresponsabilidade": 8, "negligente": 8, "negligência": 8,
    "não recomendo": 7, "nao recomendo": 7, "não indico": 7, "nao indico": 7,
    "pior": 7, "decepcionado": 7, "decepcionante": 7, "decepção": 7,
    "frustrado": 7, "frustração": 7, "frustrante": 7, "raiva": 7,
    "ódio": 8, "odio": 8, "detesto": 7, "arrependi": 7, "arrependido": 7,
    
    # Nível Médio-Alto (peso 5-6) - Insatisfação clara
    "ruim": 6, "insatisfeito": 6, "insatisfação": 6, "problema": 5, "problemas": 5,
    "atraso": 5, "atrasado": 5, "atrasaram": 5, "demorado": 5, "demora": 5,
    "demorou": 5, "lento": 5, "lentidão": 5, "erro": 5, "errado": 5,
    "errou": 5, "erros": 5, "falha": 5, "falharam": 5, "defeito": 6,
    "defeituoso": 6, "quebrado": 5, "quebrou": 5, "não funciona": 6,
    "nao funciona": 6, "não funcionou": 6, "nao funcionou": 6,
    "reclamação": 5, "reclamar": 5, "insistir": 5, "insisti": 5,
    "cobrar": 5, "cobrei": 5, "várias vezes": 5, "varias vezes": 5,
    "diversas vezes": 5, "repetidas vezes": 5, "falta": 5, "faltou": 5,
    "faltando": 5, "incompleto": 5, "mal": 5, "malfeito": 6,
    "desorganizado": 5, "desorganização": 5, "bagunça": 5, "confuso": 5,
    "confusão": 5, "perdido": 5, "perderam": 5, "sumiram": 6, "sumiu": 6,
    
    # Nível Médio (peso 3-4) - Ressalvas e incômodos
    "poderia melhorar": 4, "poderia ser melhor": 4, "esperava mais": 4,
    "deixou a desejar": 4, "regular": 3, "médio": 3, "mediano": 3,
    "normal": 2, "ok": 2, "mais ou menos": 3, "nem bom nem ruim": 3,
    "indiferente": 3, "tanto faz": 3, "razoável": 3, "razoavel": 3,
    "aceitável": 3, "aceitavel": 3, "tolerável": 3, "toleravel": 3,
    "chato": 4, "chatice": 4, "incômodo": 4, "incomodo": 4, "desconfortável": 4,
    "desconfortavel": 4, "estranho": 3, "esquisito": 3, "duvidoso": 4,
}

# Palavras positivas (quanto maior, mais positivo)
PALAVRAS_POSITIVAS = {
    # Nível Excelente (peso 10) - Encantamento total
    "perfeito": 10, "perfeita": 10, "impecável": 10, "impecavel": 10,
    "excepcional": 10, "extraordinário": 10, "extraordinario": 10,
    "maravilhoso": 10, "maravilhosa": 10, "sensacional": 10, "fantástico": 10,
    "fantastico": 10, "espetacular": 10, "incrível": 10, "incrivel": 10,
    "surpreendente": 9, "surpreendeu": 9, "superou": 9, "superaram": 9,
    "encantado": 10, "encantada": 10, "encantador": 10, "apaixonado": 9,
    "apaixonada": 9, "amei": 9, "adorei": 9, "melhor": 8, "melhor de todos": 10,
    "nota 10": 10, "nota dez": 10, "10/10": 10, "cinco estrelas": 10,
    "5 estrelas": 10, "recomendo muito": 9, "super recomendo": 10,
    "altamente recomendo": 10, "indico demais": 9,
    
    # Nível Muito Bom (peso 7-8) - Alta satisfação
    "excelente": 8, "ótimo": 8, "otimo": 8, "ótima": 8, "otima": 8,
    "muito bom": 8, "muito boa": 8, "muito bem": 8, "parabéns": 8, "parabens": 8,
    "satisfeito": 7, "satisfeita": 7, "satisfação": 7, "satisfacao": 7,
    "gostei muito": 8, "gostei demais": 8, "adorável": 8, "adoravel": 8,
    "top": 7, "top demais": 8, "show": 7, "demais": 7, "arrasou": 8,
    "mandou bem": 8, "mandaram bem": 8, "caprichado": 8, "capricharam": 8,
    "profissional": 7, "profissionais": 7, "competente": 7, "competentes": 7,
    "eficiente": 7, "eficientes": 7, "eficiência": 7, "eficiencia": 7,
    
    # Nível Bom (peso 5-6) - Satisfação clara
    "bom": 6, "boa": 6, "bem": 5, "gostei": 6, "gosto": 5, "legal": 5,
    "bacana": 5, "tranquilo": 5, "tranquila": 5, "suave": 5, "ok": 4,
    "certinho": 6, "certinha": 6, "correto": 5, "correta": 5,
    "rápido": 6, "rapido": 6, "rápida": 6, "rapida": 6, "rapidez": 6,
    "ágil": 6, "agil": 6, "agilidade": 6, "pontual": 6, "pontualidade": 6,
    "atencioso": 6, "atenciosa": 6, "atenciosos": 6, "atenção": 6, "atencao": 6,
    "educado": 6, "educada": 6, "educados": 6, "cordial": 6, "cordiais": 6,
    "cordialidade": 6, "gentil": 6, "gentis": 6, "gentileza": 6,
    "simpático": 6, "simpatico": 6, "simpática": 6, "simpatica": 6,
    "prestativo": 6, "prestativa": 6, "prestativos": 6, "solicito": 6,
    "solícito": 6, "cuidadoso": 6, "cuidadosa": 6, "cuidado": 5,
    "organizado": 6, "organizada": 6, "limpo": 5, "limpa": 5, "limpeza": 5,
    "qualidade": 6, "confiável": 6, "confiavel": 6, "confiança": 6,
    "confianca": 6, "seguro": 5, "segura": 5, "resolvi": 6, "resolveu": 6,
    "resolvido": 6, "resolveram": 6, "solução": 6, "solucao": 6,
    "funcionou": 6, "funciona": 5, "recomendo": 6, "indico": 6,
    "voltarei": 7, "voltaria": 7, "volto": 6, "retorno": 5,
    
    # Nível Neutro-Positivo (peso 3-4) - Aceitação
    "adequado": 4, "adequada": 4, "suficiente": 4, "dentro do esperado": 4,
    "como esperado": 4, "normal": 3, "padrão": 3, "padrao": 3,
    "cumpriu": 5, "cumpriram": 5, "entregou": 5, "entregaram": 5,
}

# Intensificadores (multiplicam o peso)
INTENSIFICADORES = {
    "muito": 1.5, "demais": 1.5, "extremamente": 2.0, "super": 1.7,
    "mega": 1.7, "ultra": 1.8, "hiper": 1.8, "totalmente": 1.6,
    "completamente": 1.6, "absolutamente": 1.8, "realmente": 1.3,
    "verdadeiramente": 1.4, "incrivelmente": 1.6, "absurdamente": 1.8,
    "ridiculamente": 1.7, "imensamente": 1.6, "profundamente": 1.5,
    "bastante": 1.3, "bem": 1.2, "tão": 1.4, "tanto": 1.3,
}

# Negadores (invertem o sentido)
NEGADORES = {
    "não", "nao", "nunca", "jamais", "nem", "nenhum", "nenhuma",
    "nada", "sem", "tampouco", "sequer",
}

# Indicadores de sarcasmo/ironia
INDICADORES_SARCASMO = {
    "parabéns pela": 0.7, "parabens pela": 0.7, "parabéns pelo": 0.7, 
    "parabens pelo": 0.7, "que maravilha": 0.6, "que ótimo": 0.6,
    "claro que sim": 0.7, "com certeza": 0.8, "obviamente": 0.7,
    "né": 0.8, "ne": 0.8, "viu": 0.8, "hein": 0.7,
}


# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================

def normalizar_texto(texto: str) -> str:
    """Remove acentos e converte para minúsculas."""
    if not texto:
        return ""
    # Remove acentos
    texto_norm = unicodedata.normalize('NFD', texto)
    texto_sem_acento = ''.join(c for c in texto_norm if unicodedata.category(c) != 'Mn')
    return texto_sem_acento.lower().strip()


def detectar_capslock(texto: str) -> float:
    """Detecta uso excessivo de CAPS LOCK (indica intensidade emocional)."""
    if not texto or len(texto) < 10:
        return 1.0
    
    letras = [c for c in texto if c.isalpha()]
    if not letras:
        return 1.0
    
    maiusculas = sum(1 for c in letras if c.isupper())
    proporcao = maiusculas / len(letras)
    
    # Se mais de 50% em caps, aumenta intensidade
    if proporcao > 0.5:
        return 1.3
    elif proporcao > 0.3:
        return 1.15
    return 1.0


def detectar_pontuacao_excessiva(texto: str) -> float:
    """Detecta uso excessivo de pontuação (!!!, ???)."""
    if not texto:
        return 1.0
    
    # Conta sequências de ! ou ?
    exclamacoes = len(re.findall(r'!{2,}', texto))
    interrogacoes = len(re.findall(r'\?{2,}', texto))
    
    intensidade = 1.0 + (exclamacoes * 0.1) + (interrogacoes * 0.05)
    return min(intensidade, 1.5)  # Cap em 1.5


def encontrar_palavras_com_contexto(texto: str, dicionario: dict) -> List[Tuple[str, float, int]]:
    """
    Encontra palavras do dicionário considerando contexto (negadores e intensificadores).
    Retorna lista de (palavra, peso_ajustado, posição).
    """
    if not texto:
        return []
    
    texto_norm = normalizar_texto(texto)
    palavras_texto = texto_norm.split()
    resultados = []
    
    for palavra, peso_base in dicionario.items():
        # Busca a palavra ou expressão no texto
        if palavra in texto_norm:
            # Encontra a posição
            pos = texto_norm.find(palavra)
            
            # Pega contexto anterior (3 palavras antes)
            texto_antes = texto_norm[:pos].split()[-3:]
            
            peso_final = peso_base
            
            # Verifica negadores
            tem_negador = any(neg in texto_antes for neg in NEGADORES)
            if tem_negador:
                # Inverte o sentido
                peso_final = -peso_final * 0.7
            
            # Verifica intensificadores
            for intens, mult in INTENSIFICADORES.items():
                if intens in texto_antes:
                    peso_final *= mult
                    break
            
            resultados.append((palavra, peso_final, pos))
    
    return resultados


def detectar_sarcasmo(texto: str) -> float:
    """
    Detecta possível sarcasmo no texto.
    Retorna fator de ajuste (< 1.0 se detectar sarcasmo).
    """
    if not texto:
        return 1.0
    
    texto_lower = texto.lower()
    
    for indicador, fator in INDICADORES_SARCASMO.items():
        if indicador in texto_lower:
            # Verifica se há palavras negativas no mesmo texto
            texto_norm = normalizar_texto(texto)
            tem_negativo = any(p in texto_norm for p in PALAVRAS_RISCO.keys())
            if tem_negativo:
                return fator
    
    return 1.0


def calcular_score_sentimento(descricao: str, comentario: str) -> Tuple[float, Dict]:
    """
    Calcula score de sentimento com análise detalhada.
    
    Retorna:
    - score: float (negativo = risco, positivo = satisfação)
    - detalhes: dict com breakdown da análise
    """
    texto_completo = f"{descricao or ''} {comentario or ''}".strip()
    
    if not texto_completo or len(texto_completo.strip()) < 3:
        return 0, {"motivo": "texto_vazio"}
    
    # Fatores de intensidade
    fator_caps = detectar_capslock(texto_completo)
    fator_pontuacao = detectar_pontuacao_excessiva(texto_completo)
    fator_sarcasmo = detectar_sarcasmo(texto_completo)
    
    # Encontra palavras
    palavras_negativas = encontrar_palavras_com_contexto(texto_completo, PALAVRAS_RISCO)
    palavras_positivas = encontrar_palavras_com_contexto(texto_completo, PALAVRAS_POSITIVAS)
    
    # Calcula scores
    score_negativo = sum(peso for _, peso, _ in palavras_negativas if peso > 0)
    score_negativo_invertido = sum(abs(peso) for _, peso, _ in palavras_negativas if peso < 0)
    
    score_positivo = sum(peso for _, peso, _ in palavras_positivas if peso > 0)
    score_positivo_invertido = sum(abs(peso) for _, peso, _ in palavras_positivas if peso < 0)
    
    # Aplica fatores
    intensidade_total = fator_caps * fator_pontuacao
    
    # Score final: positivo - negativo, com ajustes
    score_risco = (score_negativo * intensidade_total) - (score_negativo_invertido * 0.5)
    score_satisfacao = (score_positivo * fator_sarcasmo) - (score_positivo_invertido * 0.5)
    
    score_final = score_satisfacao - score_risco
    
    detalhes = {
        "palavras_negativas": [(p, round(w, 2)) for p, w, _ in palavras_negativas],
        "palavras_positivas": [(p, round(w, 2)) for p, w, _ in palavras_positivas],
        "score_risco": round(score_risco, 2),
        "score_satisfacao": round(score_satisfacao, 2),
        "fator_caps": round(fator_caps, 2),
        "fator_pontuacao": round(fator_pontuacao, 2),
        "fator_sarcasmo": round(fator_sarcasmo, 2),
    }
    
    return score_final, detalhes


def score_para_grau_risco(score: float, detalhes: Dict) -> str:
    """Converte score numérico para grau de risco categórico."""
    
    # Verifica casos especiais
    palavras_neg = detalhes.get("palavras_negativas", [])
    palavras_pos = detalhes.get("palavras_positivas", [])
    
    # Se tem palavras de risco crítico (peso >= 9), é Muito Alto independente
    tem_critico = any(abs(w) >= 9 for _, w in palavras_neg)
    if tem_critico:
        return "Muito Alto"
    
    # Se score muito negativo
    if score <= -15:
        return "Muito Alto"
    elif score <= -8:
        return "Alto"
    elif score <= -3:
        return "Médio"
    elif score < 5:
        # Zona neutra - depende do contexto
        if palavras_neg and not palavras_pos:
            return "Médio"
        elif palavras_pos and not palavras_neg:
            return "Baixo"
        elif palavras_neg and palavras_pos:
            return "Médio"
        else:
            return "Baixo"  # Sem indicadores claros
    else:
        return "Baixo"


def gerar_explicacao_heuristica(score: float, detalhes: Dict, comentario: str) -> str:
    """Gera explicação baseada na análise heurística."""
    
    palavras_neg = detalhes.get("palavras_negativas", [])
    palavras_pos = detalhes.get("palavras_positivas", [])
    
    # Pega trecho do comentário
    trecho = (comentario[:80] + "...") if len(comentario or "") > 80 else (comentario or "")
    trecho = trecho.replace('"', "'")
    
    if not comentario or len(comentario.strip()) < 3:
        return "Sem comentário relevante para análise."
    
    if score <= -15:
        principais = [p for p, _ in palavras_neg[:3]]
        return f"Comentário expressa forte insatisfação com indicadores críticos ({', '.join(principais)}). Requer atenção urgente."
    
    elif score <= -8:
        principais = [p for p, _ in palavras_neg[:2]]
        return f"Cliente demonstra insatisfação significativa. Termos identificados: {', '.join(principais)}."
    
    elif score <= -3:
        if palavras_neg and palavras_pos:
            return f"Feedback misto com ressalvas. Cliente menciona pontos positivos mas também críticas."
        else:
            principais = [p for p, _ in palavras_neg[:2]]
            return f"Cliente expressa incômodo ou frustração moderada ({', '.join(principais)})."
    
    elif score < 5:
        if palavras_pos:
            return f"Comentário neutro com tendência positiva. Cliente parece satisfeito com ressalvas."
        elif palavras_neg:
            return f"Comentário neutro com algumas ressalvas mencionadas."
        else:
            return f"Comentário neutro sem indicadores fortes de satisfação ou insatisfação."
    
    else:
        if score >= 15:
            principais = [p for p, _ in palavras_pos[:3]]
            return f"Cliente muito satisfeito! Elogio claro com termos positivos ({', '.join(principais)})."
        else:
            principais = [p for p, _ in palavras_pos[:2]]
            return f"Feedback positivo. Cliente demonstra satisfação ({', '.join(principais)})."


# ============================================================================
# FUNÇÃO HEURÍSTICA ROBUSTA (FALLBACK)
# ============================================================================

def heuristica_risco_explicacao(descricao: Optional[str], comentario: Optional[str]) -> Tuple[str, str]:
    """
    Análise heurística robusta para classificação de risco.
    Usada como fallback quando a IA falha.
    """
    # Combina descrição e comentário
    texto = f"{descricao or ''} {comentario or ''}".strip()
    
    if not texto or len(texto) < 3:
        return "Baixo", "Sem comentário relevante para análise."
    
    # Calcula score de sentimento
    score, detalhes = calcular_score_sentimento(descricao, comentario)
    
    # Converte para grau de risco
    grau = score_para_grau_risco(score, detalhes)
    
    # Gera explicação
    explicacao = gerar_explicacao_heuristica(score, detalhes, comentario)
    
    return grau, explicacao


# ============================================================================
# FUNÇÃO PRINCIPAL COM IA
# ============================================================================

def analisar_risco_sentimento(descricao: Optional[str], comentario: Optional[str]) -> Tuple[str, str]:
    """
    Analisa risco e sentimento usando OpenAI com fallback heurístico.
    """
    # Tratamento de comentário vazio
    if not comentario or len(str(comentario).strip()) < 3:
        return "Baixo", "Sem comentário relevante para análise."
    
    comentario = str(comentario).strip()
    descricao = str(descricao).strip() if descricao else ""
    
    # Monta o prompt otimizado
    prompt = f"""Analise o seguinte feedback de cliente de uma pesquisa NPS.

CONTEXTO DO ATENDIMENTO: {descricao if descricao else "Não especificado"}

COMENTÁRIO DO CLIENTE: "{comentario}"

TAREFA: Classifique o GRAU DE RISCO para a empresa e explique o SENTIMENTO do cliente.

REGRAS DE CLASSIFICAÇÃO (siga rigorosamente):

1. **Muito Alto** - Use APENAS quando houver:
   - Ameaça legal explícita (Procon, processo, advogado, Reclame Aqui)
   - Palavras de revolta extrema (absurdo, vergonha, inadmissível, escândalo)
   - Acusações graves (fraude, golpe, roubo, crime, calote)
   - Declaração "nunca mais volto/uso/compro"
   - CAPS LOCK com xingamentos ou ofensas

2. **Alto** - Use quando houver:
   - Insatisfação forte e clara (péssimo, horrível, terrível, nojento)
   - Declaração de arrependimento ou decepção profunda
   - Múltiplos problemas graves relatados
   - Indicação de que não recomendaria
   - Tom de raiva ou frustração intensa

3. **Médio** - Use quando houver:
   - Reclamação moderada (ruim, demorado, problema, erro)
   - Feedback misto (elogios E críticas)
   - Ressalvas ou sugestões de melhoria
   - Tom de incômodo mas sem revolta
   - Expectativas parcialmente atendidas

4. **Baixo** - Use quando houver:
   - Elogio claro (ótimo, excelente, muito bom, parabéns)
   - Satisfação expressa (gostei, recomendo, voltarei)
   - Agradecimento ou reconhecimento positivo
   - Comentário neutro sem queixas
   - Menção de experiência agradável

ATENÇÃO ESPECIAL:
- "Agendamento rápido e cordialidade no atendimento" = BAIXO (elogio claro)
- "Atendimento ok mas demorou" = MÉDIO (misto)
- "Péssimo, nunca mais volto" = MUITO ALTO (revolta + declaração)
- Comentário vazio ou sem sentido = BAIXO

Responda SOMENTE com JSON válido (sem markdown, sem texto antes/depois):
{{"grau_risco": "Muito Alto|Alto|Médio|Baixo", "explicacao": "Frase curta explicando o sentimento"}}"""

    try:
        response = client.responses.create(
            model="gpt-4o-mini",
            instructions="""Você é um especialista em análise de sentimento e experiência do cliente.
Sua tarefa é classificar o risco reputacional de feedbacks NPS.
Seja preciso: elogios claros = Baixo, críticas severas = Alto/Muito Alto.
Responda APENAS com JSON válido, sem nenhum texto adicional.""",
            input=prompt,
            max_output_tokens=250,
            temperature=0.1  # Baixa temperatura para consistência
        )
        
        # Extrai o texto da resposta
        resposta_texto = ""
        if hasattr(response, 'output'):
            for item in response.output:
                if hasattr(item, 'content'):
                    for content in item.content:
                        if hasattr(content, 'text'):
                            resposta_texto += content.text
        
        if not resposta_texto:
            return heuristica_risco_explicacao(descricao, comentario)
        
        # Tenta extrair JSON
        # Remove possíveis backticks de markdown
        resposta_texto = re.sub(r'```json\s*', '', resposta_texto)
        resposta_texto = re.sub(r'```\s*', '', resposta_texto)
        
        # Encontra o JSON
        match = re.search(r'\{[^{}]*\}', resposta_texto, re.DOTALL)
        if not match:
            return heuristica_risco_explicacao(descricao, comentario)
        
        json_str = match.group(0)
        dados = json.loads(json_str)
        
        grau = dados.get("grau_risco", "").strip()
        explicacao = dados.get("explicacao", "").strip()
        
        # Valida o grau
        graus_validos = ["Muito Alto", "Alto", "Médio", "Baixo"]
        if grau not in graus_validos:
            # Tenta normalizar
            grau_lower = grau.lower()
            if "muito alto" in grau_lower:
                grau = "Muito Alto"
            elif "alto" in grau_lower:
                grau = "Alto"
            elif "médio" in grau_lower or "medio" in grau_lower:
                grau = "Médio"
            elif "baixo" in grau_lower:
                grau = "Baixo"
            else:
                return heuristica_risco_explicacao(descricao, comentario)
        
        if not explicacao:
            _, explicacao = heuristica_risco_explicacao(descricao, comentario)
        
        return grau, explicacao
        
    except Exception as e:
        # Em caso de erro, usa heurística
        return heuristica_risco_explicacao(descricao, comentario)


# ============================================================================
# INTERFACE STREAMLIT
# ============================================================================

def main():
    # Header
    st.markdown("""
    <div style="text-align: center; padding: 1rem 0;">
        <h1>📊 Helps - Curadoria de Risco e Sentimento NPS</h1>
        <p style="color: #666; font-size: 1.1rem;">
            Análise inteligente de comentários para classificação de risco reputacional
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Sidebar com instruções
    with st.sidebar:
        st.header("📋 Como usar")
        st.markdown("""
        **Passo a passo:**
        
        1. **Upload** - Envie o arquivo Excel com a base NPS
        2. **Selecione** - Indique as colunas de descrição e comentário
        3. **Processe** - Clique para gerar a análise
        4. **Download** - Baixe o Excel enriquecido
        
        ---
        
        **Graus de Risco:**
        
        🔴 **Muito Alto** - Ameaça legal ou revolta extrema
        
        🟠 **Alto** - Forte insatisfação
        
        🟡 **Médio** - Ressalvas ou feedback misto
        
        🟢 **Baixo** - Satisfeito ou elogio
        
        ---
        
        **Sobre a análise:**
        
        O sistema usa IA (GPT-4o-mini) combinada com análise heurística para garantir precisão.
        
        ✅ Não altera o texto original
        
        ✅ Adiciona apenas classificação
        """)
    
    # Upload do arquivo
    st.subheader("1️⃣ Upload do arquivo")
    
    arquivo = st.file_uploader(
        "Envie o arquivo Excel (.xlsx) com a base NPS",
        type=["xlsx", "xls"],
        help="O título das colunas deve estar na linha 3 do arquivo"
    )
    
    if arquivo:
        try:
            # Lê o Excel com header na linha 3 (header=2)
            df = pd.read_excel(arquivo, header=2)
            
            st.success(f"✅ Arquivo carregado: **{len(df)} registros** encontrados")
            
            # Preview
            with st.expander("📄 Prévia dos dados (primeiras 5 linhas)", expanded=True):
                st.dataframe(df.head(), use_container_width=True)
            
            # Seleção de colunas
            st.subheader("2️⃣ Configuração das colunas")
            
            colunas = df.columns.tolist()
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Tenta encontrar coluna de descrição
                idx_desc = 0
                for i, c in enumerate(colunas):
                    if "descricao" in c.lower() or "descrição" in c.lower():
                        idx_desc = i + 1  # +1 porque vamos adicionar "(nenhuma)"
                        break
                
                col_descricao = st.selectbox(
                    "Coluna de descrição/contexto:",
                    ["(nenhuma)"] + colunas,
                    index=idx_desc,
                    help="Contexto do atendimento (opcional)"
                )
            
            with col2:
                # Tenta encontrar coluna de comentário
                idx_coment = 0
                for i, c in enumerate(colunas):
                    if "comentario" in c.lower() or "comentário" in c.lower():
                        idx_coment = i
                        break
                
                col_comentario = st.selectbox(
                    "Coluna de comentários do cliente:",
                    colunas,
                    index=idx_coment,
                    help="Texto livre do cliente"
                )
            
            # Nomes das novas colunas
            st.markdown("**Nomes das colunas de saída:**")
            col3, col4 = st.columns(2)
            
            with col3:
                nome_col_risco = st.text_input(
                    "Nome da coluna de risco:",
                    value="Grau de Risco"
                )
            
            with col4:
                nome_col_explicacao = st.text_input(
                    "Nome da coluna de explicação:",
                    value="Explicação do Sentimento"
                )
            
            # Botão de processamento
            st.subheader("3️⃣ Processar análise")
            
            if st.button("🚀 Gerar análise de risco e sentimento", type="primary", use_container_width=True):
                
                # Validações
                if nome_col_risco in colunas or nome_col_explicacao in colunas:
                    st.error("❌ Os nomes das novas colunas já existem na planilha. Escolha nomes diferentes.")
                    return
                
                # Processamento
                riscos = []
                explicacoes = []
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                total = len(df)
                
                for i, linha in df.iterrows():
                    # Atualiza progresso
                    progress = (i + 1) / total
                    progress_bar.progress(progress)
                    status_text.text(f"Processando: {i + 1}/{total} ({progress:.0%})")
                    
                    # Obtém valores
                    descricao_val = None if col_descricao == "(nenhuma)" else linha.get(col_descricao)
                    comentario_val = linha.get(col_comentario)
                    
                    # Converte para string se necessário
                    if pd.notna(descricao_val):
                        descricao_val = str(descricao_val)
                    else:
                        descricao_val = None
                        
                    if pd.notna(comentario_val):
                        comentario_val = str(comentario_val)
                    else:
                        comentario_val = None
                    
                    # Analisa
                    try:
                        grau, explicacao = analisar_risco_sentimento(descricao_val, comentario_val)
                    except Exception as e:
                        grau, explicacao = heuristica_risco_explicacao(descricao_val, comentario_val)
                    
                    riscos.append(grau)
                    explicacoes.append(explicacao)
                
                progress_bar.progress(1.0)
                status_text.text("✅ Processamento concluído!")
                
                # Cria DataFrame de saída
                df_saida = df.copy()
                df_saida[nome_col_risco] = riscos
                df_saida[nome_col_explicacao] = explicacoes
                
                # Resultados
                st.subheader("4️⃣ Resultados")
                
                # Estatísticas
                col_stats1, col_stats2, col_stats3, col_stats4 = st.columns(4)
                
                contagem = pd.Series(riscos).value_counts()
                
                with col_stats1:
                    st.metric(
                        "🔴 Muito Alto",
                        contagem.get("Muito Alto", 0),
                        help="Requerem atenção urgente"
                    )
                
                with col_stats2:
                    st.metric(
                        "🟠 Alto",
                        contagem.get("Alto", 0),
                        help="Insatisfação significativa"
                    )
                
                with col_stats3:
                    st.metric(
                        "🟡 Médio",
                        contagem.get("Médio", 0),
                        help="Ressalvas ou feedback misto"
                    )
                
                with col_stats4:
                    st.metric(
                        "🟢 Baixo",
                        contagem.get("Baixo", 0),
                        help="Satisfeitos ou sem problemas"
                    )
                
                # Gráfico de distribuição
                st.markdown("**Distribuição de Risco:**")
                
                # Prepara dados para o gráfico na ordem correta
                ordem_risco = ["Muito Alto", "Alto", "Médio", "Baixo"]
                dados_grafico = pd.DataFrame({
                    "Grau de Risco": ordem_risco,
                    "Quantidade": [contagem.get(r, 0) for r in ordem_risco]
                }).set_index("Grau de Risco")
                
                st.bar_chart(dados_grafico)
                
                # Preview da saída
                with st.expander("📄 Prévia do resultado (primeiras 10 linhas)", expanded=True):
                    # Mostra apenas as colunas relevantes
                    colunas_exibir = [col_comentario, nome_col_risco, nome_col_explicacao]
                    if col_descricao != "(nenhuma)":
                        colunas_exibir.insert(0, col_descricao)
                    
                    st.dataframe(df_saida[colunas_exibir].head(10), use_container_width=True)
                
                # Download
                st.subheader("5️⃣ Download")
                
                buffer = io.BytesIO()
                df_saida.to_excel(buffer, index=False, engine='openpyxl')
                buffer.seek(0)
                
                st.download_button(
                    label="📥 Baixar planilha com Grau de Risco e Explicação (.xlsx)",
                    data=buffer,
                    file_name="base_helps_curadoria_risco_sentimento.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.success("🎉 Análise concluída! Clique acima para baixar o arquivo.")
                
        except Exception as e:
            st.error(f"❌ Erro ao processar arquivo: {str(e)}")
            st.info("💡 Verifique se o arquivo está no formato correto e se os títulos estão na linha 3.")


if __name__ == "__main__":
    main()
