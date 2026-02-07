# -*- coding: utf-8 -*-
"""
Interface Streamlit para Análise de Movimentações Previdenciárias
ArcelorMittal - Sistema de Validação Automatizada
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from datetime import datetime
import io

# ============================================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================================

st.set_page_config(
    page_title="Análise ArcelorMittal",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS Customizado
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        padding: 1rem;
        background: linear-gradient(90deg, #e8f4f8 0%, #ffffff 100%);
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #1f77b4;
    }
    .error-card {
        background-color: #ffe6e6;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #ff4444;
    }
    .success-card {
        background-color: #e6ffe6;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #44ff44;
    }
    .info-card {
        background-color: #e6f3ff;
        padding: 1rem;
        border-radius: 10px;
        border-left: 4px solid #4488ff;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# BASE DE CONHECIMENTO (Cópia das células 2 e 3)
# ============================================================================


@st.cache_data
def carregar_base_conhecimento():
    """Carrega códigos e regras de negócio"""
    codigos_data = {
        'CODIGO': [11100, 11200, 16000, 15000, 14000, 13000, 12000, 21000, 22000,
                   23000, 24100, 24200, 31100, 31200, 31300, 31000, 32000, 33000, 34000],
        'DESCRICAO': ['Aposentadoria Normal', 'Aposentadoria por Invalidez',
                      'Outros Benefícios de Prestação Única',
                      'Pecúlio (Pagamento a Herdeiros)', 'Pensão por Morte',
                      'Auxílio Único (Natalidade/Funeral)', 'Auxílio Continuado (Afastamento Doença)',
                      'BPD (Benefício Proporcional Diferido)',
                      'Autopatrocinado', 'Resgate Total',
                      'Portabilidade Saída', 'Portabilidade Entrada',
                      'Ativo com Contrib. Empresa', 'Ativo com Contrib. Empresa + Participante',
                      'Ativo Contrib. Só Participante',
                      'Consolidado Ativos', 'Consolidado Aposentados',
                      'Consolidado Pensionistas', 'Designados (Dependentes)'],
        'TIPO': ['Benefício', 'Benefício', 'Benefício', 'Benefício', 'Benefício',
                 'Benefício', 'Benefício', 'Instituto', 'Instituto', 'Instituto',
                 'Instituto', 'Instituto', 'População', 'População', 'População',
                 'Consolidador', 'Consolidador', 'Consolidador', 'População']
    }
    df_codigos = pd.DataFrame(codigos_data)

    # CÓDIGOS CONSOLIDADORES QUE DEVEM SER IGNORADOS NA ANÁLISE
    CODIGOS_IGNORAR = set()  # Não ignorar mais códigos consolidadores

    # CÓDIGOS QUE CAUSAM RUÍDO EM MÚLTIPLAS SAÍDAS (devem ser filtrados ao calcular saídas líquidas)
    CODIGOS_RUIDO_SAIDA = {31100, 31200, 31300, 31000, 32000, 33000, 11000, 14000}

    CONTAS_ZERAGEM_ANUAL = {13000, 15000, 16000, 23000, 24100, 24200}
    CODIGOS_ADMISSAO = {31100, 31200}
    CODIGOS_ATIVOS = {31100, 31200, 31300}
    CODIGOS_SEM_RETORNO = {21000}
    CODIGOS_DESLIGAMENTO_RUIDO = {21000, 22000, 31300}

    CODIGOS_ENTRADA_INDEPENDENTE = {13000, 24100, 24200}

    regras_validas = [
        (31100, 11100), (31200, 11100), (31300,
                                         11100), (21000, 11100), (22000, 11100),
        (31100, 11200), (31200, 11200), (31300, 11200), (12000,
                                                         11200), (21000, 11200), (22000, 11200),
        (31100, 12000), (31200, 12000), (31300, 12000), (22000, 12000),
        (31100, 13000), (31200, 13000), (31300,
                                         13000), (11100, 13000), (11200, 13000),
        (31100, 14000), (31200, 14000), (31300, 14000), (11100,
                                                         14000), (11200, 14000), (22000, 14000),
        (31100, 15000), (31200, 15000), (31300,
                                         15000), (11100, 15000), (11200, 15000),
        (14000, 15000), (21000, 15000), (22000, 15000),
        (31100, 16000), (31200, 16000), (31300,
                                         16000), (11100, 16000), (11200, 16000),
        (21000, 16000), (22000, 16000),
        (31100, 21000), (31200, 21000), (31300, 21000), (22000, 21000),
        (31100, 22000), (31200, 22000), (31300, 22000),
        (31100, 23000), (31200, 23000), (31300,
                                         23000), (21000, 23000), (22000, 23000),
        (31100, 24100), (31200, 24100), (31300,
                                         24100), (21000, 24100), (22000, 24100),
        (31100, 24200), (31200, 24200), (31300,
                                         24200), (21000, 24200), (22000, 24200),
        (11100, 24200), (11200, 24200),
        (31200, 31100), (12000, 31100), (22000, 31100),
        (31100, 31200), (12000, 31200), (22000, 31200),
        (31100, 31300), (31200, 31300), (21000, 31300), (22000, 31300)
    ]

    return df_codigos, set(regras_validas), {
        'CODIGOS_IGNORAR': CODIGOS_IGNORAR,
        'CODIGOS_ATIVOS': CODIGOS_ATIVOS,
        'CODIGOS_ADMISSAO': CODIGOS_ADMISSAO,
        'CODIGOS_DESLIGAMENTO_RUIDO': CODIGOS_DESLIGAMENTO_RUIDO,
        'CODIGOS_RUIDO_SAIDA': CODIGOS_RUIDO_SAIDA,
        'CODIGOS_ENTRADA_INDEPENDENTE': CODIGOS_ENTRADA_INDEPENDENTE
    }


def get_descricao(codigo, df_codigos_ref):
    """Retorna descrição do código"""
    res = df_codigos_ref[df_codigos_ref['CODIGO'] == codigo]['DESCRICAO']
    return res.iloc[0] if not res.empty else f'Código Desconhecido ({codigo})'


def formatar_nome_participante(nome):
    """Formata nome do participante"""
    if pd.isna(nome):
        return "Nome não informado"
    return str(nome).strip().title()

# ============================================================================
# MOTOR DE ANÁLISE (Cópia da célula 4 - simplificada)
# ============================================================================


def analisar_movimentacoes_mes(df_mov, df_codigos, regras_validas, constantes, mes_analise=None):
    """Motor de análise principal"""

    if mes_analise is None:
        mes_analise = df_mov['ANO MES'].max()

    df_mes = df_mov[df_mov['ANO MES'] == mes_analise].copy()

    if df_mes.empty:
        return df_mes

    df_mes['ANALISE'] = 'OK'
    df_mes['TIPO_PASSO'] = 'Indefinido'
    df_mes['INTERPRETACAO'] = ''
    df_mes['GRAVIDADE'] = 'OK'

    grouped = df_mes.groupby('CODIGO ORGANIZACAO NOME')

    stats = {'total': len(grouped), 'erros': 0, 'info': 0, 'ok': 0}

    for nome_participante, group in grouped:

        # Validação 1: Múltiplas situações ativas
        if 'PLANO' in group.columns:
            for plano in group['PLANO'].unique():
                group_plano = group[group['PLANO'] == plano]
                entradas_plano = group_plano[group_plano['MOVIMENTO'] == 'ENTRADA']
                codigos_ativos_entrada = set(
                    entradas_plano['CODIGO BENEFICIO']) & constantes['CODIGOS_ATIVOS']

                if len(codigos_ativos_entrada) > 1:
                    msg = f"ERRO: Múltiplas situações ativas no Plano {plano}"
                    df_mes.loc[group_plano.index, 'ANALISE'] = msg
                    df_mes.loc[group_plano.index, 'GRAVIDADE'] = 'ERRO'
                    stats['erros'] += 1
                    continue

        # Validação 2: Pensão vs Pecúlio
        entradas = group[group['MOVIMENTO'] == 'ENTRADA']
        codigos_entrada_set = set(entradas['CODIGO BENEFICIO'])

        # Análise de transições
        saidas = group[group['MOVIMENTO'] == 'SAIDA']
        codigos_saida_set = set(saidas['CODIGO BENEFICIO'])

        # Validação 1.5: Código 14000 (Pensão) isolado sem 33000
        if 14000 in codigos_entrada_set or 14000 in codigos_saida_set:
            if 14000 in codigos_saida_set and 33000 not in codigos_saida_set:
                msg = "ERRO: Saída no código 14000 (Pensão por Morte) sem saída correspondente na conta 33000 (Consolidado Pensionistas)"
                df_mes.loc[group.index, 'ANALISE'] = msg
                df_mes.loc[group.index, 'GRAVIDADE'] = 'ERRO'
                stats['erros'] += 1
                continue
            if 14000 in codigos_entrada_set and 33000 not in codigos_entrada_set:
                msg = "ERRO: Entrada no código 14000 (Pensão por Morte) sem entrada correspondente na conta 33000 (Consolidado Pensionistas)"
                df_mes.loc[group.index, 'ANALISE'] = msg
                df_mes.loc[group.index, 'GRAVIDADE'] = 'ERRO'
                stats['erros'] += 1
                continue

        if 14000 in codigos_entrada_set and 15000 in codigos_entrada_set:
            msg = "ERRO: PENSÃO e PECÚLIO no mesmo mês"
            df_mes.loc[group.index, 'ANALISE'] = msg
            df_mes.loc[group.index, 'GRAVIDADE'] = 'ERRO'
            stats['erros'] += 1
            continue

        # Validação 3: Códigos consolidadores devem ter movimentações correspondentes
        if 'CODIGO BENEFICIO' in entradas.columns:
            # Valida código 32000 (Consolidado Aposentados)
            if 32000 in codigos_entrada_set or 32000 in codigos_saida_set:
                # 32000 deve refletir movimentações em 11000, 11100, 11200
                codigos_aposentados = {11000, 11100, 11200}
                movs_aposentados = codigos_entrada_set.union(codigos_saida_set) & codigos_aposentados
                
                if not movs_aposentados:
                    msg = "ERRO: Código 32000 lançado sem movimentação correspondente nas contas de aposentados (11000, 11100, 11200)"
                    df_mes.loc[group.index, 'ANALISE'] = msg
                    df_mes.loc[group.index, 'GRAVIDADE'] = 'ERRO'
                    stats['erros'] += 1
                    continue

            # Valida código 31300 (Consolidado Ativos - Só Participante)
            if 31300 in codigos_entrada_set or 31300 in codigos_saida_set:
                # 31300 deve refletir movimentações em 21000, 22000
                codigos_instituto = {21000, 22000}
                movs_instituto = codigos_entrada_set.union(codigos_saida_set) & codigos_instituto
                
                if not movs_instituto:
                    msg = "ERRO: Código 31300 lançado sem movimentação correspondente nas contas de instituto (21000, 22000)"
                    df_mes.loc[group.index, 'ANALISE'] = msg
                    df_mes.loc[group.index, 'GRAVIDADE'] = 'ERRO'
                    stats['erros'] += 1
                    continue
            

        # Validação 4: Código 33000 deve sempre acompanhar 14000
        if 33000 in codigos_entrada_set or 33000 in codigos_saida_set:
            if 33000 in codigos_entrada_set and 14000 not in codigos_entrada_set:
                msg = "ERRO: Entrada no código 33000 (Consolidado Pensionistas) sem entrada correspondente na conta 14000 (Pensão por Morte)"
                df_mes.loc[group.index, 'ANALISE'] = msg
                df_mes.loc[group.index, 'GRAVIDADE'] = 'ERRO'
                stats['erros'] += 1
                continue
            if 33000 in codigos_saida_set and 14000 not in codigos_saida_set:
                msg = "ERRO: Saída no código 33000 (Consolidado Pensionistas) sem saída correspondente na conta 14000 (Pensão por Morte)"
                df_mes.loc[group.index, 'ANALISE'] = msg
                df_mes.loc[group.index, 'GRAVIDADE'] = 'ERRO'
                stats['erros'] += 1
                continue

            
            # Valida código 33000 (Consolidado Pensionistas)
            if 33000 in codigos_entrada_set or 33000 in codigos_saida_set:
                # 33000 deve refletir movimentações em 14000
                if 14000 not in codigos_entrada_set and 14000 not in codigos_saida_set:
                    msg = "ERRO: Código 33000 lançado sem movimentação correspondente na conta de pensão (14000)"
                    df_mes.loc[group.index, 'ANALISE'] = msg
                    df_mes.loc[group.index, 'GRAVIDADE'] = 'ERRO'
                    stats['erros'] += 1
                    continue

        codigos_entrada_independentes = codigos_entrada_set & constantes.get(
            'CODIGOS_ENTRADA_INDEPENDENTE', set())
        # 34000 (Designados/Dependentes) também é independente
        codigos_entrada_independentes = codigos_entrada_independentes | (codigos_entrada_set & {34000})
        codigos_entrada_principal = codigos_entrada_set - codigos_entrada_independentes

        codigos_intermediarios = codigos_saida_set.intersection(
            codigos_entrada_principal)

        # Calcula saídas e entradas líquidas
        saidas_liquidas_brutas = codigos_saida_set - codigos_entrada_principal
        entradas_liquidas = codigos_entrada_principal - codigos_saida_set
        entradas_independentes_liquidas = codigos_entrada_independentes

        # FILTRAR CÓDIGOS DE RUÍDO APENAS SE HOUVER MÚLTIPLAS SAÍDAS
        # Se há apenas 1 saída, ela é legítima (mesmo sendo 31100, 31200, etc)
        # Se há múltiplas saídas, remove os códigos consolidadores que são ruído
        if len(saidas_liquidas_brutas) > 1:
            saidas_liquidas = saidas_liquidas_brutas - \
                constantes['CODIGOS_RUIDO_SAIDA']
        else:
            saidas_liquidas = saidas_liquidas_brutas

        if len(entradas_liquidas) > 1:
            entradas_liquidas_filtradas = entradas_liquidas - constantes['CODIGOS_RUIDO_SAIDA']
        else:
            entradas_liquidas_filtradas = entradas_liquidas

        # Classificação de passos
        for idx, row in group.iterrows():
            cod = row['CODIGO BENEFICIO']
            mov = row['MOVIMENTO']

            if cod in entradas_independentes_liquidas and mov == 'ENTRADA':
                df_mes.loc[idx, 'TIPO_PASSO'] = '0. Independente'
            elif cod in saidas_liquidas and mov == 'SAIDA':
                df_mes.loc[idx, 'TIPO_PASSO'] = '1. Início'
            elif cod in entradas_liquidas and mov == 'ENTRADA':
                df_mes.loc[idx, 'TIPO_PASSO'] = '3. Fim'
            elif cod in codigos_intermediarios:
                df_mes.loc[idx, 'TIPO_PASSO'] = '2. Intermediário'

        msg = ''
        gravidade = 'OK'

        # Tratamento especial para códigos consolidadores corretos
        if len(saidas_liquidas) == 0 and len(entradas_liquidas) == 0:
            # Se só tem códigos consolidadores e eles estão corretos
            codigos_consolidadores = {31000, 32000, 33000}
            tem_consolidador = bool(codigos_entrada_set & codigos_consolidadores)
            
            if tem_consolidador and not msg:  # Ainda não tem mensagem de erro
                msg = f"OK: Lançamento consolidador correto"
                gravidade = 'OK'
                stats['ok'] += 1

        # ERRO: Múltiplas saídas líquidas (após filtrar códigos de ruído)
        if len(saidas_liquidas) > 1:
            msg = f"ERRO: Participante tem múltiplas saídas finais no mesmo mês ({', '.join(map(str, saidas_liquidas))}). Isso é muito raro e pode indicar problema no sistema."
            gravidade = 'ERRO'
            stats['erros'] += 1

        elif len(entradas_liquidas_filtradas) > 1:
            if saidas_liquidas == {31200} and entradas_liquidas_filtradas == {21000, 31300}:
                msg = f"OK: Transição aceita 31200 → (21000 + 31300)"
                gravidade = 'OK'
                stats['ok'] += 1
            else:
                msg = f"ERRO: Múltiplas entradas finais"
                gravidade = 'ERRO'
                stats['erros'] += 1

        elif len(saidas_liquidas) == 1 and len(entradas_liquidas_filtradas) == 1:
            cod_origem = list(saidas_liquidas)[0]
            cod_destino = list(entradas_liquidas_filtradas)[0]

            if (cod_origem, cod_destino) in regras_validas:
                if cod_origem == 21000 and cod_destino in {31100, 31200, 31300, 22000}:
                    msg = f"ERRO: BPD não pode retornar para Ativo"
                    gravidade = 'ERRO'
                    stats['erros'] += 1
                else:
                    msg = f"OK: Transição válida {get_descricao(cod_origem, df_codigos)} → {get_descricao(cod_destino, df_codigos)}"
                    gravidade = 'OK'
                    stats['ok'] += 1
            else:
                msg = f"ERRO: Transição NÃO PERMITIDA {cod_origem} → {cod_destino}"
                gravidade = 'ERRO'
                stats['erros'] += 1

        elif len(saidas_liquidas) == 0 and len(entradas_liquidas_filtradas) > 0:
            cod_entrada = list(entradas_liquidas_filtradas)[0]
            
            # Verifica se há saída de autopatrocinado (22000) com consolidador (31300)
            if 22000 in codigos_saida_set and 31300 in codigos_saida_set and cod_entrada in {11100, 11200}:
                if 32000 in codigos_entrada_set:
                    msg = f"OK: Transição válida Autopatrocinado → Aposentadoria com consolidadores corretos"
                    gravidade = 'OK'
                    stats['ok'] += 1
                else:
                    msg = f"INFO: Processo em andamento"
                    gravidade = 'INFO'
                    stats['info'] += 1
            else:
                plano = group['PLANO'].iloc[0] if 'PLANO' in group.columns else None
                if plano == 5 and cod_entrada in constantes['CODIGOS_ADMISSAO']:
                    msg = f"INFO: Nova admissão no Plano 5"
                    gravidade = 'INFO'
                    stats['info'] += 1
                else:
                    msg = f"INFO: Processo em andamento"
                    gravidade = 'INFO'
                    stats['info'] += 1


        elif len(saidas_liquidas) == 0 and len(entradas_liquidas_filtradas) == 0 and len(entradas_independentes_liquidas) > 0:
            msg = f"OK: Lançamento(s) independente(s) ({', '.join(map(str, sorted(entradas_independentes_liquidas)))})"
            gravidade = 'OK'
            stats['ok'] += 1


        elif len(saidas_liquidas) > 0 and len(entradas_liquidas_filtradas) == 0:
            # Verifica se há códigos consolidadores correspondentes corretos
            tem_consolidador_correto = False
            
            # Para saídas de aposentadoria (11100, 11200) com 32000
            if saidas_liquidas & {11100, 11200} and 32000 in codigos_saida_set:
                tem_consolidador_correto = True
            
            # Para saída de pensão (14000) com 33000
            if 14000 in saidas_liquidas and 33000 in codigos_saida_set:
                tem_consolidador_correto = True
            
            if tem_consolidador_correto:
                msg = f"OK: Saída correta com lançamento consolidador correspondente"
                gravidade = 'OK'
                stats['ok'] += 1
            elif 22000 in saidas_liquidas:
                msg = f"ERRO: Saída de autopatrocinado (22000) sem entrada em nova situação"
                gravidade = 'ERRO'
                stats['erros'] += 1
            else:
                msg = f"INFO: Processo em andamento (aguardando conclusão)"
                gravidade = 'INFO'
                stats['info'] += 1

        if msg:
            df_mes.loc[group.index, 'ANALISE'] = msg
            df_mes.loc[group.index, 'GRAVIDADE'] = gravidade

    return df_mes, stats


def calcular_stats_participantes(df_res):
    if df_res is None or df_res.empty:
        return {'total': 0, 'erros': 0, 'info': 0, 'ok': 0}

    def pior_gravidade(series):
        vals = set(series.dropna().astype(str).tolist())
        if 'ERRO' in vals:
            return 'ERRO'
        if 'INFO' in vals:
            return 'INFO'
        return 'OK'

    por_participante = df_res.groupby('CODIGO ORGANIZACAO NOME')['GRAVIDADE'].apply(pior_gravidade)
    total = int(por_participante.shape[0])
    counts = por_participante.value_counts()

    return {
        'total': total,
        'erros': int(counts.get('ERRO', 0)),
        'info': int(counts.get('INFO', 0)),
        'ok': int(counts.get('OK', 0))
    }


def analisar_movimentacoes_periodo(df_mov, df_codigos, regras_validas, constantes, meses):
    dfs = []
    for mes in meses:
        df_mes, _stats_mes = analisar_movimentacoes_mes(
            df_mov,
            df_codigos,
            regras_validas,
            constantes,
            mes_analise=mes
        )
        if df_mes is not None and not df_mes.empty:
            dfs.append(df_mes)

    if not dfs:
        return pd.DataFrame(), {'total': 0, 'erros': 0, 'info': 0, 'ok': 0}

    df_res = pd.concat(dfs, ignore_index=True)
    stats = calcular_stats_participantes(df_res)
    return df_res, stats


def _mpl_fig_to_png_bytes(fig):
    import matplotlib.pyplot as plt
    bio = io.BytesIO()
    fig.savefig(
        bio,
        format='png',
        dpi=300,
        bbox_inches='tight',
        facecolor='white',
        edgecolor='white'
    )
    plt.close(fig)
    bio.seek(0)
    return bio.getvalue()


def gerar_pdf_relatorio_simples(titulo, subtitulo, kpis, df_res):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
        from reportlab.lib.units import cm
    except Exception as e:
        raise RuntimeError(f"Dependência ausente para PDF: {e}")

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    y = height - 2 * cm
    c.setFont("Helvetica-Bold", 14)
    c.drawString(2 * cm, y, str(titulo))
    y -= 0.8 * cm
    c.setFont("Helvetica", 10)
    c.drawString(2 * cm, y, str(subtitulo))
    y -= 1.2 * cm

    if kpis:
        c.setFont("Helvetica-Bold", 11)
        c.drawString(2 * cm, y, "Resumo")
        y -= 0.6 * cm
        c.setFont("Helvetica", 9)
        for k, v in kpis.items():
            c.drawString(2 * cm, y, f"{k}: {v}")
            y -= 0.45 * cm
        y -= 0.4 * cm

    if df_res is not None and not df_res.empty:
        c.setFont("Helvetica-Bold", 11)
        c.drawString(2 * cm, y, "Totais")
        y -= 0.6 * cm
        c.setFont("Helvetica", 9)
        try:
            total_registros = int(len(df_res))
            total_pessoas = int(df_res['CODIGO ORGANIZACAO NOME'].nunique()) if 'CODIGO ORGANIZACAO NOME' in df_res.columns else 0
            c.drawString(2 * cm, y, f"Registros: {total_registros}")
            y -= 0.45 * cm
            c.drawString(2 * cm, y, f"Participantes: {total_pessoas}")
            y -= 0.45 * cm

            if 'GRAVIDADE' in df_res.columns:
                counts = df_res['GRAVIDADE'].value_counts()
                c.drawString(2 * cm, y, f"OK: {int(counts.get('OK', 0))}")
                y -= 0.45 * cm
                c.drawString(2 * cm, y, f"INFO: {int(counts.get('INFO', 0))}")
                y -= 0.45 * cm
                c.drawString(2 * cm, y, f"ERRO: {int(counts.get('ERRO', 0))}")
                y -= 0.45 * cm
        except Exception:
            pass

    c.save()
    buffer.seek(0)
    return buffer.getvalue()


def gerar_pdf_relatorio_visual(titulo, subtitulo, kpis, df_res, df_codigos):
    import matplotlib.pyplot as plt
    import matplotlib as mpl
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import cm
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_LEFT
    except Exception as e:
        raise RuntimeError(f"Dependência ausente para PDF: {e}")

    if df_res is None or df_res.empty:
        raise RuntimeError("Não há dados para gerar o PDF")

    width, height = A4
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2.3 * cm,
        bottomMargin=1.8 * cm,
        title=str(titulo),
        author="ArcelorMittal",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        name="TituloRelatorio",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=20,
        textColor=colors.HexColor("#0B1F3A"),
        spaceAfter=4,
        alignment=TA_LEFT,
    )
    subtitle_style = ParagraphStyle(
        name="SubtituloRelatorio",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=13,
        textColor=colors.HexColor("#4B5563"),
        spaceAfter=12,
    )
    section_style = ParagraphStyle(
        name="SecaoRelatorio",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=14,
        textColor=colors.HexColor("#0B1F3A"),
        spaceBefore=10,
        spaceAfter=6,
    )
    small_style = ParagraphStyle(
        name="TextoPequeno",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#111827"),
    )
    kpi_style = ParagraphStyle(
        name="Kpi",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#111827"),
        spaceBefore=0,
        spaceAfter=0,
    )

    def desenhar_cabecalho_rodape(canvas, _doc):
        canvas.saveState()
        canvas.setFillColor(colors.HexColor("#0B1F3A"))
        canvas.rect(0, height - 1.1 * cm, width, 1.1 * cm, fill=1, stroke=0)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 10)
        canvas.drawString(2 * cm, height - 0.75 * cm, "ArcelorMittal | Relatório Estatístico")
        canvas.setFillColor(colors.HexColor("#6B7280"))
        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(width - 2 * cm, 1.1 * cm, f"Página {_doc.page}")
        canvas.restoreState()

    def _aplicar_estilo_mpl():
        try:
            plt.style.use('seaborn-v0_8-whitegrid')
        except Exception:
            try:
                plt.style.use('ggplot')
            except Exception:
                pass

        mpl.rcParams.update({
            'figure.facecolor': 'white',
            'axes.facecolor': 'white',
            'savefig.facecolor': 'white',
            'axes.edgecolor': '#E5E7EB',
            'axes.labelcolor': '#111827',
            'text.color': '#111827',
            'xtick.color': '#374151',
            'ytick.color': '#374151',
            'grid.color': '#E5E7EB',
            'grid.alpha': 0.6,
            'axes.titleweight': 'bold',
            'axes.titlesize': 14,
            'axes.labelsize': 11,
            'font.size': 11,
            'legend.frameon': False,
        })

    def _despine(ax):
        for s in ['top', 'right', 'left', 'bottom']:
            try:
                ax.spines[s].set_visible(False)
            except Exception:
                pass

    def _fmt_int(v):
        try:
            return f"{int(v):,}".replace(',', '.')
        except Exception:
            return str(v)

    def _tabela_estilizada(data, col_widths=None, header_rows=1):
        t = Table(data, colWidths=col_widths, hAlign='LEFT')
        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, header_rows - 1), colors.HexColor("#0B1F3A")),
            ("TEXTCOLOR", (0, 0), (-1, header_rows - 1), colors.white),
            ("FONTNAME", (0, 0), (-1, header_rows - 1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, header_rows - 1), 9),
            ("FONTSIZE", (0, header_rows), (-1, -1), 8.5),
            ("TEXTCOLOR", (0, header_rows), (-1, -1), colors.HexColor("#111827")),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#E5E7EB")),
            ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D1D5DB")),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]
        # Zebra
        for r in range(header_rows, len(data)):
            if (r - header_rows) % 2 == 0:
                style_cmds.append(("BACKGROUND", (0, r), (-1, r), colors.HexColor("#F9FAFB")))
            else:
                style_cmds.append(("BACKGROUND", (0, r), (-1, r), colors.white))
        t.setStyle(TableStyle(style_cmds))
        return t

    def add_mpl_fig(fig, caption=None, max_height_cm=11.0):
        img_bytes = _mpl_fig_to_png_bytes(fig)
        bio = io.BytesIO(img_bytes)
        img = Image(bio)
        target_w = doc.width
        target_h = min(max_height_cm * cm, target_w * 0.62)
        img.drawWidth = target_w
        img.drawHeight = target_h
        story.append(img)
        if caption:
            story.append(Spacer(1, 0.2 * cm))
            story.append(Paragraph(str(caption), subtitle_style))
        story.append(Spacer(1, 0.6 * cm))

    story = []
    story.append(Paragraph(str(titulo), title_style))
    story.append(Paragraph(str(subtitulo), subtitle_style))

    if kpis:
        story.append(Paragraph("Resumo", section_style))
        itens = list(kpis.items())
        ncols = 3
        linhas = []
        for i in range(0, len(itens), ncols):
            linha = []
            for j in range(ncols):
                if i + j < len(itens):
                    k, v = itens[i + j]
                    linha.append(Paragraph(f"<b>{k}</b><br/>{v}", kpi_style))
                else:
                    linha.append(Paragraph(" ", kpi_style))
            linhas.append(linha)

        t = Table(linhas, colWidths=[doc.width / ncols] * ncols)
        t.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F3F4F6")),
                    ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D1D5DB")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#E5E7EB")),
                    ("LEFTPADDING", (0, 0), (-1, -1), 10),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                    ("TOPPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ]
            )
        )
        story.append(t)
        story.append(Spacer(1, 0.8 * cm))

    total_registros = int(len(df_res))
    total_pessoas = int(df_res['CODIGO ORGANIZACAO NOME'].nunique()) if 'CODIGO ORGANIZACAO NOME' in df_res.columns else 0
    grav_counts = df_res['GRAVIDADE'].value_counts() if 'GRAVIDADE' in df_res.columns else pd.Series(dtype=int)
    story.append(Paragraph("Totais", section_style))
    story.append(
        Paragraph(
            f"Registros: <b>{total_registros}</b><br/>"
            f"Participantes: <b>{total_pessoas}</b><br/>"
            f"OK: <b>{int(grav_counts.get('OK', 0))}</b> | "
            f"INFO: <b>{int(grav_counts.get('INFO', 0))}</b> | "
            f"ERRO: <b>{int(grav_counts.get('ERRO', 0))}</b>",
            small_style,
        )
    )
    story.append(Spacer(1, 0.8 * cm))

    story.append(Paragraph("Visão Geral", section_style))
    _aplicar_estilo_mpl()

    grav = df_res.groupby('GRAVIDADE').size().reindex(['OK', 'INFO', 'ERRO']).fillna(0).astype(int)

    # Tabela resumo de gravidade
    total_geral = int(grav.sum()) if not grav.empty else 0
    tab_grav = [["Gravidade", "Quantidade", "%"]]
    for g in ['OK', 'INFO', 'ERRO']:
        qtd = int(grav.get(g, 0))
        perc = (qtd / total_geral * 100) if total_geral else 0
        tab_grav.append([g, _fmt_int(qtd), f"{perc:.1f}%"])
    story.append(_tabela_estilizada(tab_grav, col_widths=[doc.width * 0.34, doc.width * 0.33, doc.width * 0.33]))
    story.append(Spacer(1, 0.7 * cm))

    # Gráfico de gravidade (donut)
    fig, ax = plt.subplots(figsize=(7.6, 4.4))
    colors_grav = ['#22C55E', '#06B6D4', '#EF4444']
    wedges, texts, autotexts = ax.pie(
        grav.values,
        labels=grav.index.tolist(),
        autopct=lambda p: f"{p:.1f}%\n({_fmt_int(round(p/100*total_geral))})" if total_geral else "0%",
        startangle=90,
        colors=colors_grav,
        textprops={'color': '#111827', 'fontsize': 10},
        wedgeprops={'width': 0.42, 'edgecolor': 'white'}
    )
    ax.set_title('Distribuição de Gravidade')
    ax.axis('equal')
    add_mpl_fig(fig)

    if 'PLANO' in df_res.columns:
        plano_grav = df_res.groupby(['PLANO', 'GRAVIDADE']).size().unstack(fill_value=0)
        plano_grav = plano_grav.reindex(columns=['OK', 'INFO', 'ERRO'], fill_value=0)

        # Tabela por plano (Top 12)
        plano_tab = plano_grav.copy()
        plano_tab['TOTAL'] = plano_tab.sum(axis=1)
        plano_tab = plano_tab.sort_values('TOTAL', ascending=False).head(12)
        tab_plano = [["Plano", "OK", "INFO", "ERRO", "Total"]]
        for plano, row in plano_tab.iterrows():
            tab_plano.append([
                str(plano),
                _fmt_int(row.get('OK', 0)),
                _fmt_int(row.get('INFO', 0)),
                _fmt_int(row.get('ERRO', 0)),
                _fmt_int(row.get('TOTAL', 0)),
            ])
        story.append(_tabela_estilizada(
            tab_plano,
            col_widths=[doc.width * 0.34, doc.width * 0.165, doc.width * 0.165, doc.width * 0.165, doc.width * 0.165]
        ))
        story.append(Spacer(1, 0.7 * cm))

        fig, ax = plt.subplots(figsize=(9.2, 4.9))
        x = np.arange(len(plano_grav.index))
        w = 0.25
        ax.bar(x - w, plano_grav['OK'].values, width=w, label='OK', color='#22C55E')
        ax.bar(x, plano_grav['INFO'].values, width=w, label='INFO', color='#06B6D4')
        ax.bar(x + w, plano_grav['ERRO'].values, width=w, label='ERRO', color='#EF4444')
        ax.set_xticks(x)
        ax.set_xticklabels([str(p) for p in plano_grav.index], rotation=0)
        ax.set_title('Movimentações por Plano e Status')
        ax.set_xlabel('Plano')
        ax.set_ylabel('Quantidade')
        ax.legend(ncol=3, loc='upper right')
        ax.grid(axis='y', linestyle='-', linewidth=0.6)
        _despine(ax)
        add_mpl_fig(fig)

    story.append(PageBreak())
    story.append(Paragraph("Códigos e Benefícios", section_style))

    mov_por_codigo = df_res.groupby('CODIGO BENEFICIO').size().reset_index(name='count')
    mov_por_codigo = mov_por_codigo.merge(
        df_codigos[['CODIGO', 'DESCRICAO', 'TIPO']],
        left_on='CODIGO BENEFICIO',
        right_on='CODIGO',
        how='left'
    )
    total_mov = float(mov_por_codigo['count'].sum()) if not mov_por_codigo.empty else 0.0
    mov_por_codigo['percentual'] = (mov_por_codigo['count'] / total_mov * 100) if total_mov else 0.0
    top10 = mov_por_codigo.nlargest(10, 'count').sort_values('count', ascending=True)

    # Tabela Top 10
    tab_top10 = [["Código", "Descrição", "Tipo", "Qtd", "%"]]
    for _, r in top10.sort_values('count', ascending=False).iterrows():
        tab_top10.append([
            str(int(r['CODIGO BENEFICIO'])) if pd.notna(r.get('CODIGO BENEFICIO')) else "-",
            str(r.get('DESCRICAO') if pd.notna(r.get('DESCRICAO')) else "-")[:60],
            str(r.get('TIPO') if pd.notna(r.get('TIPO')) else "-"),
            _fmt_int(r.get('count', 0)),
            f"{float(r.get('percentual', 0.0)):.1f}%",
        ])
    story.append(_tabela_estilizada(
        tab_top10,
        col_widths=[doc.width * 0.13, doc.width * 0.50, doc.width * 0.17, doc.width * 0.10, doc.width * 0.10]
    ))
    story.append(Spacer(1, 0.7 * cm))

    fig, ax = plt.subplots(figsize=(9.0, 5.0))
    labels = top10['DESCRICAO'].fillna(top10['CODIGO BENEFICIO'].astype(str)).astype(str).map(lambda s: (s[:42] + '…') if len(s) > 43 else s)
    ax.barh(labels, top10['count'].values, color='#2563EB')
    ax.set_title('Top 10 Códigos Mais Utilizados')
    ax.set_xlabel('Quantidade')
    ax.grid(axis='x', linestyle='-', linewidth=0.6)
    _despine(ax)
    add_mpl_fig(fig)

    tipo_dist = mov_por_codigo.groupby('TIPO')['count'].sum().sort_values(ascending=True)
    if not tipo_dist.empty:
        tab_tipo = [["Tipo", "Quantidade", "%"]]
        total_tipo = float(tipo_dist.sum()) if float(tipo_dist.sum()) else 0.0
        for tipo, qtd in tipo_dist.sort_values(ascending=False).items():
            perc = (float(qtd) / total_tipo * 100) if total_tipo else 0
            tab_tipo.append([str(tipo), _fmt_int(qtd), f"{perc:.1f}%"])
        story.append(_tabela_estilizada(tab_tipo, col_widths=[doc.width * 0.50, doc.width * 0.25, doc.width * 0.25]))
        story.append(Spacer(1, 0.7 * cm))

        fig, ax = plt.subplots(figsize=(7.8, 4.4))
        ax.barh(tipo_dist.index.astype(str), tipo_dist.values, color='#F59E0B')
        ax.set_title('Distribuição por Tipo de Código')
        ax.set_xlabel('Quantidade')
        ax.grid(axis='x', linestyle='-', linewidth=0.6)
        _despine(ax)
        add_mpl_fig(fig)

    transicoes = df_res[df_res['MOVIMENTO'] == 'SAIDA'].merge(
        df_res[df_res['MOVIMENTO'] == 'ENTRADA'],
        on=['CODIGO ORGANIZACAO NOME', 'ANO MES'],
        suffixes=('_origem', '_destino')
    )
    if not transicoes.empty:
        story.append(PageBreak())
        story.append(Paragraph("Transições", section_style))
        trans_grouped = transicoes.groupby(['CODIGO BENEFICIO_origem', 'CODIGO BENEFICIO_destino']).size().reset_index(name='count')
        trans_grouped = trans_grouped.nlargest(15, 'count').sort_values('count', ascending=True)

        # Tabela de transições (Top 15)
        codigo_to_desc = df_codigos.set_index('CODIGO')['DESCRICAO'].to_dict() if df_codigos is not None and not df_codigos.empty else {}
        tab_trans = [["Origem", "Destino", "Qtd"]]
        for _, r in trans_grouped.sort_values('count', ascending=False).iterrows():
            o = int(r['CODIGO BENEFICIO_origem'])
            d = int(r['CODIGO BENEFICIO_destino'])
            o_desc = str(codigo_to_desc.get(o, ""))
            d_desc = str(codigo_to_desc.get(d, ""))
            origem = f"{o} - {o_desc[:35]}".strip(" -")
            destino = f"{d} - {d_desc[:35]}".strip(" -")
            tab_trans.append([origem, destino, _fmt_int(r['count'])])
        story.append(_tabela_estilizada(
            tab_trans,
            col_widths=[doc.width * 0.44, doc.width * 0.44, doc.width * 0.12]
        ))
        story.append(Spacer(1, 0.7 * cm))

        labels = trans_grouped.apply(
            lambda row: f"{int(row['CODIGO BENEFICIO_origem'])}→{int(row['CODIGO BENEFICIO_destino'])}",
            axis=1
        )
        fig, ax = plt.subplots(figsize=(9.2, 6.0))
        ax.barh(labels.tolist(), trans_grouped['count'].values, color='#10B981')
        ax.set_title('Top 15 Transições Mais Frequentes')
        ax.set_xlabel('Quantidade')
        ax.grid(axis='x', linestyle='-', linewidth=0.6)
        _despine(ax)
        add_mpl_fig(fig, max_height_cm=12.5)

        top_codes = pd.concat([
            transicoes['CODIGO BENEFICIO_origem'],
            transicoes['CODIGO BENEFICIO_destino']
        ]).value_counts().head(20).index.tolist()
        matriz = transicoes.groupby(['CODIGO BENEFICIO_origem', 'CODIGO BENEFICIO_destino']).size().reset_index(name='count')
        matriz = matriz[matriz['CODIGO BENEFICIO_origem'].isin(top_codes) & matriz['CODIGO BENEFICIO_destino'].isin(top_codes)]
        if not matriz.empty:
            pivot = matriz.pivot(index='CODIGO BENEFICIO_origem', columns='CODIGO BENEFICIO_destino', values='count').fillna(0)
            pivot = pivot.reindex(index=sorted(top_codes), columns=sorted(top_codes), fill_value=0)
            fig, ax = plt.subplots(figsize=(10.0, 7.0))
            im = ax.imshow(pivot.values, aspect='auto', cmap='RdYlGn')
            ax.set_title('Heatmap de Transições (Top 20 códigos)')
            ax.set_xlabel('Destino')
            ax.set_ylabel('Origem')
            ax.set_xticks(np.arange(len(pivot.columns)))
            ax.set_yticks(np.arange(len(pivot.index)))
            ax.set_xticklabels([str(int(c)) for c in pivot.columns], rotation=90)
            ax.set_yticklabels([str(int(i)) for i in pivot.index])
            fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
            add_mpl_fig(fig, max_height_cm=13.5)

    if 'GRAVIDADE' in df_res.columns and (df_res['GRAVIDADE'] == 'ERRO').any():
        erros_df = df_res[df_res['GRAVIDADE'] == 'ERRO'].copy()
        erros_df['TIPO_ERRO'] = erros_df['ANALISE'].astype(str).str.extract(r'ERRO: ([^.]+)')[0]
        tipo_erro_counts = erros_df.groupby('TIPO_ERRO').size().sort_values(ascending=True).tail(10)
        story.append(PageBreak())
        story.append(Paragraph("Erros", section_style))

        # Tabela tipos de erro
        if not tipo_erro_counts.empty:
            tab_erro_tipo = [["Tipo de Erro", "Qtd"]]
            for tipo, qtd in tipo_erro_counts.sort_values(ascending=False).items():
                tab_erro_tipo.append([str(tipo)[:70], _fmt_int(qtd)])
            story.append(_tabela_estilizada(tab_erro_tipo, col_widths=[doc.width * 0.82, doc.width * 0.18]))
            story.append(Spacer(1, 0.7 * cm))

        if not tipo_erro_counts.empty:
            fig, ax = plt.subplots(figsize=(9.0, 5.2))
            ax.barh(tipo_erro_counts.index.fillna('Desconhecido').astype(str), tipo_erro_counts.values, color='#EF4444')
            ax.set_title('Top 10 Tipos de Erro')
            ax.set_xlabel('Quantidade')
            ax.grid(axis='x', linestyle='-', linewidth=0.6)
            _despine(ax)
            add_mpl_fig(fig)

        if 'PLANO' in erros_df.columns:
            erros_plano = erros_df.groupby('PLANO').size()
            if not erros_plano.empty:
                tab_erros_plano = [["Plano", "Qtd"]]
                for plano, qtd in erros_plano.sort_values(ascending=False).head(12).items():
                    tab_erros_plano.append([str(plano), _fmt_int(qtd)])
                story.append(_tabela_estilizada(tab_erros_plano, col_widths=[doc.width * 0.78, doc.width * 0.22]))
                story.append(Spacer(1, 0.7 * cm))

                fig, ax = plt.subplots(figsize=(7.2, 4.4))
                ax.pie(erros_plano.values, labels=[str(p) for p in erros_plano.index], autopct='%1.1f%%', startangle=90)
                ax.set_title('Erros por Plano')
                add_mpl_fig(fig)

        cod_erro = erros_df.groupby('CODIGO BENEFICIO').size().reset_index(name='erros')
        cod_erro = cod_erro.merge(df_codigos[['CODIGO', 'DESCRICAO']], left_on='CODIGO BENEFICIO', right_on='CODIGO', how='left')
        cod_erro = cod_erro.sort_values('erros', ascending=False).head(10).sort_values('erros', ascending=True)
        if not cod_erro.empty:
            tab_cod_erro = [["Código", "Descrição", "Qtd"]]
            for _, r in cod_erro.sort_values('erros', ascending=False).iterrows():
                tab_cod_erro.append([
                    str(int(r['CODIGO BENEFICIO'])) if pd.notna(r.get('CODIGO BENEFICIO')) else "-",
                    str(r.get('DESCRICAO') if pd.notna(r.get('DESCRICAO')) else "-")[:70],
                    _fmt_int(r.get('erros', 0)),
                ])
            story.append(_tabela_estilizada(
                tab_cod_erro,
                col_widths=[doc.width * 0.15, doc.width * 0.67, doc.width * 0.18]
            ))
            story.append(Spacer(1, 0.7 * cm))

            fig, ax = plt.subplots(figsize=(9.0, 5.2))
            labels = cod_erro['DESCRICAO'].fillna(cod_erro['CODIGO BENEFICIO'].astype(str)).astype(str)
            ax.barh(labels, cod_erro['erros'].values, color='#991B1B')
            ax.set_title('Top 10 Códigos com Mais Erros')
            ax.set_xlabel('Quantidade')
            ax.grid(axis='x', linestyle='-', linewidth=0.6)
            _despine(ax)
            add_mpl_fig(fig)

    doc.build(story, onFirstPage=desenhar_cabecalho_rodape, onLaterPages=desenhar_cabecalho_rodape)
    buffer.seek(0)
    return buffer.getvalue()


def gerar_pdf_relatorio_sem_kaleido(titulo, subtitulo, kpis, df_res, df_codigos):
    import matplotlib.pyplot as plt
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
        from reportlab.lib.units import cm
        from reportlab.lib.utils import ImageReader
    except Exception as e:
        raise RuntimeError(f"Dependência ausente para PDF: {e}")

    if df_res is None or df_res.empty:
        raise RuntimeError("Não há dados para gerar o PDF")

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    y = height - 2 * cm
    c.setFont("Helvetica-Bold", 14)
    c.drawString(2 * cm, y, str(titulo))
    y -= 0.8 * cm
    c.setFont("Helvetica", 10)
    c.drawString(2 * cm, y, str(subtitulo))
    y -= 1.2 * cm

    if kpis:
        c.setFont("Helvetica-Bold", 11)
        c.drawString(2 * cm, y, "Resumo")
        y -= 0.6 * cm
        c.setFont("Helvetica", 9)
        for k, v in kpis.items():
            c.drawString(2 * cm, y, f"{k}: {v}")
            y -= 0.45 * cm
        y -= 0.4 * cm

    imagens = []

    # 1) Distribuição por gravidade
    grav = df_res.groupby('GRAVIDADE').size().reindex(['OK', 'INFO', 'ERRO']).fillna(0).astype(int)
    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    colors = ['#28a745', '#17a2b8', '#dc3545']
    ax.pie(grav.values, labels=grav.index.tolist(), autopct='%1.1f%%', startangle=90, colors=colors)
    ax.set_title('Distribuição de Gravidade')
    imagens.append(_mpl_fig_to_png_bytes(fig))

    # 2) Análise por plano (se existir)
    if 'PLANO' in df_res.columns:
        plano_grav = df_res.groupby(['PLANO', 'GRAVIDADE']).size().unstack(fill_value=0)
        plano_grav = plano_grav.reindex(columns=['OK', 'INFO', 'ERRO'], fill_value=0)
        fig, ax = plt.subplots(figsize=(8.5, 4.5))
        x = np.arange(len(plano_grav.index))
        w = 0.25
        ax.bar(x - w, plano_grav['OK'].values, width=w, label='OK', color='#28a745')
        ax.bar(x, plano_grav['INFO'].values, width=w, label='INFO', color='#17a2b8')
        ax.bar(x + w, plano_grav['ERRO'].values, width=w, label='ERRO', color='#dc3545')
        ax.set_xticks(x)
        ax.set_xticklabels([str(p) for p in plano_grav.index])
        ax.set_title('Movimentações por Plano e Status')
        ax.set_xlabel('Plano')
        ax.set_ylabel('Quantidade')
        ax.legend()
        imagens.append(_mpl_fig_to_png_bytes(fig))

    # 3) Top 10 códigos mais utilizados
    mov_por_codigo = df_res.groupby('CODIGO BENEFICIO').size().reset_index(name='count')
    mov_por_codigo = mov_por_codigo.merge(
        df_codigos[['CODIGO', 'DESCRICAO', 'TIPO']],
        left_on='CODIGO BENEFICIO',
        right_on='CODIGO',
        how='left'
    )
    top10 = mov_por_codigo.nlargest(10, 'count').sort_values('count', ascending=True)
    fig, ax = plt.subplots(figsize=(9.0, 5.0))
    ax.barh(top10['DESCRICAO'].fillna(top10['CODIGO BENEFICIO'].astype(str)).astype(str), top10['count'].values, color='#1f77b4')
    ax.set_title('Top 10 Códigos Mais Utilizados')
    ax.set_xlabel('Quantidade')
    imagens.append(_mpl_fig_to_png_bytes(fig))

    # 4) Distribuição por tipo
    tipo_dist = mov_por_codigo.groupby('TIPO')['count'].sum().sort_values(ascending=True)
    if not tipo_dist.empty:
        fig, ax = plt.subplots(figsize=(7.5, 4.2))
        ax.barh(tipo_dist.index.astype(str), tipo_dist.values, color='#ff7f0e')
        ax.set_title('Distribuição por Tipo de Código')
        ax.set_xlabel('Quantidade')
        imagens.append(_mpl_fig_to_png_bytes(fig))

    # 5) Transições (top 15)
    transicoes = df_res[df_res['MOVIMENTO'] == 'SAIDA'].merge(
        df_res[df_res['MOVIMENTO'] == 'ENTRADA'],
        on=['CODIGO ORGANIZACAO NOME', 'ANO MES'],
        suffixes=('_origem', '_destino')
    )
    if not transicoes.empty:
        trans_grouped = transicoes.groupby(['CODIGO BENEFICIO_origem', 'CODIGO BENEFICIO_destino']).size().reset_index(name='count')
        trans_grouped = trans_grouped.nlargest(15, 'count').sort_values('count', ascending=True)

        codigo_to_desc = df_codigos.set_index('CODIGO')['DESCRICAO'].to_dict()
        labels = trans_grouped.apply(
            lambda row: f"{int(row['CODIGO BENEFICIO_origem'])}→{int(row['CODIGO BENEFICIO_destino'])}",
            axis=1
        )

        fig, ax = plt.subplots(figsize=(9.5, 6.0))
        ax.barh(labels.tolist(), trans_grouped['count'].values, color='#2ca02c')
        ax.set_title('Top 15 Transições Mais Frequentes')
        ax.set_xlabel('Quantidade')
        imagens.append(_mpl_fig_to_png_bytes(fig))

        # 6) Heatmap de transições (top 20 códigos para legibilidade)
        top_codes = pd.concat([
            transicoes['CODIGO BENEFICIO_origem'],
            transicoes['CODIGO BENEFICIO_destino']
        ]).value_counts().head(20).index.tolist()

        matriz = transicoes.groupby(['CODIGO BENEFICIO_origem', 'CODIGO BENEFICIO_destino']).size().reset_index(name='count')
        matriz = matriz[matriz['CODIGO BENEFICIO_origem'].isin(top_codes) & matriz['CODIGO BENEFICIO_destino'].isin(top_codes)]
        if not matriz.empty:
            pivot = matriz.pivot(index='CODIGO BENEFICIO_origem', columns='CODIGO BENEFICIO_destino', values='count').fillna(0)
            pivot = pivot.reindex(index=sorted(top_codes), columns=sorted(top_codes), fill_value=0)

            fig, ax = plt.subplots(figsize=(10.0, 7.0))
            im = ax.imshow(pivot.values, aspect='auto', cmap='RdYlGn')
            ax.set_title('Heatmap de Transições (Top 20 códigos)')
            ax.set_xlabel('Destino')
            ax.set_ylabel('Origem')
            ax.set_xticks(np.arange(len(pivot.columns)))
            ax.set_yticks(np.arange(len(pivot.index)))
            ax.set_xticklabels([str(int(c)) for c in pivot.columns], rotation=90)
            ax.set_yticklabels([str(int(i)) for i in pivot.index])
            fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
            imagens.append(_mpl_fig_to_png_bytes(fig))

    # 7) Erros (se houver)
    if 'GRAVIDADE' in df_res.columns and (df_res['GRAVIDADE'] == 'ERRO').any():
        erros_df = df_res[df_res['GRAVIDADE'] == 'ERRO'].copy()
        erros_df['TIPO_ERRO'] = erros_df['ANALISE'].astype(str).str.extract(r'ERRO: ([^.]+)')[0]
        tipo_erro_counts = erros_df.groupby('TIPO_ERRO').size().sort_values(ascending=True).tail(10)
        if not tipo_erro_counts.empty:
            fig, ax = plt.subplots(figsize=(9.0, 5.0))
            ax.barh(tipo_erro_counts.index.fillna('Desconhecido').astype(str), tipo_erro_counts.values, color='#dc3545')
            ax.set_title('Top 10 Tipos de Erro')
            ax.set_xlabel('Quantidade')
            imagens.append(_mpl_fig_to_png_bytes(fig))

        if 'PLANO' in erros_df.columns:
            erros_plano = erros_df.groupby('PLANO').size()
            if not erros_plano.empty:
                fig, ax = plt.subplots(figsize=(7.0, 4.2))
                ax.pie(erros_plano.values, labels=[str(p) for p in erros_plano.index], autopct='%1.1f%%', startangle=90)
                ax.set_title('Erros por Plano')
                imagens.append(_mpl_fig_to_png_bytes(fig))

        cod_erro = erros_df.groupby('CODIGO BENEFICIO').size().reset_index(name='erros')
        cod_erro = cod_erro.merge(df_codigos[['CODIGO', 'DESCRICAO']], left_on='CODIGO BENEFICIO', right_on='CODIGO', how='left')
        cod_erro = cod_erro.sort_values('erros', ascending=False).head(10).sort_values('erros', ascending=True)
        if not cod_erro.empty:
            fig, ax = plt.subplots(figsize=(9.0, 5.0))
            labels = cod_erro['DESCRICAO'].fillna(cod_erro['CODIGO BENEFICIO'].astype(str)).astype(str)
            ax.barh(labels, cod_erro['erros'].values, color='crimson')
            ax.set_title('Top 10 Códigos com Mais Erros')
            ax.set_xlabel('Quantidade')
            imagens.append(_mpl_fig_to_png_bytes(fig))

    # Renderiza imagens no PDF
    for img_bytes in imagens:
        img = ImageReader(io.BytesIO(img_bytes))
        img_w = width - 4 * cm
        img_h = img_w * (700 / 1200)

        if y - img_h < 2 * cm:
            c.showPage()
            y = height - 2 * cm

        c.drawImage(img, 2 * cm, y - img_h, width=img_w, height=img_h, preserveAspectRatio=True, anchor='c')
        y -= img_h + 1 * cm

    c.save()
    buffer.seek(0)
    return buffer.getvalue()


def gerar_pdf_relatorio(titulo, subtitulo, kpis, figuras):
    raise RuntimeError("Exportação via Kaleido desativada. Use gerar_pdf_relatorio_simples().")

# ============================================================================
# INTERFACE PRINCIPAL
# ============================================================================


def main():
    st.markdown('<div class="main-header">📊 Sistema de Análise de Movimentações Previdenciárias<br>ArcelorMittal</div>', unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.image("https://companieslogo.com/img/orig/MT_BIG.D-48309f61.png?t=1741059352",
                 use_container_width=True, width=150, caption="Projeto PrevicAuto - V0.1", clamp=True, output_format="PNG", channels="RGB")
        st.markdown("### ⚙️ Configurações")

        modo = st.radio(
            "Modo de Operação:",
            ["📁 Upload de Arquivo"],
            help="Outros modos de operação serão implementados futuramente."
        )

        st.markdown("---")
        st.markdown("### 📖 Sobre")
        st.info("""
        Sistema automatizado para validação de movimentações previdenciárias.
        
        **Recursos:**
        - ✅ Validação de transições
        - 📊 Estatísticas detalhadas
        - 📈 Visualizações interativas
        - 📥 Export para Excel
        """)

    # Carrega base de conhecimento
    df_codigos, regras_validas, constantes = carregar_base_conhecimento()

    # Tabs principais
    tab1, tab2, tab3, tab4 = st.tabs(
        ["🏠 Análise", "📊 Estatísticas", "🔍 Busca", "📚 Documentação"])

    with tab1:
        st.markdown("## 📂 Importação de Dados")

        df_para_analise = None

        if modo == "📁 Upload de Arquivo":
            uploaded_file = st.file_uploader(
                "Selecione o arquivo Excel ou CSV",
                type=['xlsx', 'csv'],
                help="Arquivo deve conter as colunas: CODIGO ORGANIZACAO PESSOA, NOME, PLANO, CODIGO BENEFICIO, ANO MES, MOVIMENTO"
            )

            if uploaded_file is not None:
                try:
                    with st.spinner('🔄 Processando arquivo...'):
                        if uploaded_file.name.endswith('.xlsx'):
                            df_bruto = pd.read_excel(uploaded_file)
                        else:
                            df_bruto = pd.read_csv(
                                uploaded_file, sep=';', on_bad_lines='skip')

                        # Limpeza e preparação
                        df_bruto.columns = df_bruto.columns.str.strip()

                        column_mapping = {
                            'CODIGO ORGANIZACAO PESSOA': 'CODIGO_ORG',
                            'NOME': 'NOME',
                            'PLANO': 'PLANO',
                            'CODIGO BENEFICIO': 'CODIGO BENEFICIO',
                            'ANO MES': 'ANO MES',
                            'MOVIMENTO': 'MOVIMENTO'
                        }

                        df_bruto.rename(columns=column_mapping, inplace=True)
                        df_para_analise = df_bruto.copy()

                        # Conversões
                        df_para_analise['CODIGO BENEFICIO'] = pd.to_numeric(
                            df_para_analise['CODIGO BENEFICIO'], errors='coerce')
                        df_para_analise['ANO MES'] = pd.to_numeric(
                            df_para_analise['ANO MES'], errors='coerce')

                        # Remove NAs
                        df_para_analise.dropna(
                            subset=['CODIGO BENEFICIO', 'ANO MES', 'CODIGO_ORG', 'NOME'], inplace=True)

                        # Conversões finais
                        df_para_analise['CODIGO BENEFICIO'] = df_para_analise['CODIGO BENEFICIO'].astype(
                            int)
                        df_para_analise['ANO MES'] = df_para_analise['ANO MES'].astype(
                            int)

                        # Remove códigos ignorados
                        # df_para_analise = df_para_analise[~df_para_analise['CODIGO BENEFICIO'].isin(
                        #     constantes['CODIGOS_IGNORAR'])].copy()

                        # Remove duplicatas
                        df_para_analise = df_para_analise.sort_values('ANO MES').drop_duplicates(
                            subset=['CODIGO_ORG', 'NOME',
                                    'CODIGO BENEFICIO', 'MOVIMENTO'],
                            keep='first'
                        )

                        # Cria identificador
                        df_para_analise['CODIGO ORGANIZACAO NOME'] = (
                            df_para_analise['CODIGO_ORG'].astype(str) + " - " +
                            df_para_analise['NOME'].apply(
                                formatar_nome_participante)
                        )

                        st.success(
                            f"✅ Arquivo carregado: {len(df_para_analise)} registros válidos")

                        st.session_state['df_dados'] = df_para_analise
                        st.session_state.pop('df_resultado_geral', None)
                        st.session_state.pop('stats_geral', None)
                        st.session_state.pop('pdf_bytes', None)

                except Exception as e:
                    st.error(f"❌ Erro ao processar arquivo: {e}")

        else:  # Modo teste
            st.info("🧪 **Modo de Teste Ativado** - Dados simulados serão gerados")

            col1, col2 = st.columns(2)
            with col1:
                n_participantes = st.slider(
                    "Número de participantes:", 50, 500, 200)
            with col2:
                mes_teste = st.number_input(
                    "Mês de análise:", 202401, 202512, 202501)

            if st.button("🎲 Gerar Dados de Teste", type="primary"):
                with st.spinner('🔄 Gerando dados...'):
                    # Importa gerador (simplificado aqui)
                    from datetime import datetime
                    import random

                    random.seed(42)
                    dados = []

                    for i in range(n_participantes):
                        codigo_org = 50000000 + i
                        nome = f"Participante Teste {i+1}"
                        plano = random.choice([3, 4, 5, 6, 7])

                        # Transição simples
                        origem = random.choice([31100, 31200])
                        destino = random.choice([11100, 21000, 22000])

                        dados.append({
                            'CODIGO_ORG': codigo_org,
                            'NOME': nome,
                            'PLANO': plano,
                            'ANO MES': mes_teste,
                            'CODIGO BENEFICIO': origem,
                            'MOVIMENTO': 'SAIDA'
                        })

                        dados.append({
                            'CODIGO_ORG': codigo_org,
                            'NOME': nome,
                            'PLANO': plano,
                            'ANO MES': mes_teste,
                            'CODIGO BENEFICIO': destino,
                            'MOVIMENTO': 'ENTRADA'
                        })

                    df_para_analise = pd.DataFrame(dados)
                    df_para_analise['CODIGO ORGANIZACAO NOME'] = (
                        df_para_analise['CODIGO_ORG'].astype(
                            str) + " - " + df_para_analise['NOME']
                    )

                    st.success(
                        f"✅ {len(df_para_analise)} registros de teste gerados")

        # ANÁLISE
        if df_para_analise is not None and not df_para_analise.empty:
            st.markdown("---")
            st.markdown("## 🔬 Análise de Movimentações")

            meses_disponiveis = sorted(df_para_analise['ANO MES'].unique())
            mes_selecionado = st.selectbox(
                "Selecione o mês para análise:",
                meses_disponiveis,
                index=len(meses_disponiveis)-1
            )

            if st.button("▶️ Executar Análise", type="primary", use_container_width=True):
                with st.spinner('🔄 Analisando movimentações...'):
                    df_resultado, stats = analisar_movimentacoes_mes(
                        df_para_analise,
                        df_codigos,
                        regras_validas,
                        constantes,
                        mes_analise=mes_selecionado
                    )

                    # Salva no session state
                    st.session_state['df_resultado'] = df_resultado
                    st.session_state['stats'] = stats
                    st.session_state['mes_analisado'] = mes_selecionado
                    st.session_state['df_dados'] = df_para_analise
                    st.session_state.pop('pdf_bytes', None)

                    st.success("✅ Análise concluída!")

                    # Métricas
                    st.markdown("### 📊 Resultados Gerais")
                    col1, col2, col3, col4 = st.columns(4)

                    with col1:
                        st.metric("👥 Total", stats['total'])
                    with col2:
                        st.metric("✅ OK", stats['ok'], delta_color="normal")
                    with col3:
                        st.metric("ℹ️ Info", stats['info'], delta_color="off")
                    with col4:
                        st.metric(
                            "❌ Erros", stats['erros'], delta_color="inverse")

                    # Gráfico de pizza
                    fig = go.Figure(data=[go.Pie(
                        labels=['OK', 'INFO', 'ERRO'],
                        values=[stats['ok'], stats['info'], stats['erros']],
                        marker_colors=['#44ff44', '#4488ff', '#ff4444'],
                        hole=0.4
                    )])
                    fig.update_layout(
                        title="Distribuição das Análises", height=400)
                    st.plotly_chart(fig, use_container_width=True)

                    # Tabela de erros
                    if stats['erros'] > 0:
                        st.markdown("### ❌ Erros Encontrados")
                        erros = df_resultado[df_resultado['GRAVIDADE'] == 'ERRO']
                        st.dataframe(
                            erros[['CODIGO ORGANIZACAO NOME', 'PLANO',
                                   'CODIGO BENEFICIO', 'MOVIMENTO', 'ANALISE']],
                            use_container_width=True,
                            height=400
                        )

                        st.markdown("### 📥 Baixar Resultados da Análise")

                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            # Aba com toda a análise
                            df_resultado.to_excel(
                                writer, index=False, sheet_name='Analise Completa')

                            # Aba apenas com erros (opcional)
                            erros = df_resultado[df_resultado['GRAVIDADE'] == 'ERRO']
                            if not erros.empty:
                                erros.to_excel(
                                    writer, index=False, sheet_name='Erros')

                            # Aba de estatísticas (opcional)
                            pd.DataFrame([stats]).to_excel(
                                writer, index=False, sheet_name='Resumo')

                            writer.close()

                        st.download_button(
                            label="📊 Download Análise Completa (XLSX)",
                            data=buffer.getvalue(),
                            file_name=f"analise_completa_{mes_selecionado}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

    with tab2:
        st.markdown("## 📈 Estatísticas Detalhadas")

        if 'df_dados' in st.session_state and st.session_state['df_dados'] is not None and not st.session_state['df_dados'].empty:
            df_base = st.session_state['df_dados']
            meses_disponiveis_stats = sorted(df_base['ANO MES'].unique())
            anos_disponiveis = ["Todos"] + sorted({int(m // 100) for m in meses_disponiveis_stats})

            col_filtro_1, col_filtro_2, col_filtro_3 = st.columns([1.2, 1, 1])
            with col_filtro_1:
                visao_stats = st.radio(
                    "Tipo de Estatística:",
                    ["Estatística Geral/Anual", "Estatística Mensal"],
                    horizontal=False
                )

            mes_alvo = None
            with col_filtro_2:
                ano_sel = st.selectbox("Ano:", anos_disponiveis, index=len(anos_disponiveis) - 1)
            with col_filtro_3:
                if ano_sel == "Todos":
                    meses_do_ano = sorted({int(m % 100) for m in meses_disponiveis_stats})
                else:
                    meses_do_ano = sorted({int(m % 100) for m in meses_disponiveis_stats if int(m // 100) == int(ano_sel)})
                if not meses_do_ano:
                    meses_do_ano = [1]
                mes_sel = st.selectbox("Mês:", meses_do_ano, index=len(meses_do_ano) - 1)
                if ano_sel == "Todos":
                    mes_alvo = meses_disponiveis_stats[-1]
                else:
                    mes_alvo = int(ano_sel) * 100 + int(mes_sel)

            df_res = None
            stats = {}

            if visao_stats == "Estatística Mensal":
                with st.spinner('🔄 Calculando estatísticas mensais...'):
                    df_res, stats = analisar_movimentacoes_mes(
                        df_base,
                        df_codigos,
                        regras_validas,
                        constantes,
                        mes_analise=mes_alvo
                    )
            else:
                if ano_sel == "Todos":
                    meses_geral = meses_disponiveis_stats
                else:
                    meses_geral = [m for m in meses_disponiveis_stats if int(m // 100) == int(ano_sel)]

                cache_key = f"geral_{ano_sel}"
                if 'df_resultado_geral' not in st.session_state:
                    st.session_state['df_resultado_geral'] = {}
                if 'stats_geral' not in st.session_state:
                    st.session_state['stats_geral'] = {}

                if cache_key not in st.session_state['df_resultado_geral'] or cache_key not in st.session_state['stats_geral']:
                    with st.spinner('🔄 Calculando estatísticas gerais...'):
                        df_res_geral, stats_geral = analisar_movimentacoes_periodo(
                            df_base,
                            df_codigos,
                            regras_validas,
                            constantes,
                            meses=meses_geral
                        )
                        st.session_state['df_resultado_geral'][cache_key] = df_res_geral
                        st.session_state['stats_geral'][cache_key] = stats_geral

                df_res = st.session_state['df_resultado_geral'][cache_key]
                stats = st.session_state['stats_geral'].get(cache_key, {})

            if df_res is None or df_res.empty:
                st.info("ℹ️ Não há dados para exibir com os filtros selecionados")
                st.stop()

            figuras_pdf = []

            # ============================================================================
            # SEÇÃO 1: VISÃO GERAL COM KPIS
            # ============================================================================
            st.markdown("### 🎯 Indicadores Chave de Performance (KPIs)")

            col1, col2, col3, col4, col5 = st.columns(5)

            total_movs = len(df_res)
            total_participantes = df_res['CODIGO ORGANIZACAO NOME'].nunique()
            taxa_erro = (stats.get('erros', 0) / total_participantes *
                         100) if total_participantes > 0 else 0
            taxa_conformidade = 100 - taxa_erro
            media_movs_participante = total_movs / \
                total_participantes if total_participantes > 0 else 0

            with col1:
                st.metric("👥 Participantes", f"{total_participantes:,}",
                          help="Total de participantes únicos analisados")
            with col2:
                st.metric(
                    "📋 Movimentações", f"{total_movs:,}", help="Total de registros de movimentação")
            with col3:
                st.metric("✅ Taxa Conformidade", f"{taxa_conformidade:.1f}%",
                          delta=f"{taxa_conformidade - 85:.1f}%" if taxa_conformidade >= 85 else None,
                          help="Percentual de participantes sem erros")
            with col4:
                st.metric("⚠️ Taxa de Erro", f"{taxa_erro:.1f}%",
                          delta=f"{taxa_erro - 15:.1f}%" if taxa_erro > 0 else "0%",
                          delta_color="inverse",
                          help="Percentual de participantes com erros críticos")
            with col5:
                st.metric("📊 Média Movs/Pessoa", f"{media_movs_participante:.1f}",
                          help="Número médio de movimentações por participante")
            st.markdown("---")

            # ============================================================================
            # SEÇÃO 2: ANÁLISE TEMPORAL E TENDÊNCIAS
            # ============================================================================
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("### 📅 Distribuição de Gravidade")

                # Gráfico de rosca com percentuais
                gravidade_counts = df_res.groupby(
                    'GRAVIDADE').size().reset_index(name='count')
                gravidade_counts['percentual'] = (
                    gravidade_counts['count'] / gravidade_counts['count'].sum() * 100).round(1)

                colors_gravidade = {'OK': '#28a745',
                                    'INFO': '#17a2b8', 'ERRO': '#dc3545'}

                fig = go.Figure(data=[go.Pie(
                    labels=gravidade_counts['GRAVIDADE'],
                    values=gravidade_counts['count'],
                    hole=0.5,
                    marker_colors=[colors_gravidade.get(
                        g, '#999') for g in gravidade_counts['GRAVIDADE']],
                    textinfo='label+percent',
                    textposition='outside',
                    hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
                )])

                fig.update_layout(
                    title="Classificação das Análises",
                    height=400,
                    showlegend=True,
                    annotations=[dict(
                        text=f'{total_participantes}<br>Participantes', x=0.5, y=0.5, font_size=16, showarrow=False)]
                )
                st.plotly_chart(fig, use_container_width=True)
                figuras_pdf.append(fig)

            with col2:
                st.markdown("### 🏢 Análise por Plano")

                if 'PLANO' in df_res.columns:
                    plano_gravidade = df_res.groupby(
                        ['PLANO', 'GRAVIDADE']).size().reset_index(name='count')

                    fig = px.bar(
                        plano_gravidade,
                        x='PLANO',
                        y='count',
                        color='GRAVIDADE',
                        title="Movimentações por Plano e Status",
                        color_discrete_map=colors_gravidade,
                        barmode='group',
                        text='count'
                    )

                    fig.update_traces(textposition='outside')
                    fig.update_layout(
                        xaxis_title="Plano",
                        yaxis_title="Quantidade",
                        height=400,
                        hovermode='x unified'
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    figuras_pdf.append(fig)
                else:
                    st.info("ℹ️ Coluna PLANO não disponível nos dados")

            st.markdown("---")

            # ============================================================================
            # SEÇÃO 3: ANÁLISE DE CÓDIGOS E BENEFÍCIOS
            # ============================================================================
            st.markdown("### 💼 Análise de Códigos de Benefício")

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("#### 📊 Top 10 Códigos Mais Utilizados")

                mov_por_codigo = df_res.groupby(
                    'CODIGO BENEFICIO').size().reset_index(name='count')
                mov_por_codigo = mov_por_codigo.merge(
                    df_codigos[['CODIGO', 'DESCRICAO', 'TIPO']],
                    left_on='CODIGO BENEFICIO',
                    right_on='CODIGO'
                )
                mov_por_codigo['percentual'] = (
                    mov_por_codigo['count'] / mov_por_codigo['count'].sum() * 100).round(2)

                top10 = mov_por_codigo.nlargest(10, 'count')

                fig = go.Figure(data=[go.Bar(
                    x=top10['count'],
                    y=top10['DESCRICAO'],
                    orientation='h',
                    text=top10['count'],
                    textposition='auto',
                    marker=dict(
                        color=top10['count'],
                        colorscale='Blues',
                        showscale=True,
                        colorbar=dict(title="Quantidade")
                    ),
                    hovertemplate='<b>%{y}</b><br>Código: ' + top10['CODIGO'].astype(
                        str) + '<br>Quantidade: %{x}<br>Percentual: ' + top10['percentual'].astype(str) + '%<extra></extra>'
                )])

                fig.update_layout(
                    title="Códigos com Maior Volume",
                    xaxis_title="Quantidade de Movimentações",
                    yaxis_title="",
                    height=450,
                    yaxis=dict(autorange="reversed")
                )
                st.plotly_chart(fig, use_container_width=True)
                figuras_pdf.append(fig)

            with col2:
                st.markdown("#### 🎭 Distribuição por Tipo de Código")

                tipo_dist = mov_por_codigo.groupby(
                    'TIPO')['count'].sum().reset_index()
                tipo_dist = tipo_dist.sort_values('count', ascending=False)

                colors_tipo = {'Benefício': '#ff7f0e', 'Instituto': '#2ca02c',
                               'População': '#1f77b4', 'Consolidador': '#d62728'}

                fig = go.Figure(data=[go.Bar(
                    x=tipo_dist['TIPO'],
                    y=tipo_dist['count'],
                    text=tipo_dist['count'],
                    textposition='auto',
                    marker_color=[colors_tipo.get(t, '#999')
                                  for t in tipo_dist['TIPO']],
                    hovertemplate='<b>%{x}</b><br>Quantidade: %{y}<extra></extra>'
                )])

                fig.update_layout(
                    title="Volume por Categoria",
                    xaxis_title="Tipo de Código",
                    yaxis_title="Quantidade",
                    height=450,
                    showlegend=False
                )
                st.plotly_chart(fig, use_container_width=True)
                figuras_pdf.append(fig)

            # Tabela detalhada
            with st.expander("📋 Ver Tabela Completa de Códigos"):
                st.dataframe(
                    mov_por_codigo[['CODIGO', 'DESCRICAO',
                                    'TIPO', 'count', 'percentual']]
                    .sort_values('count', ascending=False)
                    .rename(columns={'count': 'Quantidade', 'percentual': 'Percentual (%)', 'CODIGO': 'Código', 'DESCRICAO': 'Descrição'}),
                    use_container_width=True,
                    height=400
                )

            st.markdown("---")

            # ============================================================================
            # SEÇÃO 4: ANÁLISE DE TRANSIÇÕES (GRÁFICO DE BARRAS AGRUPADAS)
            # ============================================================================
            st.markdown("### 🔄 Análise de Transições")

            transicoes = df_res[df_res['MOVIMENTO'] == 'SAIDA'].merge(
                df_res[df_res['MOVIMENTO'] == 'ENTRADA'],
                on=['CODIGO ORGANIZACAO NOME', 'ANO MES'],
                suffixes=('_origem', '_destino')
            )

            if not transicoes.empty:
                col1, col2 = st.columns([2, 1])

                with col1:
                    st.markdown("#### 📊 Top 15 Transições Mais Frequentes")

                    # Prepara dados para o gráfico
                    trans_grouped = transicoes.groupby(
                        ['CODIGO BENEFICIO_origem', 'CODIGO BENEFICIO_destino']
                    ).size().reset_index(name='count')
                    trans_grouped = trans_grouped.nlargest(15, 'count')

                    # Mapeia códigos para nomes
                    codigo_to_desc = df_codigos.set_index(
                        'CODIGO')['DESCRICAO'].to_dict()

                    # Cria labels de transição
                    trans_grouped['transicao'] = trans_grouped.apply(
                        lambda row: f"{codigo_to_desc.get(row['CODIGO BENEFICIO_origem'], str(row['CODIGO BENEFICIO_origem']))[:20]}\n→\n{codigo_to_desc.get(row['CODIGO BENEFICIO_destino'], str(row['CODIGO BENEFICIO_destino']))[:20]}",
                        axis=1
                    )

                    trans_grouped['transicao_hover'] = trans_grouped.apply(
                        lambda row: f"{row['CODIGO BENEFICIO_origem']} → {row['CODIGO BENEFICIO_destino']}<br>{codigo_to_desc.get(row['CODIGO BENEFICIO_origem'], 'Desconhecido')}<br>para<br>{codigo_to_desc.get(row['CODIGO BENEFICIO_destino'], 'Desconhecido')}",
                        axis=1
                    )

                    # Cria gráfico de barras horizontais
                    fig = go.Figure(data=[go.Bar(
                        y=trans_grouped['transicao'],
                        x=trans_grouped['count'],
                        orientation='h',
                        text=trans_grouped['count'],
                        textposition='auto',
                        marker=dict(
                            color=trans_grouped['count'],
                            colorscale='Viridis',
                            showscale=True,
                            colorbar=dict(title="Quantidade")
                        ),
                        hovertemplate='<b>%{customdata}</b><br>Quantidade: %{x}<extra></extra>',
                        customdata=trans_grouped['transicao_hover']
                    )])

                    fig.update_layout(
                        title="Fluxos de Transição (Origem → Destino)",
                        xaxis_title="Quantidade de Participantes",
                        yaxis_title="",
                        height=600,
                        yaxis=dict(autorange="reversed"),
                        font=dict(size=10)
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    figuras_pdf.append(fig)

                with col2:
                    st.markdown("#### 📈 Estatísticas de Transições")

                    total_trans = len(transicoes)
                    trans_unicas = transicoes.groupby(
                        ['CODIGO BENEFICIO_origem', 'CODIGO BENEFICIO_destino']).size().shape[0]
                    trans_mais_comum = transicoes.groupby(
                        ['CODIGO BENEFICIO_origem', 'CODIGO BENEFICIO_destino']).size().idxmax()
                    trans_mais_comum_count = transicoes.groupby(
                        ['CODIGO BENEFICIO_origem', 'CODIGO BENEFICIO_destino']).size().max()

                    st.metric("🔀 Total de Transições", f"{total_trans:,}")
                    st.metric("🎯 Tipos Únicos", f"{trans_unicas}")

                    st.markdown("**🏆 Transição Mais Comum:**")
                    origem_desc = get_descricao(
                        trans_mais_comum[0], df_codigos)
                    destino_desc = get_descricao(
                        trans_mais_comum[1], df_codigos)
                    st.info(
                        f"{origem_desc[:25]}...\n\n↓\n\n{destino_desc[:25]}...\n\n**{trans_mais_comum_count} casos**")

                    # Top 5 transições
                    st.markdown("**📊 Top 5 Transições:**")
                    top5_trans = transicoes.groupby(['CODIGO BENEFICIO_origem', 'CODIGO BENEFICIO_destino']).size(
                    ).nlargest(5).reset_index(name='count')

                    for idx, row in top5_trans.iterrows():
                        origem = get_descricao(
                            row['CODIGO BENEFICIO_origem'], df_codigos)
                        destino = get_descricao(
                            row['CODIGO BENEFICIO_destino'], df_codigos)
                        st.markdown(
                            f"{idx+1}. `{row['CODIGO BENEFICIO_origem']}→{row['CODIGO BENEFICIO_destino']}` ({row['count']}x)")
                        st.caption(f"   {origem[:20]}... → {destino[:20]}...")

                # Heatmap de transições
                st.markdown("#### 🔥 Matriz de Calor - Transições")

                matriz = transicoes.groupby(
                    ['CODIGO BENEFICIO_origem', 'CODIGO BENEFICIO_destino']).size().reset_index(name='count')
                matriz_pivot = matriz.pivot(
                    index='CODIGO BENEFICIO_origem', columns='CODIGO BENEFICIO_destino', values='count').fillna(0)

                # Adiciona labels
                origem_labels = [
                    f"{int(idx)}<br>{get_descricao(int(idx), df_codigos)[:15]}" for idx in matriz_pivot.index]
                destino_labels = [
                    f"{int(col)}<br>{get_descricao(int(col), df_codigos)[:15]}" for col in matriz_pivot.columns]

                fig = go.Figure(data=go.Heatmap(
                    z=matriz_pivot.values,
                    x=destino_labels,
                    y=origem_labels,
                    colorscale='RdYlGn',
                    text=matriz_pivot.values,
                    texttemplate='%{text}',
                    textfont={"size": 10},
                    hovertemplate='Origem: %{y}<br>Destino: %{x}<br>Quantidade: %{z}<extra></extra>',
                    colorbar=dict(title="Qtd")
                ))

                fig.update_layout(
                    title="Mapa de Calor das Transições (Origem vs Destino)",
                    xaxis_title="Código Destino",
                    yaxis_title="Código Origem",
                    height=600,
                    xaxis={'side': 'bottom'},
                )
                st.plotly_chart(fig, use_container_width=True)
                figuras_pdf.append(fig)

            st.markdown("---")

            # ============================================================================
            # SEÇÃO 5: ANÁLISE DE ERROS
            # ============================================================================
            if stats.get('erros', 0) > 0:
                st.markdown("### ⚠️ Análise Detalhada de Erros")

                erros_df = df_res[df_res['GRAVIDADE'] == 'ERRO'].copy()

                col1, col2 = st.columns(2)

                with col1:
                    st.markdown("#### 🎯 Tipos de Erro Mais Comuns")

                    # Extrai tipo de erro da mensagem
                    erros_df['TIPO_ERRO'] = erros_df['ANALISE'].str.extract(
                        r'ERRO: ([^.]+)')[0]
                    tipo_erro_counts = erros_df.groupby('TIPO_ERRO').size().reset_index(
                        name='count').sort_values('count', ascending=False)

                    fig = px.bar(
                        tipo_erro_counts.head(10),
                        y='TIPO_ERRO',
                        x='count',
                        orientation='h',
                        title="Top 10 Tipos de Erro",
                        color='count',
                        color_continuous_scale='Reds',
                        text='count'
                    )
                    fig.update_traces(textposition='outside')
                    fig.update_layout(height=400, yaxis={
                                      'categoryorder': 'total ascending'})
                    st.plotly_chart(fig, use_container_width=True)
                    figuras_pdf.append(fig)

                with col2:
                    st.markdown("#### 🏢 Erros por Plano")

                    if 'PLANO' in erros_df.columns:
                        erros_plano = erros_df.groupby(
                            'PLANO').size().reset_index(name='count')

                        fig = go.Figure(data=[go.Pie(
                            labels=erros_plano['PLANO'],
                            values=erros_plano['count'],
                            hole=0.4,
                            marker_colors=px.colors.sequential.Reds[2:],
                            textinfo='label+value+percent'
                        )])

                        fig.update_layout(
                            title="Distribuição de Erros por Plano",
                            height=400
                        )
                        st.plotly_chart(fig, use_container_width=True)
                        figuras_pdf.append(fig)
                    else:
                        st.info("ℹ️ Coluna PLANO não disponível")

                # Ranking de códigos com erro
                st.markdown("#### 🚨 Códigos Mais Problemáticos")

                cod_erro = erros_df.groupby(
                    'CODIGO BENEFICIO').size().reset_index(name='erros')
                cod_erro = cod_erro.merge(
                    df_codigos[['CODIGO', 'DESCRICAO']], left_on='CODIGO BENEFICIO', right_on='CODIGO')
                cod_erro = cod_erro.sort_values(
                    'erros', ascending=False).head(10)

                fig = go.Figure(data=[go.Bar(
                    x=cod_erro['DESCRICAO'],
                    y=cod_erro['erros'],
                    text=cod_erro['erros'],
                    textposition='auto',
                    marker_color='crimson'
                )])

                fig.update_layout(
                    title="Top 10 Códigos com Mais Erros",
                    xaxis_title="Código",
                    yaxis_title="Quantidade de Erros",
                    height=400,
                    xaxis_tickangle=-45
                )
                st.plotly_chart(fig, use_container_width=True)
                figuras_pdf.append(fig)

            st.markdown("---")

            # ==========================================================================
            # SEÇÃO 6: INSIGHTS E RECOMENDAÇÕES
            # ==========================================================================
            st.markdown("### 💡 Insights e Recomendações")

            col1, col2, col3 = st.columns(3)

            with col1:
                st.markdown("#### ✅ Pontos Fortes")
                if taxa_conformidade >= 90:
                    st.success("✓ Excelente taxa de conformidade")
                if media_movs_participante < 5:
                    st.success(
                        "✓ Processos estão sendo concluídos rapidamente")
                if stats.get('info', 0) < stats.get('total', 1) * 0.3:
                    st.success("✓ Poucos processos pendentes")

            with col2:
                st.markdown("#### ⚠️ Pontos de Atenção")
                if taxa_erro > 10:
                    st.warning(
                        f"⚠ Taxa de erro acima de 10% ({taxa_erro:.1f}%)")
                if stats.get('info', 0) > stats.get('total', 1) * 0.3:
                    st.warning(
                        f"⚠ Muitos processos em andamento ({stats.get('info', 0)})")
                if media_movs_participante > 6:
                    st.warning("⚠ Muitas movimentações por participante")

            with col3:
                st.markdown("#### 🎯 Próximos Passos")
                if taxa_erro > 5:
                    st.info("→ Revisar casos com erro crítico")
                if stats.get('info', 0) > 20:
                    st.info(
                        f"→ Acompanhar {stats.get('info', 0)} processos pendentes")
                st.info("→ Monitorar tendências mensais")

            st.markdown("### 📄 Exportação")
            col_pdf_1, col_pdf_2 = st.columns([1, 2])
            with col_pdf_1:
                gerar_pdf = st.button(
                    "📄 Gerar PDF",
                    type="secondary",
                    use_container_width=True,
                    key=f"gerar_pdf_{visao_stats}_{ano_sel}_{mes_alvo}"
                )

            if gerar_pdf:
                try:
                    titulo = "Relatório de Estatísticas - ArcelorMittal"
                    if visao_stats == "Estatística Mensal":
                        subtitulo = f"Período: {mes_alvo}"
                    else:
                        if ano_sel == "Todos":
                            subtitulo = f"Período: {meses_disponiveis_stats[0]} a {meses_disponiveis_stats[-1]}"
                        else:
                            subtitulo = f"Ano: {ano_sel}"

                    kpis_pdf = {
                        'Participantes': f"{total_participantes:,}",
                        'Movimentações': f"{total_movs:,}",
                        'Taxa Conformidade': f"{taxa_conformidade:.1f}%",
                        'Taxa de Erro': f"{taxa_erro:.1f}%",
                        'Média Movs/Pessoa': f"{media_movs_participante:.1f}"
                    }

                    st.session_state['pdf_bytes'] = gerar_pdf_relatorio_visual(
                        titulo,
                        subtitulo,
                        kpis_pdf,
                        df_res,
                        df_codigos
                    )
                    st.success("✅ PDF gerado")
                except Exception as e:
                    st.error(f"❌ Não foi possível gerar o PDF: {e}")

            with col_pdf_2:
                if 'pdf_bytes' in st.session_state and st.session_state['pdf_bytes']:
                    sufixo = mes_alvo if visao_stats == 'Estatística Mensal' else (ano_sel if ano_sel != 'Todos' else 'geral')
                    nome_pdf = f"relatorio_estatisticas_{sufixo}.pdf"
                    st.download_button(
                        label="📥 Download PDF",
                        data=st.session_state['pdf_bytes'],
                        file_name=nome_pdf,
                        mime="application/pdf",
                        use_container_width=True
                    )

        else:
            st.info("ℹ️ Carregue um arquivo e/ou execute uma análise primeiro")

    with tab3:
        st.markdown("## 🔍 Busca de Participante")

        if 'df_resultado' in st.session_state:
            df_res = st.session_state['df_resultado']

            nome_busca = st.text_input(
                "Digite o nome ou código do participante:")

            if nome_busca:
                resultados = df_res[df_res['CODIGO ORGANIZACAO NOME'].str.contains(
                    nome_busca, case=False, na=False)]

                if not resultados.empty:
                    st.success(
                        f"✅ {len(resultados)} registro(s) encontrado(s)")

                    for participante in resultados['CODIGO ORGANIZACAO NOME'].unique():
                        with st.expander(f"👤 {participante}"):
                            dados_part = resultados[resultados['CODIGO ORGANIZACAO NOME']
                                                    == participante]
                            st.dataframe(
                                dados_part[['PLANO', 'CODIGO BENEFICIO',
                                            'MOVIMENTO', 'GRAVIDADE', 'ANALISE']],
                                use_container_width=True
                            )
                else:
                    st.warning("⚠️ Nenhum participante encontrado")
        else:
            st.info("ℹ️ Execute uma análise primeiro")

    with tab4:
        st.markdown("## 📚 Documentação")

        st.markdown("""
        ### 🎯 Objetivo do Sistema
        
        Este sistema automatiza a validação de movimentações previdenciárias, identificando:
        - ✅ Transições válidas conforme regras de negócio
        - ❌ Erros e inconsistências
        - ℹ️ Processos em andamento
        
        ### 📋 Códigos Principais
        """)

        st.dataframe(df_codigos, use_container_width=True, height=400)

        st.markdown("""
        ### 🔄 Como Usar
        
        1. **Upload**: Carregue seu arquivo Excel/CSV ou use dados de teste
        2. **Análise**: Selecione o mês e execute a análise
        3. **Resultados**: Visualize métricas, gráficos e exporte relatórios
        4. **Busca**: Encontre participantes específicos
        5. **Estatísticas**: Explore padrões e tendências
        
        ### ⚠️ Tipos de Alertas
        
        - **✅ OK**: Transição válida conforme regras
        - **ℹ️ INFO**: Processo em andamento (normal)
        - **❌ ERRO**: Inconsistência que precisa correção
        """)


if __name__ == "__main__":
    main()
