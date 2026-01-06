# -*- coding: utf-8 -*-
"""
Interface Streamlit para Análise de Movimentações Previdenciárias
ArcelorMittal - Sistema de Validação Automatizada - V2.0
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, Image
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import tempfile

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
# BASE DE CONHECIMENTO
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

    CODIGOS_IGNORAR = {31000, 32000, 33000, 34000}
    CODIGOS_RUIDO_SAIDA = {31100, 31200, 31300, 11000, 14000}
    CONTAS_ZERAGEM_ANUAL = {13000, 15000, 16000, 23000, 24100, 24200}
    CODIGOS_ADMISSAO = {31100, 31200}
    CODIGOS_ATIVOS = {31100, 31200, 31300}
    CODIGOS_SEM_RETORNO = {21000}
    CODIGOS_DESLIGAMENTO_RUIDO = {21000, 22000, 31300}
    
    # NOVOS: Códigos que podem existir isolados (sem vínculo)
    CODIGOS_ISOLADOS_PERMITIDOS = {24100, 24200, 13000}

    regras_validas = [
        (31100, 11100), (31200, 11100), (31300, 11100), (21000, 11100), (22000, 11100),
        (31100, 11200), (31200, 11200), (31300, 11200), (12000, 11200), (21000, 11200), (22000, 11200),
        (31100, 12000), (31200, 12000), (31300, 12000), (22000, 12000),
        (31100, 13000), (31200, 13000), (31300, 13000), (11100, 13000), (11200, 13000),
        (31100, 14000), (31200, 14000), (31300, 14000), (11100, 14000), (11200, 14000), (22000, 14000),
        (31100, 15000), (31200, 15000), (31300, 15000), (11100, 15000), (11200, 15000),
        (14000, 15000), (21000, 15000), (22000, 15000),
        (31100, 16000), (31200, 16000), (31300, 16000), (11100, 16000), (11200, 16000),
        (21000, 16000), (22000, 16000),
        (31100, 21000), (31200, 21000), (31300, 21000), (22000, 21000),
        (31100, 22000), (31200, 22000), (31300, 22000),
        (31100, 23000), (31200, 23000), (31300, 23000), (21000, 23000), (22000, 23000),
        (31100, 24100), (31200, 24100), (31300, 24100), (21000, 24100), (22000, 24100),
        (31100, 24200), (31200, 24200), (31300, 24200), (21000, 24200), (22000, 24200),
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
        'CODIGOS_ISOLADOS_PERMITIDOS': CODIGOS_ISOLADOS_PERMITIDOS
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
# MOTOR DE ANÁLISE APRIMORADO
# ============================================================================


def analisar_movimentacoes_mes(df_mov, df_codigos, regras_validas, constantes, mes_analise=None):
    """Motor de análise principal - VERSÃO APRIMORADA"""

    if mes_analise is None:
        mes_analise = df_mov['ANO MES'].max()

    df_mes = df_mov[df_mov['ANO MES'] == mes_analise].copy()

    if df_mes.empty:
        return df_mes, {'total': 0, 'erros': 0, 'info': 0, 'ok': 0}

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

        if 14000 in codigos_entrada_set and 15000 in codigos_entrada_set:
            msg = "ERRO: PENSÃO e PECÚLIO no mesmo mês"
            df_mes.loc[group.index, 'ANALISE'] = msg
            df_mes.loc[group.index, 'GRAVIDADE'] = 'ERRO'
            stats['erros'] += 1
            continue

        # NOVA VALIDAÇÃO: Verifica se existem códigos isolados permitidos
        codigos_isolados_presentes = codigos_entrada_set & constantes['CODIGOS_ISOLADOS_PERMITIDOS']
        
        # Se existem apenas códigos isolados (24100, 24200, 13000), marca como OK
        if codigos_isolados_presentes and len(codigos_entrada_set) == len(codigos_isolados_presentes):
            codigos_desc = ', '.join([get_descricao(c, df_codigos) for c in codigos_isolados_presentes])
            msg = f"OK: Lançamento independente - {codigos_desc}"
            df_mes.loc[group.index, 'ANALISE'] = msg
            df_mes.loc[group.index, 'GRAVIDADE'] = 'OK'
            df_mes.loc[group.index, 'TIPO_PASSO'] = '3. Fim'
            stats['ok'] += 1
            continue

        # Análise de transições
        saidas = group[group['MOVIMENTO'] == 'SAIDA']
        codigos_saida_set = set(saidas['CODIGO BENEFICIO'])
        codigos_intermediarios = codigos_saida_set.intersection(codigos_entrada_set)

        # Remove códigos isolados permitidos da análise de entradas líquidas
        codigos_entrada_sem_isolados = codigos_entrada_set - constantes['CODIGOS_ISOLADOS_PERMITIDOS']

        # Calcula saídas e entradas líquidas
        saidas_liquidas_brutas = codigos_saida_set - codigos_entrada_set

        if len(saidas_liquidas_brutas) > 1:
            saidas_liquidas = saidas_liquidas_brutas - constantes['CODIGOS_RUIDO_SAIDA']
        else:
            saidas_liquidas = saidas_liquidas_brutas

        entradas_liquidas = codigos_entrada_sem_isolados - codigos_saida_set

        # NOVA REGRA: Aceita saída de 31200 com entrada em 21000 E 31300
        if (31200 in saidas_liquidas and 
            21000 in entradas_liquidas and 
            31300 in entradas_liquidas and
            len(saidas_liquidas) == 1):
            
            msg = "OK: Transição válida - Ativo para BPD + Autopatrocinado"
            df_mes.loc[group.index, 'ANALISE'] = msg
            df_mes.loc[group.index, 'GRAVIDADE'] = 'OK'
            
            # Marca os passos corretamente
            for idx, row in group.iterrows():
                if row['CODIGO BENEFICIO'] == 31200 and row['MOVIMENTO'] == 'SAIDA':
                    df_mes.loc[idx, 'TIPO_PASSO'] = '1. Início'
                elif row['CODIGO BENEFICIO'] in {21000, 31300} and row['MOVIMENTO'] == 'ENTRADA':
                    df_mes.loc[idx, 'TIPO_PASSO'] = '3. Fim'
            
            stats['ok'] += 1
            continue

        # Classificação de passos
        for idx, row in group.iterrows():
            cod = row['CODIGO BENEFICIO']
            mov = row['MOVIMENTO']

            # Códigos isolados sempre são considerados "Fim"
            if cod in constantes['CODIGOS_ISOLADOS_PERMITIDOS'] and mov == 'ENTRADA':
                df_mes.loc[idx, 'TIPO_PASSO'] = '3. Fim'
            elif cod in saidas_liquidas and mov == 'SAIDA':
                df_mes.loc[idx, 'TIPO_PASSO'] = '1. Início'
            elif cod in entradas_liquidas and mov == 'ENTRADA':
                df_mes.loc[idx, 'TIPO_PASSO'] = '3. Fim'
            elif cod in codigos_intermediarios:
                df_mes.loc[idx, 'TIPO_PASSO'] = '2. Intermediário'

        msg = ''
        gravidade = 'OK'

        # ERRO: Múltiplas saídas líquidas
        if len(saidas_liquidas) > 1:
            msg = f"ERRO: Participante tem múltiplas saídas finais no mesmo mês ({', '.join(map(str, saidas_liquidas))})"
            gravidade = 'ERRO'
            stats['erros'] += 1

        elif len(entradas_liquidas) > 1:
            msg = f"ERRO: Múltiplas entradas finais distintas"
            gravidade = 'ERRO'
            stats['erros'] += 1

        elif len(saidas_liquidas) == 1 and len(entradas_liquidas) == 1:
            cod_origem = list(saidas_liquidas)[0]
            cod_destino = list(entradas_liquidas)[0]

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

        elif len(saidas_liquidas) == 0 and len(entradas_liquidas) > 0:
            # Se tem códigos isolados junto, não marca como INFO
            if not codigos_isolados_presentes:
                cod_entrada = list(entradas_liquidas)[0]
                plano = group['PLANO'].iloc[0] if 'PLANO' in group.columns else None

                if plano == 5 and cod_entrada in constantes['CODIGOS_ADMISSAO']:
                    msg = f"INFO: Nova admissão no Plano 5"
                    gravidade = 'INFO'
                    stats['info'] += 1
                else:
                    msg = f"INFO: Processo em andamento"
                    gravidade = 'INFO'
                    stats['info'] += 1

        elif len(saidas_liquidas) > 0 and len(entradas_liquidas) == 0:
            msg = f"INFO: Processo em andamento (aguardando conclusão)"
            gravidade = 'INFO'
            stats['info'] += 1

        if msg:
            df_mes.loc[group.index, 'ANALISE'] = msg
            df_mes.loc[group.index, 'GRAVIDADE'] = gravidade

    return df_mes, stats

# ============================================================================
# FUNÇÃO DE EXPORTAÇÃO PDF
# ============================================================================

def gerar_pdf_relatorio(df_resultado, stats, mes_analise, df_codigos):
    """Gera relatório em PDF com estatísticas e gráficos"""
    
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), 
                           rightMargin=30, leftMargin=30,
                           topMargin=50, bottomMargin=30)
    
    elements = []
    styles = getSampleStyleSheet()
    
    # Estilo customizado para título
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#1f77b4'),
        spaceAfter=30,
        alignment=TA_CENTER
    )
    
    # Título
    elements.append(Paragraph("Relatório de Análise de Movimentações Previdenciárias", title_style))
    elements.append(Paragraph(f"ArcelorMittal - Mês: {mes_analise}", styles['Heading2']))
    elements.append(Paragraph(f"Data de geração: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles['Normal']))
    elements.append(Spacer(1, 20))
    
    # Resumo Executivo
    elements.append(Paragraph("📊 Resumo Executivo", styles['Heading2']))
    
    resumo_data = [
        ['Métrica', 'Valor'],
        ['Total de Participantes', f"{stats['total']:,}"],
        ['Transações OK', f"{stats['ok']:,}"],
        ['Informações', f"{stats['info']:,}"],
        ['Erros Críticos', f"{stats['erros']:,}"],
        ['Taxa de Conformidade', f"{(stats['ok']/stats['total']*100):.1f}%" if stats['total'] > 0 else "N/A"],
        ['Taxa de Erro', f"{(stats['erros']/stats['total']*100):.1f}%" if stats['total'] > 0 else "N/A"]
    ]
    
    table = Table(resumo_data, colWidths=[4*inch, 2*inch])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f77b4')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    elements.append(table)
    elements.append(Spacer(1, 30))
    
    # Top 10 Códigos
    elements.append(Paragraph("💼 Top 10 Códigos Mais Utilizados", styles['Heading2']))
    
    mov_por_codigo = df_resultado.groupby('CODIGO BENEFICIO').size().reset_index(name='count')
    mov_por_codigo = mov_por_codigo.merge(
        df_codigos[['CODIGO', 'DESCRICAO']],
        left_on='CODIGO BENEFICIO',
        right_on='CODIGO'
    )
    top10 = mov_por_codigo.nlargest(10, 'count')
    
    top10_data = [['Código', 'Descrição', 'Quantidade']]
    for _, row in top10.iterrows():
        top10_data.append([str(row['CODIGO']), row['DESCRICAO'][:40], str(row['count'])])
    
    table2 = Table(top10_data, colWidths=[1*inch, 4*inch, 1.5*inch])
    table2.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2ca02c')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
    ]))
    
    elements.append(table2)
    elements.append(PageBreak())
    
    # Seção de Erros
    if stats['erros'] > 0:
        elements.append(Paragraph("⚠️ Análise de Erros Críticos", styles['Heading2']))
        
        erros_df = df_resultado[df_resultado['GRAVIDADE'] == 'ERRO']
        
        # Resumo de erros
        erros_resumo = [['Categoria', 'Quantidade']]
        
        # Extrai tipos de erro
        if not erros_df.empty:
            erros_df_copy = erros_df.copy()
            erros_df_copy['TIPO_ERRO'] = erros_df_copy['ANALISE'].str.extract(r'ERRO: ([^.]+)')[0]
            tipo_erro_counts = erros_df_copy.groupby('TIPO_ERRO').size().reset_index(name='count')
            
            for _, row in tipo_erro_counts.head(10).iterrows():
                erros_resumo.append([row['TIPO_ERRO'][:50], str(row['count'])])
        
        table3 = Table(erros_resumo, colWidths=[5*inch, 1.5*inch])
        table3.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#dc3545')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightpink),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        elements.append(table3)
    
    # Rodapé
    elements.append(Spacer(1, 40))
    elements.append(Paragraph(
        "Sistema de Validação Automatizada - ArcelorMittal v2.0", 
        ParagraphStyle('footer', parent=styles['Normal'], fontSize=8, textColor=colors.grey, alignment=TA_CENTER)
    ))
    
    doc.build(elements)
    buffer.seek(0)
    return buffer

# ============================================================================
# INTERFACE PRINCIPAL
# ============================================================================


def main():
    st.markdown('<div class="main-header">📊 Sistema de Análise de Movimentações Previdenciárias<br>ArcelorMittal v2.0</div>', unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.image("https://companieslogo.com/img/orig/MT_BIG.D-48309f61.png?t=1741059352",
                 use_container_width=True, width=150, caption="Projeto PrevicAuto - V2.0", clamp=True, output_format="PNG", channels="RGB")
        st.markdown("### ⚙️ Configurações")

        modo = st.radio(
            "Modo de Operação:",
            ["📁 Upload de Arquivo", "🧪 Dados de Teste"],
            help="Escolha entre carregar seus dados ou usar dados simulados"
        )

        st.markdown("---")
        st.markdown("### 📖 Sobre")
        st.info("""
        **Novidades v2.0:**
        - ✨ Filtro mensal em estatísticas
        - 📄 Exportação em PDF
        - 🔧 Regras de validação aprimoradas
        
        **Recursos:**
        - ✅ Validação de transições
        - 📊 Estatísticas detalhadas
        - 📈 Visualizações interativas
        - 📥 Export Excel e PDF
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
                            df_bruto = pd.read_csv(uploaded_file, sep=';', on_bad_lines='skip')

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

                        df_para_analise['CODIGO BENEFICIO'] = pd.to_numeric(
                            df_para_analise['CODIGO BENEFICIO'], errors='coerce')
                        df_para_analise['ANO MES'] = pd.to_numeric(
                            df_para_analise['ANO MES'], errors='coerce')

                        df_para_analise.dropna(
                            subset=['CODIGO BENEFICIO', 'ANO MES', 'CODIGO_ORG', 'NOME'], inplace=True)

                        df_para_analise['CODIGO BENEFICIO'] = df_para_analise['CODIGO BENEFICIO'].astype(int)
                        df_para_analise['ANO MES'] = df_para_analise['ANO MES'].astype(int)

                        df_para_analise = df_para_analise[~df_para_analise['CODIGO BENEFICIO'].isin(
                            constantes['CODIGOS_IGNORAR'])].copy()

                        df_para_analise = df_para_analise.sort_values('ANO MES').drop_duplicates(
                            subset=['CODIGO_ORG', 'NOME', 'CODIGO BENEFICIO', 'MOVIMENTO'],
                            keep='first'
                        )

                        df_para_analise['CODIGO ORGANIZACAO NOME'] = (
                            df_para_analise['CODIGO_ORG'].astype(str) + " - " +
                            df_para_analise['NOME'].apply(formatar_nome_participante)
                        )

                        st.success(f"✅ Arquivo carregado: {len(df_para_analise)} registros válidos")

                except Exception as e:
                    st.error(f"❌ Erro ao processar arquivo: {e}")

        else:  # Modo teste
            st.info("🧪 **Modo de Teste Ativado**")

            col1, col2 = st.columns(2)
            with col1:
                n_participantes = st.slider("Número de participantes:", 50, 500, 200)
            with col2:
                mes_teste = st.number_input("Mês de análise:", 202401, 202512, 202501)

            if st.button("🎲 Gerar Dados de Teste", type="primary"):
                with st.spinner('🔄 Gerando dados...'):
                    import random
                    random.seed(42)
                    dados = []

                    for i in range(n_participantes):
                        codigo_org = 50000000 + i
                        nome = f"Participante Teste {i+1}"
                        plano = random.choice([3, 4, 5, 6, 7])

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
                        df_para_analise['CODIGO_ORG'].astype(str) + " - " + df_para_analise['NOME']
                    )

                    st.success(f"✅ {len(df_para_analise)} registros de teste gerados")

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

                    st.session_state['df_resultado'] = df_resultado
                    st.session_state['df_completo'] = df_para_analise
                    st.session_state['stats'] = stats
                    st.session_state['mes_analise'] = mes_selecionado

                    st.success("✅ Análise concluída!")

                    # Métricas
                    st.markdown("### 📊 Resultados Gerais")
                    col1, col2, col3, col4 = st.columns(4)

                    with col1:
                        st.metric("👥 Total", stats['total'])
                    with col2:
                        st.metric("✅ OK", stats['ok'])
                    with col3:
                        st.metric("ℹ️ Info", stats['info'])
                    with col4:
                        st.metric("❌ Erros", stats['erros'])

                    # Gráfico de pizza
                    fig = go.Figure(data=[go.Pie(
                        labels=['OK', 'INFO', 'ERRO'],
                        values=[stats['ok'], stats['info'], stats['erros']],
                        marker_colors=['#44ff44', '#4488ff', '#ff4444'],
                        hole=0.4
                    )])
                    fig.update_layout(title="Distribuição das Análises", height=400)
                    st.plotly_chart(fig, use_container_width=True)

                    # Tabela de erros
                    if stats['erros'] > 0:
                        st.markdown("### ❌ Erros Encontrados")
                        erros = df_resultado[df_resultado['GRAVIDADE'] == 'ERRO']
                        st.dataframe(
                            erros[['CODIGO ORGANIZACAO NOME', 'PLANO', 'CODIGO BENEFICIO', 'MOVIMENTO', 'ANALISE']],
                            use_container_width=True,
                            height=400
                        )

                    # Downloads
                    st.markdown("### 📥 Baixar Resultados")
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                            df_resultado.to_excel(writer, index=False, sheet_name='Analise Completa')
                            erros = df_resultado[df_resultado['GRAVIDADE'] == 'ERRO']
                            if not erros.empty:
                                erros.to_excel(writer, index=False, sheet_name='Erros')
                            pd.DataFrame([stats]).to_excel(writer, index=False, sheet_name='Resumo')

                        st.download_button(
                            label="📊 Download Excel",
                            data=buffer.getvalue(),
                            file_name=f"analise_{mes_selecionado}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                    with col2:
                        pdf_buffer = gerar_pdf_relatorio(df_resultado, stats, mes_selecionado, df_codigos)
                        st.download_button(
                            label="📄 Download PDF",
                            data=pdf_buffer.getvalue(),
                            file_name=f"relatorio_{mes_selecionado}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )

    with tab2:
        st.markdown("## 📈 Estatísticas Detalhadas")

        if 'df_completo' in st.session_state and 'df_resultado' in st.session_state:
            df_completo = st.session_state['df_completo']
            
            # NOVO: Filtro de visualização
            st.markdown("### 🎯 Modo de Visualização")
            modo_estatistica = st.radio(
                "Escolha o tipo de estatística:",
                ["📅 Estatística Mensal", "📊 Estatística Consolidada (Todos os Meses)"],
                horizontal=True
            )
            
            if modo_estatistica == "📅 Estatística Mensal":
                # Modo mensal
                meses_disponiveis = sorted(df_completo['ANO MES'].unique())
                mes_escolhido = st.selectbox(
                    "Selecione o mês:",
                    meses_disponiveis,
                    index=meses_disponiveis.index(st.session_state.get('mes_analise', meses_disponiveis[-1]))
                )
                
                # Re-analisa para o mês escolhido
                with st.spinner('🔄 Carregando estatísticas do mês...'):
                    df_resultado_filtrado, stats_filtrado = analisar_movimentacoes_mes(
                        df_completo,
                        df_codigos,
                        regras_validas,
                        constantes,
                        mes_analise=mes_escolhido
                    )
                
                df_res = df_resultado_filtrado
                stats = stats_filtrado
                titulo_periodo = f"Mês: {mes_escolhido}"
                
            else:
                # Modo consolidado - analisa todos os meses
                st.info("📊 Analisando todos os meses disponíveis...")
                
                todos_resultados = []
                stats_consolidado = {'total': 0, 'ok': 0, 'info': 0, 'erros': 0}
                
                meses_disponiveis = sorted(df_completo['ANO MES'].unique())
                progress_bar = st.progress(0)
                
                for idx, mes in enumerate(meses_disponiveis):
                    df_mes, stats_mes = analisar_movimentacoes_mes(
                        df_completo,
                        df_codigos,
                        regras_validas,
                        constantes,
                        mes_analise=mes
                    )
                    todos_resultados.append(df_mes)
                    
                    # Acumula estatísticas
                    for key in stats_consolidado.keys():
                        stats_consolidado[key] += stats_mes.get(key, 0)
                    
                    progress_bar.progress((idx + 1) / len(meses_disponiveis))
                
                df_res = pd.concat(todos_resultados, ignore_index=True)
                stats = stats_consolidado
                titulo_periodo = f"Período: {meses_disponiveis[0]} a {meses_disponiveis[-1]}"
                progress_bar.empty()

            # Exibe as estatísticas
            st.markdown(f"### 📊 Análise - {titulo_periodo}")
            
            # KPIs
            col1, col2, col3, col4, col5 = st.columns(5)

            total_movs = len(df_res)
            total_participantes = df_res['CODIGO ORGANIZACAO NOME'].nunique()
            taxa_erro = (stats['erros'] / total_participantes * 100) if total_participantes > 0 else 0
            taxa_conformidade = 100 - taxa_erro
            media_movs_participante = total_movs / total_participantes if total_participantes > 0 else 0

            with col1:
                st.metric("👥 Participantes", f"{total_participantes:,}")
            with col2:
                st.metric("📋 Movimentações", f"{total_movs:,}")
            with col3:
                st.metric("✅ Taxa Conformidade", f"{taxa_conformidade:.1f}%")
            with col4:
                st.metric("⚠️ Taxa de Erro", f"{taxa_erro:.1f}%")
            with col5:
                st.metric("📊 Média Movs/Pessoa", f"{media_movs_participante:.1f}")

            st.markdown("---")

            # Gráficos
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("### 📅 Distribuição de Gravidade")
                gravidade_counts = df_res.groupby('GRAVIDADE').size().reset_index(name='count')
                colors_gravidade = {'OK': '#28a745', 'INFO': '#17a2b8', 'ERRO': '#dc3545'}

                fig = go.Figure(data=[go.Pie(
                    labels=gravidade_counts['GRAVIDADE'],
                    values=gravidade_counts['count'],
                    hole=0.5,
                    marker_colors=[colors_gravidade.get(g, '#999') for g in gravidade_counts['GRAVIDADE']],
                    textinfo='label+percent'
                )])
                fig.update_layout(title="Classificação das Análises", height=400)
                st.plotly_chart(fig, use_container_width=True)

            with col2:
                st.markdown("### 🏢 Análise por Plano")
                if 'PLANO' in df_res.columns:
                    plano_gravidade = df_res.groupby(['PLANO', 'GRAVIDADE']).size().reset_index(name='count')
                    fig = px.bar(
                        plano_gravidade,
                        x='PLANO',
                        y='count',
                        color='GRAVIDADE',
                        color_discrete_map=colors_gravidade,
                        barmode='group',
                        text='count'
                    )
                    fig.update_layout(height=400)
                    st.plotly_chart(fig, use_container_width=True)

            st.markdown("---")

            # Top Códigos
            st.markdown("### 💼 Top 10 Códigos Mais Utilizados")
            mov_por_codigo = df_res.groupby('CODIGO BENEFICIO').size().reset_index(name='count')
            mov_por_codigo = mov_por_codigo.merge(
                df_codigos[['CODIGO', 'DESCRICAO', 'TIPO']],
                left_on='CODIGO BENEFICIO',
                right_on='CODIGO'
            )
            top10 = mov_por_codigo.nlargest(10, 'count')

            fig = go.Figure(data=[go.Bar(
                x=top10['count'],
                y=top10['DESCRICAO'],
                orientation='h',
                text=top10['count'],
                textposition='auto',
                marker=dict(color=top10['count'], colorscale='Blues')
            )])
            fig.update_layout(height=450, yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig, use_container_width=True)

            # Análise de Erros
            if stats['erros'] > 0:
                st.markdown("### ⚠️ Análise de Erros")
                erros_df = df_res[df_res['GRAVIDADE'] == 'ERRO'].copy()
                erros_df['TIPO_ERRO'] = erros_df['ANALISE'].str.extract(r'ERRO: ([^.]+)')[0]
                tipo_erro_counts = erros_df.groupby('TIPO_ERRO').size().reset_index(name='count')

                fig = px.bar(
                    tipo_erro_counts.head(10),
                    y='TIPO_ERRO',
                    x='count',
                    orientation='h',
                    color='count',
                    color_continuous_scale='Reds'
                )
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)

            # Botão de download do PDF para estatísticas
            st.markdown("---")
            st.markdown("### 📥 Exportar Estatísticas")
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("📄 Gerar Relatório PDF", type="primary", use_container_width=True):
                    with st.spinner('📄 Gerando PDF...'):
                        mes_referencia = mes_escolhido if modo_estatistica == "📅 Estatística Mensal" else "CONSOLIDADO"
                        pdf_buffer = gerar_pdf_relatorio(df_res, stats, mes_referencia, df_codigos)
                        
                        st.download_button(
                            label="💾 Download Relatório PDF",
                            data=pdf_buffer.getvalue(),
                            file_name=f"estatisticas_{mes_referencia}_{datetime.now().strftime('%Y%m%d')}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )

        else:
            st.info("ℹ️ Execute uma análise primeiro na aba 'Análise'")

    with tab3:
        st.markdown("## 🔍 Busca de Participante")

        if 'df_resultado' in st.session_state:
            df_res = st.session_state['df_resultado']

            nome_busca = st.text_input("Digite o nome ou código do participante:")

            if nome_busca:
                resultados = df_res[df_res['CODIGO ORGANIZACAO NOME'].str.contains(
                    nome_busca, case=False, na=False)]

                if not resultados.empty:
                    st.success(f"✅ {len(resultados)} registro(s) encontrado(s)")

                    for participante in resultados['CODIGO ORGANIZACAO NOME'].unique():
                        with st.expander(f"👤 {participante}"):
                            dados_part = resultados[resultados['CODIGO ORGANIZACAO NOME'] == participante]
                            st.dataframe(
                                dados_part[['PLANO', 'CODIGO BENEFICIO', 'MOVIMENTO', 'GRAVIDADE', 'ANALISE']],
                                use_container_width=True
                            )
                else:
                    st.warning("⚠️ Nenhum participante encontrado")
        else:
            st.info("ℹ️ Execute uma análise primeiro")

    with tab4:
        st.markdown("## 📚 Documentação do Sistema v2.0")

        st.markdown("""
        ### 🎯 Objetivo do Sistema
        
        Este sistema automatiza a validação de movimentações previdenciárias, identificando:
        - ✅ Transições válidas conforme regras de negócio
        - ❌ Erros e inconsistências
        - ℹ️ Processos em andamento
        
        ### 🆕 Novidades da Versão 2.0
        
        #### 1. 📊 Estatísticas Mensais e Consolidadas
        - Visualize estatísticas de um mês específico
        - Ou analise dados consolidados de todos os meses
        - Filtro intuitivo para alternar entre modos
        
        #### 2. 📄 Exportação em PDF
        - Gere relatórios profissionais em PDF
        - Inclui todos os gráficos e análises
        - Ideal para apresentações e documentação
        
        #### 3. 🔧 Regras de Validação Aprimoradas
        
        **a) Transição Ativo → BPD + Autopatrocinado**
        - ✅ Agora aceita: Saída de 31200 com entrada em 21000 E 31300
        - Esta é uma movimentação válida quando participante vai para BPD e mantém autopatrocínio
        
        **b) Portabilidade (24100 e 24200)**
        - ✅ Podem existir de forma isolada
        - Sempre são ENTRADA, nunca saída
        - Não exigem vinculação com outras contas
        
        **c) Auxílio Funeral (13000)**
        - ✅ Pode aparecer sozinho ou com outras movimentações
        - Sempre ENTRADA, nunca saída
        - Independente de outras transições
        
        ### 📋 Códigos Principais
        """)

        st.dataframe(df_codigos, use_container_width=True, height=400)

        st.markdown("""
        ### 🔄 Como Usar o Sistema
        
        #### 📂 Passo 1: Importação de Dados
        1. Escolha entre **Upload de Arquivo** ou **Dados de Teste**
        2. Para upload: selecione arquivo Excel (.xlsx) ou CSV
        3. Aguarde o processamento e validação dos dados
        
        #### 🔬 Passo 2: Análise
        1. Selecione o mês que deseja analisar
        2. Clique em **▶️ Executar Análise**
        3. Visualize os resultados: métricas, gráficos e erros
        4. Baixe relatórios em Excel ou PDF
        
        #### 📊 Passo 3: Estatísticas
        1. Acesse a aba **📊 Estatísticas**
        2. Escolha entre:
           - **📅 Estatística Mensal**: dados de um mês específico
           - **📊 Estatística Consolidada**: todos os meses juntos
        3. Explore gráficos detalhados e insights
        4. Exporte estatísticas em PDF
        
        #### 🔍 Passo 4: Busca
        1. Procure participantes específicos
        2. Visualize histórico detalhado
        3. Identifique problemas individuais
        
        ### ⚠️ Interpretação dos Alertas
        
        | Status | Significado | Ação Recomendada |
        |--------|-------------|------------------|
        | ✅ **OK** | Transição válida conforme regras | Nenhuma ação necessária |
        | ℹ️ **INFO** | Processo em andamento (normal) | Monitorar evolução |
        | ❌ **ERRO** | Inconsistência que precisa correção | Revisar imediatamente |
        
        ### 📈 Exemplos de Validação
        
        #### ✅ Casos Válidos:
        
        1. **Aposentadoria Normal**
           - Saída: 31200 (Ativo) → Entrada: 11100 (Aposentadoria Normal)
           
        2. **BPD + Autopatrocínio** *(NOVO)*
           - Saída: 31200 (Ativo) → Entrada: 21000 (BPD) + 31300 (Autopatrocinado)
           
        3. **Portabilidade Isolada** *(NOVO)*
           - Entrada: 24100 ou 24200 (sem necessidade de saída vinculada)
           
        4. **Auxílio Funeral Independente** *(NOVO)*
           - Entrada: 13000 (pode aparecer sozinho ou com outras movimentações)
        
        #### ❌ Casos de Erro:
        
        1. **Múltiplas Situações Ativas**
           - Participante com 31100 E 31200 ativos simultaneamente no mesmo plano
           
        2. **Pensão + Pecúlio**
           - Entrada de 14000 (Pensão) E 15000 (Pecúlio) no mesmo mês
           
        3. **BPD Retornando para Ativo**
           - Saída: 21000 (BPD) → Entrada: 31100/31200/31300 (qualquer ativo)
        
        ### 💡 Dicas de Uso
        
        - 📊 Use **Estatística Mensal** para análise detalhada de um período
        - 📈 Use **Estatística Consolidada** para identificar tendências
        - 📄 Gere PDFs para documentar análises e compartilhar com equipe
        - 🔍 Use a busca para investigar casos específicos
        - 📥 Baixe Excel para análises customizadas adicionais
        
        ### 🆘 Suporte
        
        Em caso de dúvidas:
        1. Revise esta documentação
        2. Verifique os exemplos de validação
        3. Entre em contato com a equipe de TI
        """)

        st.markdown("---")
        st.info("💼 **ArcelorMittal - Sistema de Validação Automatizada v2.0** | Desenvolvido para otimizar processos previdenciários")


if __name__ == "__main__":
    main()
