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
    CODIGOS_IGNORAR = {31000, 32000, 33000, 34000}

    # CÓDIGOS QUE CAUSAM RUÍDO EM MÚLTIPLAS SAÍDAS (devem ser filtrados ao calcular saídas líquidas)
    CODIGOS_RUIDO_SAIDA = {31100, 31200, 31300, 11000, 14000}

    CONTAS_ZERAGEM_ANUAL = {13000, 15000, 16000, 23000, 24100, 24200}
    CODIGOS_ADMISSAO = {31100, 31200}
    CODIGOS_ATIVOS = {31100, 31200, 31300}
    CODIGOS_SEM_RETORNO = {21000}
    CODIGOS_DESLIGAMENTO_RUIDO = {21000, 22000, 31300}

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
        'CODIGOS_RUIDO_SAIDA': CODIGOS_RUIDO_SAIDA
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

        if 14000 in codigos_entrada_set and 15000 in codigos_entrada_set:
            msg = "ERRO: PENSÃO e PECÚLIO no mesmo mês"
            df_mes.loc[group.index, 'ANALISE'] = msg
            df_mes.loc[group.index, 'GRAVIDADE'] = 'ERRO'
            stats['erros'] += 1
            continue

        # Análise de transições
        saidas = group[group['MOVIMENTO'] == 'SAIDA']
        codigos_saida_set = set(saidas['CODIGO BENEFICIO'])
        codigos_intermediarios = codigos_saida_set.intersection(
            codigos_entrada_set)

        # Calcula saídas e entradas líquidas
        saidas_liquidas_brutas = codigos_saida_set - codigos_entrada_set
        entradas_liquidas = codigos_entrada_set - codigos_saida_set

        # FILTRAR CÓDIGOS DE RUÍDO APENAS SE HOUVER MÚLTIPLAS SAÍDAS
        # Se há apenas 1 saída, ela é legítima (mesmo sendo 31100, 31200, etc)
        # Se há múltiplas saídas, remove os códigos consolidadores que são ruído
        if len(saidas_liquidas_brutas) > 1:
            saidas_liquidas = saidas_liquidas_brutas - \
                constantes['CODIGOS_RUIDO_SAIDA']
        else:
            saidas_liquidas = saidas_liquidas_brutas

        # Classificação de passos
        for idx, row in group.iterrows():
            cod = row['CODIGO BENEFICIO']
            mov = row['MOVIMENTO']

            if cod in saidas_liquidas and mov == 'SAIDA':
                df_mes.loc[idx, 'TIPO_PASSO'] = '1. Início'
            elif cod in entradas_liquidas and mov == 'ENTRADA':
                df_mes.loc[idx, 'TIPO_PASSO'] = '3. Fim'
            elif cod in codigos_intermediarios:
                df_mes.loc[idx, 'TIPO_PASSO'] = '2. Intermediário'

        msg = ''
        gravidade = 'OK'

        # ERRO: Múltiplas saídas líquidas (após filtrar códigos de ruído)
        if len(saidas_liquidas) > 1:
            msg = f"ERRO: Participante tem múltiplas saídas finais no mesmo mês ({', '.join(map(str, saidas_liquidas))}). Isso é muito raro e pode indicar problema no sistema."
            gravidade = 'ERRO'
            stats['erros'] += 1

        elif len(entradas_liquidas) > 1:
            msg = f"ERRO: Múltiplas entradas finais"
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
# INTERFACE PRINCIPAL
# ============================================================================


def main():
    st.markdown('<div class="main-header">📊 Sistema de Análise de Movimentações Previdenciárias<br>ArcelorMittal</div>', unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.image("https://via.placeholder.com/200x80/1f77b4/FFFFFF?text=ArcelorMittal",
                 use_container_width=True)
        st.markdown("### ⚙️ Configurações")

        modo = st.radio(
            "Modo de Operação:",
            ["📁 Upload de Arquivo", "🧪 Dados de Teste"],
            help="Escolha entre analisar seus dados ou usar dados de teste"
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
                        df_para_analise = df_para_analise[~df_para_analise['CODIGO BENEFICIO'].isin(
                            constantes['CODIGOS_IGNORAR'])].copy()

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

                        # Download
                        csv = erros.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            "📥 Download Erros (XLSX)",
                            csv,
                            f"erros_{mes_selecionado}.csv",
                            "text/csv"
                        )

    with tab2:
        st.markdown("## 📈 Estatísticas Detalhadas")

        if 'df_resultado' in st.session_state:
            df_res = st.session_state['df_resultado']
            stats = st.session_state.get('stats', {})

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
                on='CODIGO ORGANIZACAO NOME',
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

            st.markdown("---")

            # ============================================================================
            # SEÇÃO 6: INSIGHTS E RECOMENDAÇÕES
            # ============================================================================
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

        else:
            st.info("ℹ️ Execute uma análise primeiro na aba 'Análise'")

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
