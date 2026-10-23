import io
import os
import json
import base64
import textwrap
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import datetime

from fechamento_engine import (
    processar_lote,
    gerar_tsv_colar_excel,
    gerar_excel_fechamento,
    extrair_chaves_historico,
    formatar_moeda_br
)

from auditar_engine import (
    ENDERECOS_BASES,
    ENDERECOS_BASES_DETALHADOS,
    OPCOES_BASES_FORMATADAS,
    detectar_base_por_texto,
    obter_endereco_base,
    carregar_planilha_segura,
    processar_auditoria_completa,
    exportar_relatorio_excel
)

def carregar_simbolo_consigaz_b64() -> str:
    caminho = os.path.join(os.path.dirname(__file__), "assets", "consigaz_symbol.png")
    if os.path.exists(caminho):
        try:
            with open(caminho, "rb") as f:
                return base64.b64encode(f.read()).decode("utf-8")
        except Exception:
            return ""
    return ""

caminho_icone = os.path.join(os.path.dirname(__file__), "assets", "consigaz_symbol.png")
page_icon = caminho_icone if os.path.exists(caminho_icone) else "🔥"

st.set_page_config(
    page_title="Consigaz - Auditoria de KM por Dia",
    page_icon=page_icon,
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS completo: MODO NOTURNO (Dark Mode) Corporativo Elegante e de Alto Contraste
st.markdown("""
<style>
/* Fundo global e layout Dark Mode */
.stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"], [data-testid="stToolbar"] {
    background-color: #0F172A !important;
    color: #F8FAFC !important;
}

[data-testid="stSidebar"], [data-testid="stSidebarContent"] {
    background-color: #0B1120 !important;
    color: #F8FAFC !important;
}

label, [data-testid="stWidgetLabel"] p, [data-testid="stMarkdownContainer"] p, [data-testid="stCaptionContainer"] p, div[data-baseweb="select"] span {
    color: #F8FAFC !important;
}

h1, h2, h3, h4, h5, h6 {
    color: #38BDF8 !important;
}

/* Tabs Dark Mode */
button[data-baseweb="tab"] p {
    color: #94A3B8 !important;
    font-weight: 600 !important;
    font-size: 15px !important;
}

button[data-baseweb="tab"][aria-selected="true"] p {
    color: #38BDF8 !important;
    font-weight: 700 !important;
}

/* Botões de Ação */
.stButton>button {
    background-color: #0284C7 !important;
    color: #FFFFFF !important;
    font-weight: 700 !important;
    border-radius: 6px !important;
    height: 44px !important;
    border: none !important;
    box-shadow: 0 2px 8px rgba(2, 132, 199, 0.3) !important;
    transition: all 0.2s ease !important;
}

.stButton>button p, .stButton>button span {
    color: #FFFFFF !important;
}

.stButton>button:hover {
    background-color: #0369A1 !important;
    color: #FFFFFF !important;
    box-shadow: 0 4px 12px rgba(2, 132, 199, 0.45) !important;
}

/* Caixas de Métricas Streamlit */
[data-testid="stMetric"] {
    background-color: #1E293B !important;
    border: 1px solid #334155 !important;
    border-radius: 8px !important;
    padding: 12px 14px !important;
}

[data-testid="stMetricLabel"] p {
    color: #94A3B8 !important;
    font-size: 12px !important;
    font-weight: 700 !important;
}

[data-testid="stMetricValue"] div {
    color: #F8FAFC !important;
    font-weight: 800 !important;
}

/* Formulários e Inputs */
input, textarea {
    background-color: #1E293B !important;
    color: #F8FAFC !important;
    border: 1px solid #475569 !important;
    border-radius: 6px !important;
}

div[data-baseweb="select"] > div {
    background-color: #1E293B !important;
    color: #F8FAFC !important;
    border-color: #475569 !important;
}

/* Expanders */
[data-testid="stExpander"] {
    background-color: #1E293B !important;
    border: 1px solid #334155 !important;
    border-radius: 8px !important;
}

/* Cabeçalho */
.main-header {
    font-size: 26px;
    font-weight: 800;
    color: #F8FAFC !important;
    margin-bottom: 2px;
    letter-spacing: -0.3px;
}

.sub-header {
    font-size: 14.5px;
    color: #94A3B8 !important;
    margin-bottom: 18px;
    line-height: 1.4;
}

/* Tabelas Dark Mode */
.table-km {
    width: 100%;
    border-collapse: collapse;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    font-size: 13.5px;
    margin-top: 8px;
    margin-bottom: 18px;
    background-color: #0F172A;
    border: 1px solid #334155;
    border-radius: 6px;
    overflow: hidden;
}

.table-km th {
    background-color: #1E3A8A !important;
    color: #FFFFFF !important;
    text-align: center;
    padding: 10px 8px;
    font-weight: 700;
    border: 1px solid #294A9A;
}

.table-km td {
    padding: 9px 10px;
    border: 1px solid #334155;
    color: #F1F5F9 !important;
    vertical-align: middle;
}

.table-km tr:nth-child(even) {
    background-color: #162032;
}

.table-km tr:hover {
    background-color: #243046;
}

.row-highlight {
    background-color: #382B10 !important;
}

.row-highlight td {
    background-color: #382B10 !important;
    color: #FDE68A !important;
    border-color: #78350F !important;
}

.row-total td {
    background-color: #1E3A8A !important;
    color: #FFFFFF !important;
    font-weight: 800 !important;
    font-size: 14px !important;
}

/* Badges de Status Dark Mode */
.badge-alto {
    background-color: #991B1B !important;
    color: #FEE2E2 !important;
    border: 1px solid #EF4444 !important;
    padding: 4px 10px;
    border-radius: 12px;
    font-size: 12.5px;
    font-weight: 700;
    display: inline-block;
}

.badge-normal {
    background-color: #14532D !important;
    color: #DCFCE7 !important;
    border: 1px solid #22C55E !important;
    padding: 4px 10px;
    border-radius: 12px;
    font-size: 12.5px;
    font-weight: 700;
    display: inline-block;
}

.badge-mapa {
    background-color: #0369A1 !important;
    color: #E0F2FE !important;
    border: 1px solid #38BDF8 !important;
    padding: 4px 10px;
    border-radius: 12px;
    font-size: 13px;
    font-weight: 700;
    display: inline-block;
}

.badge-cobrando {
    background-color: #854D0E !important;
    color: #FEF3C7 !important;
    border: 1px solid #F59E0B !important;
    padding: 3px 8px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 700;
    display: inline-block;
    margin: 2px 0;
}

.badge-sem-km {
    background-color: #334155 !important;
    color: #94A3B8 !important;
    border: 1px solid #475569 !important;
    padding: 2px 7px;
    border-radius: 6px;
    font-size: 11.5px;
    font-weight: 600;
    display: inline-block;
}

.btn-maps {
    background-color: #0284C7 !important;
    color: #FFFFFF !important;
    padding: 6px 11px;
    border-radius: 5px;
    text-decoration: none !important;
    font-weight: 700;
    font-size: 12px;
    display: inline-block;
    margin: 2px;
    transition: background-color 0.15s ease;
}

.btn-maps:hover {
    background-color: #0369A1 !important;
    color: #FFFFFF !important;
}

.btn-maps-sec {
    background-color: #334155 !important;
    color: #E2E8F0 !important;
    border: 1px solid #475569 !important;
    padding: 5px 9px;
    border-radius: 5px;
    text-decoration: none !important;
    font-weight: 600;
    font-size: 11px;
    display: inline-block;
    margin: 2px;
}

.btn-maps-sec:hover {
    background-color: #475569 !important;
    color: #FFFFFF !important;
}

/* Cards de Dia ("Tudo num Card só") */
.card-dia {
    background-color: #1E293B !important;
    border: 1.5px solid #334155 !important;
    border-radius: 12px !important;
    padding: 20px 22px !important;
    margin-bottom: 26px !important;
    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.3) !important;
}

.card-dia:hover {
    border-color: #475569 !important;
    box-shadow: 0 6px 22px rgba(0, 0, 0, 0.45) !important;
}

.card-dia-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 1.5px solid #334155;
    padding-bottom: 12px;
    margin-bottom: 14px;
    flex-wrap: wrap;
    gap: 10px;
}

.card-dia-data {
    font-size: 19px;
    font-weight: 800;
    color: #38BDF8 !important;
}

.card-dia-metricas {
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    background-color: #0F172A !important;
    border-radius: 8px;
    padding: 14px 18px;
    margin-bottom: 16px;
    border: 1px solid #334155;
}

.metrica-box {
    flex: 1 1 200px;
    display: flex;
    flex-direction: column;
}

.metrica-label {
    font-size: 11px;
    font-weight: 700;
    color: #94A3B8 !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 3px;
}

.metrica-val {
    font-size: 18px;
    font-weight: 800;
}

.metrica-sub {
    font-size: 12px;
    color: #94A3B8 !important;
    margin-top: 2px;
    line-height: 1.4;
}

/* Transforma os quadradinhos do uploader em lista vertical (listinha) */
[data-testid="stFileUploaderFileList"] {
    display: flex !important;
    flex-direction: column !important;
    gap: 8px !important;
    width: 100% !important;
    margin-top: 10px !important;
}

[data-testid="stFileUploaderFile"] {
    display: flex !important;
    flex-direction: row !important;
    align-items: center !important;
    justify-content: space-between !important;
    width: 100% !important;
    background-color: #1E293B !important;
    border: 1px solid #334155 !important;
    border-radius: 8px !important;
    padding: 10px 14px !important;
    margin: 0 !important;
}

[data-testid="stFileUploaderFileName"] {
    color: #F8FAFC !important;
    font-weight: 600 !important;
    font-size: 14px !important;
}

.item-lista-arquivo {
    display: flex;
    align-items: center;
    justify-content: space-between;
    background-color: #1E293B;
    border: 1px solid #334155;
    border-radius: 8px;
    padding: 9px 14px;
    margin-bottom: 6px;
    font-size: 13.5px;
    transition: all 0.2s ease;
}

.item-lista-arquivo:hover {
    border-color: #38BDF8;
    background-color: #243248;
}
</style>
""", unsafe_allow_html=True)

simbolo_b64 = carregar_simbolo_consigaz_b64()
if simbolo_b64:
    tag_logo = f'<img src="data:image/png;base64,{simbolo_b64}" alt="Consigaz" style="height:48px;width:auto;object-fit:contain;filter:drop-shadow(0 2px 6px rgba(220,62,49,0.45));" />'
    tag_logo_sidebar = f'<img src="data:image/png;base64,{simbolo_b64}" alt="Consigaz" style="height:32px;width:auto;object-fit:contain;filter:drop-shadow(0 1px 4px rgba(220,62,49,0.45));" />'
else:
    tag_logo = '<span style="font-size:32px;">🔥</span>'
    tag_logo_sidebar = '<span style="font-size:24px;">🔥</span>'



def render_auditoria_km(tag_logo, tag_logo_sidebar):
    # Título Principal com Símbolo Oficial da Consigaz
    st.markdown(
        f"""
        <div style="display:flex;align-items:center;gap:14px;margin-bottom:6px;padding-bottom:4px;">
            {tag_logo}
            <div>
                <div class="main-header">Consigaz — Auditoria Automática de KM por Dia</div>
                <div class="sub-header">Soma todos os KMs ditos no dia, calcula a maior rota rodoviária possível no Google Maps (teto máximo com Base no início e no fim) e mostra se bate ou se há cobrança a maior</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Barra Lateral - Configurações
    st.sidebar.markdown(
        f"""
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:14px;">
            {tag_logo_sidebar}
            <span style="font-size:18px;font-weight:800;color:#F8FAFC;letter-spacing:0.3px;">Parâmetros da Auditoria</span>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    lista_labels = list(OPCOES_BASES_FORMATADAS.keys())
    indice_padrao = 0
    for idx, lbl in enumerate(lista_labels):
        if "CAN" in lbl or "Base 6" in lbl:
            indice_padrao = idx
            break
    
    label_base_selecionada = st.sidebar.selectbox(
        "Base Operacional (Tabela Oficial Consigaz):",
        options=lista_labels,
        index=indice_padrao,
        help="Selecione a base correspondente para carregar o endereço e estado (UF) oficial de partida e retorno."
    )
    
    dados_base = OPCOES_BASES_FORMATADAS[label_base_selecionada]
    base_selecionada = dados_base["sigla"]
    endereco_base_sugerido = dados_base["endereco"]
    
    endereco_partida = st.sidebar.text_area(
        "Endereço de Partida e Retorno (Base Consigaz):",
        value=endereco_base_sugerido,
        height=80,
        help="Endereço oficial da Base Consigaz utilizado como ponto inicial e final da rota."
    )
    
    tolerancia_km = st.sidebar.number_input(
        "Margem de Erro / Tolerância Aceitável (KM):",
        min_value=0.0,
        max_value=200.0,
        value=40.0,
        step=5.0,
        help="Margem de erro que você aceita. Exemplo: se você deixar 40 km, numa OS o técnico cobra 300 km e no Maps der 270 km (diferença de 30 km), ele é aprovado automaticamente!"
    )
    
    with st.sidebar.expander("⚙️ Franquia Contratual (Opcional)", expanded=False):
        raio_franquia = st.number_input(
            "Franquia de Raio (KM):",
            min_value=0.0,
            max_value=200.0,
            value=0.0,
            step=5.0,
            help="Deixe em 0 km se você paga o KM cheio. Ajuste apenas se o contrato do prestador descontar um raio gratuito inicial para pagamento."
        )
    
    senha_planilha = st.sidebar.text_input(
        "Senha da Planilha (se houver):",
        value="FECHAMENTO",
        type="password",
        help="Senha padrão de fechamento utilizada para desbloquear os arquivos (testa FECHAMENTO e fechamento automaticamente)."
    )
    
    st.sidebar.markdown("---")
    st.sidebar.markdown(
        """
        <div style="text-align: center; padding: 14px 10px; background-color: #1E293B; border-radius: 8px; border: 1.5px solid #334155; margin-top: 18px; box-shadow: 0 4px 12px rgba(0,0,0,0.3);">
            <span style="font-size: 16px; font-weight: 800; color: #38BDF8; letter-spacing: 0.5px;">Desenvolvido por Chiku</span>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Área de Upload
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. Planilha de Fechamento")
        arquivo_fechamento = st.file_uploader(
            "Arraste a planilha enviada pela base (.xlsx)",
            type=["xlsx", "xlsm"],
            key="upload_fechamento"
        )
        if arquivo_fechamento:
            base_detectada = detectar_base_por_texto(arquivo_fechamento.name)
            st.success(f"Arquivo carregado: **{arquivo_fechamento.name}** (Base sugerida pelo nome: **{base_detectada}**)")
    
    with col2:
        st.subheader("2. Relatório Totvs (ESOS)")
        arquivo_totvs = st.file_uploader(
            "Arraste a planilha de consulta Totvs com as OSs (.xlsx)",
            type=["xlsx", "xlsm"],
            key="upload_totvs"
        )
        if arquivo_totvs:
            st.success(f"Arquivo carregado: **{arquivo_totvs.name}**")
    
    st.markdown("---")
    
    if arquivo_fechamento and arquivo_totvs:
        if st.button("🚀 Calcular KMs por Dia Automaticamente", use_container_width=True):
            with st.spinner("Cruzando OSs do dia, calculando quilometragem rodoviária automática e gerando rotas sem erro..."):
                try:
                    wb_fech = carregar_planilha_segura(arquivo_fechamento, senha=senha_planilha)
                    wb_totvs = carregar_planilha_segura(arquivo_totvs, senha=senha_planilha)
    
                    resultado = processar_auditoria_completa(
                        wb_fechamento=wb_fech,
                        wb_totvs=wb_totvs,
                        base_codigo=base_selecionada,
                        base_endereco_custom=endereco_partida,
                        raio_franquia=float(raio_franquia),
                        tolerancia_km=float(tolerancia_km)
                    )
    
                    st.session_state["resultado_auditoria"] = resultado
                    st.session_state["nome_fechamento"] = arquivo_fechamento.name
    
                except Exception as e:
                    st.error(f"Erro ao processar as planilhas: {e}")
    
    # Exibição dos Resultados
    if "resultado_auditoria" in st.session_state:
        res = st.session_state["resultado_auditoria"]
        por_dia = res.get("por_dia", [])
        linhas_os = res.get("resumo", [])
    
        tab_dia, tab_relatorio_pdf, tab_os, tab_exportar = st.tabs([
            "📅 1. Soma de KMs por Dia vs KM Calculado da Rota",
            "📄 2. Relatório de Deslocamento (Estilo PDF Consigaz)",
            "📋 3. Tabela Individual por OS Cobrada",
            "📥 4. Baixar Relatório em Excel"
        ])
    
        # =========================================================================
        # ABA 1: AUDITORIA POR DIA — CONFERÊNCIA COM RUAS (TUDO NUM CARD SÓ)
        # =========================================================================
        with tab_dia:
            if not por_dia:
                st.warning("Nenhuma cobrança de KM foi encontrada nas abas desta planilha de fechamento.")
            else:
                total_dias = len(por_dia)
                total_km_cobrado = sum(float(d["Soma KM Cobrado no Dia"]) for d in por_dia)
                total_km_mapa = sum(float(d["KM Calculado no Mapa"]) for d in por_dia)
                dias_com_glosa = [d for d in por_dia if d["Precisa Glosar"]]
                dias_aprovados = [d for d in por_dia if not d["Precisa Glosar"]]
    
                st.markdown("#### 📅 Auditoria por Dia — Conferência Completa com Ruas (Tudo num Card só)")
                st.caption(
                    "Cada Card abaixo consolida tudo do dia: as **OSs que cobram KM**, a **Quilometragem calculada pela Maior Rota (Google Maps)**, "
                    "o **Veredito**, os botões de rota e a **Planilha completa de todas as OSs do dia com os endereços das ruas visitadas**."
                )
    
                # Métricas Globais no Topo
                m1, m2, m3, m4 = st.columns(4)
                with m1:
                    st.metric("Total de KM Cobrado", f"{total_km_cobrado:,.1f} km", help="Soma de todos os KMs cobrados na planilha de fechamento")
                with m2:
                    st.metric("Total Calculado (Maior Rota)", f"{total_km_mapa:,.1f} km", help="Soma das maiores rotas rodoviárias calculadas pelo Google Maps / OSRM com Base no início e no fim")
                with m3:
                    st.metric("Dias Aprovados (Bateu)", f"{len(dias_aprovados)} dias", delta="✅ 100% Coberto" if len(dias_aprovados) > 0 else None)
                with m4:
                    st.metric("Dias com Alerta de Glosa", f"{len(dias_com_glosa)} dias", delta=f"🚨 {len(dias_com_glosa)}" if len(dias_com_glosa) > 0 else "Nenhuma glosa", delta_color="inverse")
    
                st.markdown("---")
    
                # Filtro de Exibição
                opcao_filtro = st.radio(
                    "Filtrar Cards dos Dias:",
                    options=["Todos os Dias", f"🚨 Apenas com Alerta de Glosa ({len(dias_com_glosa)})", f"✅ Apenas Aprovados ({len(dias_aprovados)})"],
                    horizontal=True,
                    index=0
                )
    
                if "Alerta" in opcao_filtro:
                    dias_exibir = dias_com_glosa
                elif "Aprovados" in opcao_filtro:
                    dias_exibir = dias_aprovados
                else:
                    dias_exibir = por_dia
    
                # RENDERIZAÇÃO DOS CARDS ("TUDO NUM CARD SÓ")
                for d in dias_exibir:
                    soma_km = float(d["Soma KM Cobrado no Dia"])
                    km_mapa = float(d["KM Calculado no Mapa"])
                    dif_km = round(soma_km - km_mapa, 1)
                    is_glosa = d["Precisa Glosar"]
                    todos_trechos = d["Trechos Maps"].get("todos_trechos", [])
                    qtd_total_dia = int(d["Qtd Total OSs no Dia"])
                    lista_oss_dia = d["Todas as OSs do Dia"]
                    oss_com_km = d["Lista OSs Cobrando"]
    
                    margem_aceita = float(d.get("Tolerancia KM", tolerancia_km))
                    if is_glosa:
                        excesso = round(dif_km - margem_aceita, 1)
                        badge_veredito = f'<span class="badge-alto">🚨 EXCEDEU A MARGEM (+{dif_km:,.1f} km acima do Maps)</span>'
                        cor_dif = "#F87171"
                        status_folga = f'<span style="color:#F87171;font-weight:700;">Cobrou +{dif_km:,.1f} km acima do Maps. Ultrapassou a margem de {margem_aceita:,.0f} km em <b>+{excesso:,.1f} km</b>.</span>'
                    else:
                        badge_veredito = f'<span class="badge-normal">✅ APROVADO (Dentro da margem de {margem_aceita:,.0f} km)</span>'
                        cor_dif = "#4ADE80"
                        if dif_km <= 0:
                            status_folga = f'<span style="color:#4ADE80;font-weight:700;">Rota do Maps cobre totalmente a cobrança com folga de {abs(dif_km):,.1f} km.</span>'
                        else:
                            status_folga = f'<span style="color:#4ADE80;font-weight:700;">Cobrou +{dif_km:,.1f} km acima do Maps, mas está <b>APROVADO</b> (dentro da margem de {margem_aceita:,.0f} km).</span>'
    
                    chips_cobrando = []
                    for o_cob in oss_com_km:
                        chips_cobrando.append(f'<b>OS {o_cob["os"]}</b> ({o_cob["km_cobrado"]:,.1f} km - {o_cob["cidade"]})')
                    txt_oss_cobrando = " + ".join(chips_cobrando)
    
                    kms_trechos = d.get("KMs Trechos Lista", [])
                    if len(kms_trechos) > 1:
                        txt_trechos = " + ".join(f"T{i}: {k:,.0f} km" for i, k in enumerate(kms_trechos, 1))
                    elif len(kms_trechos) == 1:
                        txt_trechos = f"Trecho Único: {kms_trechos[0]:,.0f} km"
                    else:
                        txt_trechos = f"{km_mapa:,.0f} km apurado"
    
                    botoes_maps_html = []
                    for idx_t, tr in enumerate(todos_trechos, start=1):
                        nome_btn = f"🌐 Abrir Trecho {idx_t} no Google Maps ({tr['qtd_paradas']} paradas)" if len(todos_trechos) > 1 else f"🌐 Abrir Rota com Ruas no Google Maps ({tr['qtd_paradas']} paradas)"
                        botoes_maps_html.append(
                            f'<a href="{tr["url"]}" target="_blank" class="btn-maps" style="font-size:13px;padding:8px 14px;margin-right:8px;margin-bottom:6px;">{nome_btn}</a>'
                        )
                    url_cidades = d.get("Link Maps (Por Cidades - Sem Erro)")
                    if url_cidades:
                        botoes_maps_html.append(
                            f'<a href="{url_cidades}" target="_blank" class="btn-maps-sec" style="font-size:12px;padding:7px 12px;">🗺️ Abrir Rota Cidades (Auxiliar)</a>'
                        )
    
                    html_linhas_tabela = []
                    for idx_ordem, o_item in enumerate(lista_oss_dia, start=1):
                        cob = o_item["km_cobrado"] > 0
                        r_cls = ' class="row-highlight"' if cob else ''
                        badge_st = f'<span class="badge-cobrando">💰 {o_item["km_cobrado"]:,.1f} km</span>' if cob else '<span class="badge-sem-km">0 km</span>'
                        end_rua_formatado = o_item["endereco_limpo"]
    
                        # Aviso se a OS não constava no Totvs
                        if not o_item.get("encontrada_totvs", True):
                            tag_totvs = '<br><span style="color:#FDE68A;font-size:10.5px;font-weight:700;background:#451A03;padding:1px 5px;border-radius:4px;border:1px solid #B45309;" title="Esta OS foi informada no Fechamento mas não foi localizada no relatório Totvs extraído. O sistema utilizou os dados da própria planilha de fechamento.">⚠️ Não consta no Totvs</span>'
                            end_rua_formatado = f"{end_rua_formatado} <small style='color:#94A3B8;'>(Origem: Fechamento)</small>"
                        else:
                            tag_totvs = ""
    
                        t_num = 1
                        if len(todos_trechos) > 1:
                            t_num = 1 if idx_ordem <= len(todos_trechos[0].get("paradas", [])) else 2
    
                        html_linhas_tabela.append(
                            f'<tr{r_cls}>'
                            f'<td style="text-align:center;font-weight:700;">{idx_ordem}</td>'
                            f'<td style="text-align:center;font-weight:800;font-size:13.5px;color:#38BDF8;">OS {o_item["os"]}{tag_totvs}</td>'
                            f'<td style="text-align:center;">{badge_st}</td>'
                            f'<td style="text-align:left;font-weight:700;color:#F8FAFC;font-size:13px;">🏠 {end_rua_formatado}</td>'
                            f'<td style="text-align:center;font-weight:600;color:#E2E8F0;">{o_item["cidade"]}</td>'
                            f'<td style="text-align:center;color:#94A3B8;">{o_item["cod_cliente"]}</td>'
                            f'<td style="text-align:left;color:#CBD5E1;">{o_item["servico"]}</td>'
                            f'<td style="text-align:center;font-weight:600;color:#38BDF8;">Trecho {t_num}</td>'
                            f'</tr>'
                        )
    
                    card_html = textwrap.dedent(f"""
                    <div class="card-dia">
                        <div class="card-dia-header">
                            <div>
                                <span class="card-dia-data">📅 DIA {d['Data']}</span>
                                <span style="margin-left: 10px; font-size: 13.5px; color: #94A3B8; font-weight:600;">
                                    ({qtd_total_dia} OSs atendidas no roteiro)
                                </span>
                            </div>
                            <div>
                                {badge_veredito}
                            </div>
                        </div>
                        
                        <div class="card-dia-metricas">
                            <div class="metrica-box">
                                <span class="metrica-label">💰 SOMA DE KM COBRADO NO DIA</span>
                                <span class="metrica-val" style="color: #F8FAFC;">{soma_km:,.1f} km</span>
                                <span class="metrica-sub">{txt_oss_cobrando}</span>
                            </div>
                            <div class="metrica-box">
                                <span class="metrica-label">🛣️ MAIOR ROTA POSSÍVEL (MAPS)</span>
                                <span class="metrica-val" style="color: #38BDF8;">{km_mapa:,.1f} km</span>
                                <span class="metrica-sub">({txt_trechos})</span>
                            </div>
                            <div class="metrica-box">
                                <span class="metrica-label">⚖️ DIFERENÇA / VEREDITO</span>
                                <span class="metrica-val" style="color: {cor_dif};">{dif_km:+,.1f} km</span>
                                <span class="metrica-sub">{status_folga}</span>
                            </div>
                        </div>
    
                        <div style="margin-bottom: 14px;">
                            <span style="font-size: 13px; font-weight: 700; color: #F8FAFC; margin-right: 8px;">
                                📍 Conferir Rota das Ruas no Google Maps:
                            </span>
                            <div style="margin-top: 6px;">
                                {"".join(botoes_maps_html)}
                            </div>
                        </div>
    
                        <div style="margin-top: 14px;">
                            <div style="font-weight: 700; color: #38BDF8; font-size: 14px; margin-bottom: 6px;">
                                📋 Planilha de Todas as {qtd_total_dia} OSs do Dia com Endereço das Ruas:
                            </div>
                            <table class="table-km" style="font-size: 12.5px; margin-bottom: 4px;">
                                <thead>
                                    <tr>
                                        <th style="width: 40px;">#</th>
                                        <th style="width: 100px;">OS</th>
                                        <th style="width: 110px;">KM Cobrado</th>
                                        <th>Rua / Endereço (Google Maps)</th>
                                        <th style="width: 130px;">Cidade</th>
                                        <th style="width: 100px;">Cód. Cliente</th>
                                        <th>Serviço Prestado</th>
                                        <th style="width: 80px;">Trecho</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {"".join(html_linhas_tabela)}
                                </tbody>
                            </table>
                        </div>
                    </div>
                    """).strip()
    
                    if hasattr(st, "html"):
                        st.html(card_html)
                    else:
                        st.markdown(card_html, unsafe_allow_html=True)
    
                    with st.expander(f"📸 Conferir Prints ou Ajustar KM dos Trechos do Dia {d['Data']} (Opcional)"):
                        cols_tr = st.columns(len(todos_trechos))
                        km_digitados_card = []
                        for idx_t, t_obj in enumerate(todos_trechos, start=1):
                            with cols_tr[idx_t - 1]:
                                st.markdown(f"**📍 Trecho {idx_t}** ({t_obj.get('qtd_paradas', 0)} paradas)")
                                st.link_button(f"🌐 Abrir Trecho {idx_t} no Maps", t_obj["url"], use_container_width=True)
                                val_t = float(t_obj.get("km_calculado_ruas", 0.0))
                                if val_t <= 0.0:
                                    val_t = float(km_mapa / len(todos_trechos))
                                km_inf = st.number_input(
                                    f"KM Maps Trecho {idx_t}:",
                                    min_value=0.0,
                                    max_value=3000.0,
                                    value=val_t,
                                    step=1.0,
                                    key=f"km_inp_{d['chave_dia']}_{idx_t}"
                                )
                                km_digitados_card.append(km_inf)
                                img_t = st.file_uploader(
                                    f"Print Trecho {idx_t}:",
                                    type=["png", "jpg", "jpeg"],
                                    key=f"img_tr_{d['chave_dia']}_{idx_t}"
                                )
                                if img_t:
                                    st.image(img_t, caption=f"Print Trecho {idx_t}", use_container_width=True)
    
                        soma_ajustada = sum(km_digitados_card)
                        if soma_ajustada != km_mapa:
                            st.info(f"Soma recalculada com os valores informados: **{soma_ajustada:,.1f} km** (Cobrado: {soma_km:,.1f} km)")
    
        # =========================================================================
        # ABA 2: RELATÓRIO DE DESLOCAMENTO (ESTILO PDF CONSIGAZ PRONTO PARA PRINT)
        # =========================================================================
        with tab_relatorio_pdf:
            dossies = res.get("dossies", {})
            if not dossies:
                st.warning("Nenhum dia com cobrança de KM localizado.")
            else:
                lista_opcoes = list(dossies.keys())
                opcao_sel = st.selectbox(
                    "Selecione o Dia para gerar a folha de Relatório de Deslocamento:",
                    options=lista_opcoes,
                    index=0
                )
                dossie = dossies[opcao_sel]
    
                c_up1, c_up2 = st.columns(2)
                with c_up1:
                    km_maps_folha = st.number_input(
                        "KM Apurado no Mapa (já preenchido automaticamente pela rota calculada):",
                        min_value=0.0,
                        max_value=3000.0,
                        value=float(dossie["KM Calculado no Mapa"]),
                        step=1.0,
                        key=f"km_folha_{dossie['chave_dia']}"
                    )
                with c_up2:
                    prints_maps = st.file_uploader(
                        "Anexar Print(s) do Google Maps na folha (Opcional):",
                        type=["png", "jpg", "jpeg"],
                        accept_multiple_files=True,
                        key=f"prints_{dossie['chave_dia']}"
                    )
    
                tag_logo_pdf = f'<img src="data:image/png;base64,{simbolo_b64}" style="height:24px;width:auto;vertical-align:middle;margin-right:6px;" />' if simbolo_b64 else ''
                folha_pdf_html = [
                    '<div style="background-color:#FFFFFF;border:2px solid #CBD5E1;border-radius:8px;padding:22px;margin-top:10px;margin-bottom:20px;font-family:Calibri,Arial,sans-serif;color:#000000;box-shadow:0 4px 14px rgba(0,0,0,0.5);">',
                    '<table style="width:100%;border-collapse:collapse;margin-bottom:12px;border:1.5px solid #000;">',
                    '<tr>',
                    '<td style="width:55%;padding:12px;border:1.5px solid #000;vertical-align:middle;">',
                    f'<span style="display:inline-flex;align-items:center;font-size:24px;font-weight:900;color:#004B87;letter-spacing:1px;">{tag_logo_pdf}CONSIGAZ</span> &nbsp;&nbsp;',
                    '<span style="font-size:22px;font-weight:900;color:#B91C1C;">GASBALL</span> &nbsp;&nbsp;',
                    '<span style="font-size:22px;font-weight:900;color:#15803D;">PROPANGÁS</span>',
                    '</td>',
                    '<td style="width:45%;padding:12px;border:1.5px solid #000;text-align:center;vertical-align:middle;">',
                    '<div style="font-size:24px;font-weight:900;color:#DC2626;line-height:1.1;">RELATÓRIO DE DESLOCAMENTO</div>',
                    f'<div style="font-size:15px;font-weight:800;color:#000000;margin-top:4px;">DATA DO DIA: {dossie["Data"]}</div>',
                    '</td>',
                    '</tr>',
                    '</table>',
                    '<table style="width:100%;border-collapse:collapse;font-size:13px;color:#000000;border:1.5px solid #000;">',
                    '<thead><tr style="background-color:#F1F5F9;font-weight:800;text-align:center;">',
                    '<th style="border:1px solid #000;padding:6px;color:#000;">ABA PLANILHA:</th>',
                    '<th style="border:1px solid #000;padding:6px;color:#000;">ORDEM DE SERVIÇO:</th>',
                    '<th style="border:1px solid #000;padding:6px;color:#000;">CÓDIGO CLIENTE:</th>',
                    '<th style="border:1px solid #000;padding:6px;color:#000;">CIDADE:</th>',
                    '<th style="border:1px solid #000;padding:6px;color:#000;">ENDEREÇO SIMPLIFICADO:</th>',
                    '<th style="border:1px solid #000;padding:6px;color:#000;">KM TOTAL:</th>',
                    '<th style="border:1px solid #000;padding:6px;color:#000;">PEDÁGIO:</th>',
                    '<th style="border:1px solid #000;padding:6px;color:#000;">SERVIÇO PRESTADO:</th>',
                    '</tr></thead><tbody>'
                ]
    
                for o_row in dossie["Todas as OSs do Dia"]:
                    tem_km = o_row["km_cobrado"] > 0
                    bg_row = 'background-color:#FEF3C7;font-weight:700;' if tem_km else 'background-color:#FFFFFF;'
                    km_cell = f'{o_row["km_cobrado"]:,.1f} km' if tem_km else ''
                    ped_cell = f'{o_row["pedagio"]:,.2f}' if o_row["pedagio"] > 0 else ''
                    folha_pdf_html.append(
                        f'<tr style="{bg_row}text-align:center;">'
                        f'<td style="border:1px solid #000;padding:5px;color:#000;">{o_row["aba"]}</td>'
                        f'<td style="border:1px solid #000;padding:5px;color:#000;font-weight:800;">{o_row["os"]}</td>'
                        f'<td style="border:1px solid #000;padding:5px;color:#000;">{o_row["cod_cliente"]}</td>'
                        f'<td style="border:1px solid #000;padding:5px;color:#000;">{o_row["cidade"]}</td>'
                        f'<td style="border:1px solid #000;padding:5px;color:#000;text-align:left;">{o_row["endereco_limpo"]}</td>'
                        f'<td style="border:1px solid #000;padding:5px;color:#000;font-weight:800;">{km_cell}</td>'
                        f'<td style="border:1px solid #000;padding:5px;color:#000;">{ped_cell}</td>'
                        f'<td style="border:1px solid #000;padding:5px;color:#000;text-align:left;">{o_row["servico"]}</td>'
                        '</tr>'
                    )
    
                folha_pdf_html.append(
                    '<tr style="background-color:#E2E8F0;font-weight:900;text-align:center;">'
                    f'<td colspan="5" style="border:1.5px solid #000;padding:7px;text-align:right;color:#000;">SOMA DOS KMs REGISTRADOS NO DIA ({dossie["Data"]}):</td>'
                    f'<td style="border:1.5px solid #000;padding:7px;color:#B91C1C;font-size:15px;">{dossie["Soma KM Cobrado no Dia"]:,.1f} km</td>'
                    f'<td style="border:1.5px solid #000;padding:7px;color:#000;">{dossie["Soma Pedágio no Dia"]:,.2f}</td>'
                    f'<td style="border:1.5px solid #000;padding:7px;text-align:left;color:#000;">Maior KM Calculado da Rota: <b>{km_maps_folha:,.1f} km</b> ({dossie["Sequência da Rota"]})</td>'
                    '</tr>'
                )
                folha_pdf_html.append('</tbody></table>')
                folha_pdf_html.append('</div>')
                st.markdown("\n".join(folha_pdf_html), unsafe_allow_html=True)
    
                if prints_maps:
                    for p_img in prints_maps:
                        st.image(p_img, use_container_width=True)
    
        # =========================================================================
        # ABA 3: TABELA INDIVIDUAL POR OS COBRADA
        # =========================================================================
        with tab_os:
            df_resumo = pd.DataFrame(linhas_os)
            linhas_html = [
                '<table class="table-km">',
                '<thead><tr>',
                '<th>Aba</th>',
                '<th>OS Cobrada</th>',
                '<th>Data</th>',
                '<th>KM na OS</th>',
                '<th>SOMA DE KM DITO NO DIA</th>',
                '<th>MAIOR KM CALCULADO DA ROTA</th>',
                '<th>Total de OSs no Dia</th>',
                '<th>Cidade</th>',
                '<th>Rota por Cidades (Sem Erro)</th>',
                '</tr></thead><tbody>'
            ]
            for _, r in df_resumo.iterrows():
                km_val = float(r["KM Cobrado"])
                soma_d = float(r["Soma KM do Dia"])
                km_m = float(r["KM Calculado no Mapa"])
                is_high = (soma_d - km_m) > float(tolerancia_km)
                row_class = ' class="row-highlight"' if is_high else ''
                badge_os = f'<span class="badge-cobrando">{km_val:,.1f} km</span>'
                badge_dia = f'<span class="badge-alto">⚠️ {soma_d:,.1f} km</span>' if is_high else f'<span class="badge-normal">{soma_d:,.1f} km</span>'
                badge_mp = f'<span class="badge-mapa">🛣️ {km_m:,.1f} km</span>'
                link_maps = r["Link Maps (Cidades)"]
                btn_link = f'<a href="{link_maps}" target="_blank" class="btn-maps">🗺️ Abrir Rota Cidades</a>' if link_maps else '-'
    
                linhas_html.append(
                    f'<tr{row_class}>'
                    f'<td style="text-align:center;font-weight:600;">{r["Aba"]}</td>'
                    f'<td style="text-align:center;font-weight:800;font-size:14px;color:#38BDF8;">OS {r["OS Cobrada"]}</td>'
                    f'<td style="text-align:center;font-weight:700;">{r["Data"]}</td>'
                    f'<td style="text-align:center;">{badge_os}</td>'
                    f'<td style="text-align:center;">{badge_dia}</td>'
                    f'<td style="text-align:center;">{badge_mp}</td>'
                    f'<td style="text-align:center;font-weight:700;">{r["Qtd OS Total Dia"]} OSs</td>'
                    f'<td style="text-align:center;font-weight:600;">{r["Cidade"]}</td>'
                    f'<td style="text-align:center;">{btn_link}</td>'
                    f'</tr>'
                )
            linhas_html.append('</tbody></table>')
            st.markdown("\n".join(linhas_html), unsafe_allow_html=True)
    
        # =========================================================================
        # ABA 4: EXPORTAR EXCEL
        # =========================================================================
        with tab_exportar:
            st.markdown("#### 📥 Baixar Relatório Consolidado por Dia em Excel")
            buffer = io.BytesIO()
            exportar_relatorio_excel(res, buffer)
            buffer.seek(0)
    
            nome_arquivo_excel = f"Auditoria_KM_Por_Dia_{st.session_state.get('nome_fechamento', 'Consigaz')}.xlsx"
            st.download_button(
                label="📥 Baixar Planilha de Auditoria por Dia (.xlsx)",
                data=buffer,
                file_name=nome_arquivo_excel,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("👆 Arraste a **Planilha de Fechamento** e o **Relatório Totvs** nos campos acima e clique em **Calcular KMs por Dia Automaticamente**.")
    



def render_botao_copiar_clipboard(tsv_texto: str):
    tsv_json = json.dumps(tsv_texto)
    html_code = f"""
    <div style="margin-top: 6px; margin-bottom: 10px;">
        <button id="copy-excel-btn" onclick="copiarParaExcel()" style="
            width: 100%;
            background: linear-gradient(135deg, #0284C7 0%, #0369A1 100%);
            color: #FFFFFF;
            border: 1px solid #38BDF8;
            border-radius: 8px;
            padding: 12px 18px;
            font-size: 15px;
            font-weight: 700;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
            box-shadow: 0 4px 14px rgba(2, 132, 199, 0.4);
            transition: all 0.2s ease;
        ">
            📋 Copiar Dados para Colar no Excel (Ctrl + V)
        </button>
        <div id="copy-status" style="display:none; text-align:center; color:#4ADE80; font-size:13px; font-weight:700; margin-top:6px;">
            ✅ Dados copiados para a área de transferência! Vá ao Excel e aperte Ctrl + V
        </div>
    </div>
    <script>
    function copiarParaExcel() {{
        const texto = {tsv_json};
        navigator.clipboard.writeText(texto).then(function() {{
            const btn = document.getElementById('copy-excel-btn');
            const status = document.getElementById('copy-status');
            btn.style.background = 'linear-gradient(135deg, #16A34A 0%, #15803D 100%)';
            btn.style.borderColor = '#4ADE80';
            btn.innerHTML = '✅ Dados Copiados com Sucesso!';
            status.style.display = 'block';
            setTimeout(function() {{
                btn.style.background = 'linear-gradient(135deg, #0284C7 0%, #0369A1 100%)';
                btn.style.borderColor = '#38BDF8';
                btn.innerHTML = '📋 Copiar Dados para Colar no Excel (Ctrl + V)';
                status.style.display = 'none';
            }}, 3500);
        }}).catch(function(err) {{
            alert('Não foi possível copiar automaticamente pelo navegador: ' + err);
        }});
    }}
    </script>
    """
    components.html(html_code, height=75)


def render_fechamento_lote(tag_logo, tag_logo_sidebar):
    st.markdown(
        f"""
        <div style="display:flex;align-items:center;gap:14px;margin-bottom:6px;padding-bottom:4px;">
            {tag_logo}
            <div>
                <div class="main-header">Consigaz — Fechamento em Lote & Leitor de E-mails</div>
                <div class="sub-header">Arraste múltiplos e-mails (.msg) ou planilhas (.xlsx) recebidas para extração instantânea do consolidado, conferência e cópia direta para o Excel</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.sidebar.markdown(
        f"""
        <div style="display:flex;align-items:center;gap:10px;margin-bottom:14px;">
            {tag_logo_sidebar}
            <span style="font-size:18px;font-weight:800;color:#F8FAFC;letter-spacing:0.3px;">Opções do Fechamento</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    incluir_cabecalho = st.sidebar.checkbox(
        "Incluir cabeçalho na cópia (Base, Assunto, Cód...)",
        value=False,
        help="Se marcado, a primeira linha copiada conterá os títulos das colunas."
    )

    with st.sidebar.expander("🔍 Histórico para Comparação (Opcional)", expanded=False):
        arquivo_historico = st.file_uploader(
            "Planilha de Controle Atual (.xlsx)",
            type=["xlsx"],
            help="Carregue sua planilha de controle atual para o sistema avisar se algum desses e-mails já tinha sido lançado anteriormente."
        )

    st.sidebar.markdown("---")
    st.sidebar.markdown(
        """
        <div style="text-align: center; padding: 14px 10px; background-color: #1E293B; border-radius: 8px; border: 1.5px solid #334155; margin-top: 18px; box-shadow: 0 4px 12px rgba(0,0,0,0.3);">
            <span style="font-size: 16px; font-weight: 800; color: #38BDF8; letter-spacing: 0.5px;">Desenvolvido por Chiku</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("### 📥 1. Carregar E-mails ou Planilhas em Lote")
    st.caption("💡 **Dica:** Você pode ir soltando e-mails aos poucos (de 3 em 3 ou 4 em 4) que o sistema **vai acumulando todos na listinha** sem apagar os anteriores! Ou selecione vários de uma pasta para subir dezenas de uma só vez.")

    if "fechamento_arquivos_acumulados" not in st.session_state:
        st.session_state["fechamento_arquivos_acumulados"] = {}

    arquivos_upload = st.file_uploader(
        "Arraste e solte arquivos aqui (.xlsx, .xlsm ou .msg do Outlook):",
        type=["xlsx", "xlsm", "msg"],
        accept_multiple_files=True,
        key="upload_lote_fechamento",
        help="Você pode arrastar diretamente anexos ou e-mails salvos"
    )

    # Acumula novos arquivos soltos sem substituir os anteriores
    if arquivos_upload:
        for arq in arquivos_upload:
            if arq.name not in st.session_state["fechamento_arquivos_acumulados"]:
                st.session_state["fechamento_arquivos_acumulados"][arq.name] = {
                    "nome": arq.name,
                    "size": arq.size,
                    "bytes": arq.getvalue()
                }

    arquivos_acumulados = list(st.session_state["fechamento_arquivos_acumulados"].values())

    if arquivos_acumulados:
        # Exibe cabeçalho da listinha com botão de limpar
        col_list_header, col_list_action = st.columns([3, 1])
        with col_list_header:
            st.markdown(f"#### 📋 Lista de Arquivos Adicionados ({len(arquivos_acumulados)} arquivo{'s' if len(arquivos_acumulados) != 1 else ''}):")
        with col_list_action:
            if st.button("🗑️ Limpar Lista de Arquivos", use_container_width=True, help="Zera a lista de arquivos para iniciar um novo fechamento"):
                st.session_state["fechamento_arquivos_acumulados"] = {}
                st.rerun()

        # Renderiza a Listinha Vertical com numeração, ícone e tamanho
        linhas_lista = []
        for idx, arq_info in enumerate(arquivos_acumulados, 1):
            tam_kb = arq_info["size"] / 1024
            is_msg = arq_info["nome"].upper().endswith(".MSG")
            badge_tipo = "✉️ E-mail" if is_msg else "📊 Excel"
            cor_badge = "#0284C7" if is_msg else "#10B981"
            linhas_lista.append(
                f'''<div class="item-lista-arquivo">
                    <div style="display:flex;align-items:center;gap:12px;">
                        <span style="font-weight:800;color:#38BDF8;font-size:13px;width:26px;">#{idx}</span>
                        <span style="background-color:{cor_badge};color:#FFFFFF;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:700;">{badge_tipo}</span>
                        <span style="font-weight:600;color:#F8FAFC;font-size:14px;">{arq_info['nome']}</span>
                    </div>
                    <span style="color:#94A3B8;font-size:12.5px;font-weight:600;">{tam_kb:,.1f} KB</span>
                </div>'''
            )
        st.markdown("\n".join(linhas_lista), unsafe_allow_html=True)
        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

        with st.spinner("Processando arquivos, aba CONSOLIDADO, fornecedores e valores..."):
            chaves_hist = extrair_chaves_historico(arquivo_historico) if arquivo_historico else set()
            registros = processar_lote(arquivos_acumulados, chaves_hist)

        total_arquivos = len(registros)
        valor_total = sum(r.get("valor", 0.0) or 0.0 for r in registros if not str(r.get("status", "")).startswith("ERRO"))
        duplicados = [r for r in registros if "DUPLICADO" in str(r.get("status", "")).upper()]
        erros = [r for r in registros if "ERRO" in str(r.get("status", "")).upper()]
        ok_count = total_arquivos - len(duplicados) - len(erros)

        col_k1, col_k2, col_k3, col_k4 = st.columns(4)
        with col_k1:
            st.metric("Total Processado", f"{total_arquivos} itens")
        with col_k2:
            st.metric("Valor Total Consolidado", formatar_moeda_br(valor_total))
        with col_k3:
            st.metric("Prontos / OK", f"{ok_count} itens")
        with col_k4:
            st.metric("Alertas / Atenção", f"{len(duplicados) + len(erros)} itens")

        if duplicados:
            st.warning(f"⚠️ Atenção: Detectamos **{len(duplicados)} fechamento(s) duplicado(s)** no lote ou histórico. Verifique antes de colar no Excel.")
        if erros:
            st.error(f"🚨 Existem **{len(erros)} arquivo(s) com erro** de leitura ou sem aba CONSOLIDADO.")

        st.markdown("---")
        st.markdown("### 📊 2. Grade de Fechamento (Visualização e Conferência)")

        col_f1, col_f2 = st.columns(2)
        with col_f1:
            bases_disp = sorted(list(set(r.get("base", "N/D") for r in registros)))
            filtro_bases = st.multiselect("Filtrar por Base:", options=bases_disp, default=bases_disp)
        with col_f2:
            status_disp = sorted(list(set(r.get("status", "") for r in registros)))
            filtro_status = st.multiselect("Filtrar por Status:", options=status_disp, default=status_disp)

        registros_filtrados = [
            r for r in registros
            if r.get("base") in filtro_bases and r.get("status") in filtro_status
        ]

        if not registros_filtrados:
            st.info("Nenhum registro corresponde aos filtros selecionados acima.")
        else:
            df_view = pd.DataFrame([
                {
                    "Base": r.get("base", ""),
                    "Identificação / Assunto": r.get("identificacao", ""),
                    "Código / CNPJ": r.get("codigo_fornecedor", ""),
                    "Fornecedor": r.get("fornecedor", ""),
                    "Valor (R$)": r.get("valor", 0.0),
                    "Status": r.get("status", "PENDENTE"),
                    "Data": r.get("data", ""),
                    "Arquivo Original": r.get("arquivo", ""),
                    "Avisos / Detalhes": r.get("detalhe_erro", "")
                }
                for r in registros_filtrados
            ])

            st.dataframe(
                df_view,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Valor (R$)": st.column_config.NumberColumn(
                        "Valor (R$)",
                        format="R$ %.2f"
                    ),
                    "Base": st.column_config.TextColumn("Base", width="small"),
                    "Código / CNPJ": st.column_config.TextColumn("Código / CNPJ", width="small"),
                    "Status": st.column_config.TextColumn("Status", width="medium"),
                }
            )

            st.markdown("### 📋 3. Copiar e Colar no Excel ou Baixar")
            tsv_texto = gerar_tsv_colar_excel(registros_filtrados, incluir_cabecalho=incluir_cabecalho)

            col_btn1, col_btn2 = st.columns([1.2, 1])
            with col_btn1:
                render_botao_copiar_clipboard(tsv_texto)
            with col_btn2:
                excel_bytes = gerar_excel_fechamento(registros_filtrados)
                nome_dl = f"Fechamento_Consolidado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                st.download_button(
                    label="📥 Baixar Planilha Consolidada (.xlsx)",
                    data=excel_bytes,
                    file_name=nome_dl,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            with st.expander("👁️ Ver texto tabulado para cópia manual (Ctrl+A e Ctrl+C)", expanded=False):
                st.caption("Você também pode selecionar o texto abaixo e copiar diretamente para a planilha do Excel:")
                st.code(tsv_texto, language=None)
    else:
        st.info("👆 Arraste e solte seus e-mails (.msg) ou planilhas (.xlsx) no campo acima para carregar a grade automaticamente.")



st.sidebar.markdown(
    f"""
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px;">
        {tag_logo_sidebar}
        <span style="font-size:16px;font-weight:800;color:#38BDF8;letter-spacing:0.5px;">CONSIGAZ TOOLS</span>
    </div>
    """,
    unsafe_allow_html=True
)

modulo_ativo = st.sidebar.radio(
    "Selecione a Ferramenta:",
    options=["🚗 Auditoria de KM por Dia", "✉️ Fechamento por E-mails / Lote"],
    index=0
)
st.sidebar.markdown("---")

if modulo_ativo == "🚗 Auditoria de KM por Dia":
    render_auditoria_km(tag_logo, tag_logo_sidebar)
else:
    render_fechamento_lote(tag_logo, tag_logo_sidebar)

st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: #94A3B8; font-size: 14px; font-weight: 700; padding: 22px 0;">
        Consigaz &bull; Gestão & Automação &bull; Desenvolvido por Chiku
    </div>
    """,
    unsafe_allow_html=True
)
