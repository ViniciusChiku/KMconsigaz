import io
import os
import textwrap
import streamlit as st
import pandas as pd
from datetime import datetime

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

st.set_page_config(
    page_title="Consigaz - Auditoria de KM por Dia",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS completo para garantir Modo Claro perfeito e sem textos brancos invisíveis
st.markdown("""
<style>
.stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"], [data-testid="stToolbar"] {
    background-color: #FFFFFF !important;
    color: #1E293B !important;
}
[data-testid="stSidebar"], [data-testid="stSidebarContent"] {
    background-color: #F8FAFC !important;
    color: #1E293B !important;
}
label, [data-testid="stWidgetLabel"] p, [data-testid="stMarkdownContainer"] p, [data-testid="stCaptionContainer"] p, div[data-baseweb="select"] span {
    color: #1E293B !important;
}
h1, h2, h3, h4, h5, h6 {
    color: #1F4E79 !important;
}
button[data-baseweb="tab"] p {
    color: #475569 !important;
    font-weight: 600 !important;
    font-size: 15px !important;
}
button[data-baseweb="tab"][aria-selected="true"] p {
    color: #1F4E79 !important;
    font-weight: 700 !important;
}
.stButton>button {
    background-color: #1F4E79 !important;
    color: #FFFFFF !important;
    font-weight: 600 !important;
    border-radius: 6px !important;
    height: 44px !important;
    border: none !important;
}
.stButton>button p, .stButton>button span {
    color: #FFFFFF !important;
}
.stButton>button:hover {
    background-color: #163857 !important;
    color: #FFFFFF !important;
}
.main-header {
    font-size: 26px;
    font-weight: 800;
    color: #1F4E79 !important;
    margin-bottom: 2px;
}
.sub-header {
    font-size: 15px;
    color: #475569 !important;
    margin-bottom: 18px;
}
.table-km {
    width: 100%;
    border-collapse: collapse;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    font-size: 13.5px;
    margin-top: 8px;
    margin-bottom: 18px;
    background-color: #FFFFFF;
}
.table-km th {
    background-color: #1F4E79 !important;
    color: #FFFFFF !important;
    text-align: center;
    padding: 10px 8px;
    font-weight: 600;
    border: 1px solid #163857;
}
.table-km td {
    padding: 9px 10px;
    border: 1px solid #CBD5E1;
    color: #1E293B !important;
    vertical-align: middle;
}
.table-km tr:nth-child(even) {
    background-color: #F8FAFC;
}
.table-km tr:hover {
    background-color: #E2E8F0;
}
.row-highlight {
    background-color: #FEF3C7 !important;
}
.row-highlight td {
    background-color: #FEF3C7 !important;
    color: #78350F !important;
}
.row-total td {
    background-color: #1F4E79 !important;
    color: #FFFFFF !important;
    font-weight: 800 !important;
    font-size: 14px !important;
}
.badge-alto {
    background-color: #DC2626 !important;
    color: #FFFFFF !important;
    padding: 4px 10px;
    border-radius: 12px;
    font-size: 12.5px;
    font-weight: 700;
    display: inline-block;
}
.badge-normal {
    background-color: #16A34A !important;
    color: #FFFFFF !important;
    padding: 4px 10px;
    border-radius: 12px;
    font-size: 12.5px;
    font-weight: 700;
    display: inline-block;
}
.badge-mapa {
    background-color: #0284C7 !important;
    color: #FFFFFF !important;
    padding: 4px 10px;
    border-radius: 12px;
    font-size: 13px;
    font-weight: 700;
    display: inline-block;
}
.badge-cobrando {
    background-color: #D97706 !important;
    color: #FFFFFF !important;
    padding: 3px 8px;
    border-radius: 6px;
    font-size: 12px;
    font-weight: 700;
    display: inline-block;
    margin: 2px 0;
}
.badge-sem-km {
    background-color: #E2E8F0 !important;
    color: #475569 !important;
    padding: 2px 7px;
    border-radius: 6px;
    font-size: 11.5px;
    font-weight: 600;
    display: inline-block;
}
.btn-maps {
    background-color: #2563EB !important;
    color: #FFFFFF !important;
    padding: 6px 11px;
    border-radius: 5px;
    text-decoration: none !important;
    font-weight: 600;
    font-size: 12px;
    display: inline-block;
    margin: 2px;
}
.btn-maps:hover {
    background-color: #1D4ED8 !important;
    color: #FFFFFF !important;
}
.btn-maps-sec {
    background-color: #475569 !important;
    color: #FFFFFF !important;
    padding: 5px 9px;
    border-radius: 5px;
    text-decoration: none !important;
    font-weight: 600;
    font-size: 11px;
    display: inline-block;
    margin: 2px;
}
.card-dia {
    background-color: #FFFFFF !important;
    border: 1.5px solid #CBD5E1 !important;
    border-radius: 12px !important;
    padding: 20px 22px !important;
    margin-bottom: 26px !important;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04) !important;
}
.card-dia:hover {
    border-color: #94A3B8 !important;
    box-shadow: 0 4px 14px rgba(0, 0, 0, 0.08) !important;
}
.card-dia-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-bottom: 2px solid #E2E8F0;
    padding-bottom: 12px;
    margin-bottom: 14px;
    flex-wrap: wrap;
    gap: 10px;
}
.card-dia-data {
    font-size: 19px;
    font-weight: 800;
    color: #1F4E79 !important;
}
.card-dia-metricas {
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    background-color: #F8FAFC !important;
    border-radius: 8px;
    padding: 14px 18px;
    margin-bottom: 16px;
    border: 1px solid #E2E8F0;
}
.metrica-box {
    flex: 1 1 200px;
    display: flex;
    flex-direction: column;
}
.metrica-label {
    font-size: 11px;
    font-weight: 700;
    color: #64748B !important;
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
    color: #475569 !important;
    margin-top: 2px;
    line-height: 1.4;
}
</style>
""", unsafe_allow_html=True)

# Título Principal
st.markdown('<div class="main-header">🔥 Consigaz — Auditoria Automática de KM por Dia</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Soma todos os KMs ditos no dia, calcula automaticamente quantos KMs deu a rota das OSs do dia e mostra se bate ou se há cobrança a maior</div>', unsafe_allow_html=True)

# Barra Lateral - Configurações
st.sidebar.header("⚙️ Parâmetros da Auditoria")

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
    <div style="text-align: center; padding: 14px 10px; background-color: #FFFFFF; border-radius: 8px; border: 1.5px solid #CBD5E1; margin-top: 18px; box-shadow: 0 1px 3px rgba(0,0,0,0.03);">
        <span style="font-size: 15px; font-weight: 800; color: #1F4E79; letter-spacing: 0.3px;">Desenvolvido por Chiku</span>
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
                "Cada Card abaixo consolida tudo do dia: as **OSs que cobram KM**, a **Quilometragem calculada pelas Ruas (Google Maps)**, "
                "o **Veredito**, os botões de rota e a **Planilha completa de todas as OSs do dia com os endereços das ruas visitadas**."
            )

            # Métricas Globais no Topo
            m1, m2, m3, m4 = st.columns(4)
            with m1:
                st.metric("Total de KM Cobrado", f"{total_km_cobrado:,.1f} km", help="Soma de todos os KMs cobrados na planilha de fechamento")
            with m2:
                st.metric("Total Calculado pelas Ruas", f"{total_km_mapa:,.1f} km", help="Soma das rotas rodoviárias calculadas pelo Google Maps / OSRM")
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
                    cor_dif = "#DC2626"
                    status_folga = f'<span style="color:#DC2626;font-weight:700;">Cobrou +{dif_km:,.1f} km acima do Maps. Ultrapassou a margem de {margem_aceita:,.0f} km em <b>+{excesso:,.1f} km</b>.</span>'
                else:
                    badge_veredito = f'<span class="badge-normal">✅ APROVADO (Dentro da margem de {margem_aceita:,.0f} km)</span>'
                    cor_dif = "#16A34A"
                    if dif_km <= 0:
                        status_folga = f'<span style="color:#16A34A;font-weight:700;">Rota do Maps cobre totalmente a cobrança com folga de {abs(dif_km):,.1f} km.</span>'
                    else:
                        status_folga = f'<span style="color:#16A34A;font-weight:700;">Cobrou +{dif_km:,.1f} km acima do Maps, mas está <b>APROVADO</b> (dentro da margem de {margem_aceita:,.0f} km).</span>'

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
                        tag_totvs = '<br><span style="color:#B45309;font-size:10.5px;font-weight:700;background:#FEF3C7;padding:1px 5px;border-radius:4px;border:1px solid #FCD34D;" title="Esta OS foi informada no Fechamento mas não foi localizada no relatório Totvs extraído. O sistema utilizou os dados da própria planilha de fechamento.">⚠️ Não consta no Totvs</span>'
                        end_rua_formatado = f"{end_rua_formatado} <small style='color:#64748B;'>(Origem: Fechamento)</small>"
                    else:
                        tag_totvs = ""

                    t_num = 1
                    if len(todos_trechos) > 1:
                        t_num = 1 if idx_ordem <= len(todos_trechos[0].get("paradas", [])) else 2

                    html_linhas_tabela.append(
                        f'<tr{r_cls}>'
                        f'<td style="text-align:center;font-weight:700;">{idx_ordem}</td>'
                        f'<td style="text-align:center;font-weight:800;font-size:13.5px;color:#1F4E79;">OS {o_item["os"]}{tag_totvs}</td>'
                        f'<td style="text-align:center;">{badge_st}</td>'
                        f'<td style="text-align:left;font-weight:700;color:#0F172A;font-size:13px;">🏠 {end_rua_formatado}</td>'
                        f'<td style="text-align:center;font-weight:600;color:#1E293B;">{o_item["cidade"]}</td>'
                        f'<td style="text-align:center;color:#475569;">{o_item["cod_cliente"]}</td>'
                        f'<td style="text-align:left;color:#334155;">{o_item["servico"]}</td>'
                        f'<td style="text-align:center;font-weight:600;color:#0284C7;">Trecho {t_num}</td>'
                        f'</tr>'
                    )

                card_html = textwrap.dedent(f"""
                <div class="card-dia">
                    <div class="card-dia-header">
                        <div>
                            <span class="card-dia-data">📅 DIA {d['Data']}</span>
                            <span style="margin-left: 10px; font-size: 13.5px; color: #64748B; font-weight:600;">
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
                            <span class="metrica-val" style="color: #0F172A;">{soma_km:,.1f} km</span>
                            <span class="metrica-sub">{txt_oss_cobrando}</span>
                        </div>
                        <div class="metrica-box">
                            <span class="metrica-label">🛣️ KM CALCULADO PELAS RUAS (MAPS)</span>
                            <span class="metrica-val" style="color: #0284C7;">{km_mapa:,.1f} km</span>
                            <span class="metrica-sub">({txt_trechos})</span>
                        </div>
                        <div class="metrica-box">
                            <span class="metrica-label">⚖️ DIFERENÇA / VEREDITO</span>
                            <span class="metrica-val" style="color: {cor_dif};">{dif_km:+,.1f} km</span>
                            <span class="metrica-sub">{status_folga}</span>
                        </div>
                    </div>

                    <div style="margin-bottom: 14px;">
                        <span style="font-size: 13px; font-weight: 700; color: #1E293B; margin-right: 8px;">
                            📍 Conferir Rota das Ruas no Google Maps:
                        </span>
                        <div style="margin-top: 6px;">
                            {"".join(botoes_maps_html)}
                        </div>
                    </div>

                    <div style="margin-top: 14px;">
                        <div style="font-weight: 700; color: #1F4E79; font-size: 14px; margin-bottom: 6px;">
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

            folha_pdf_html = [
                '<div style="background-color:#FFFFFF;border:2px solid #000000;padding:20px;margin-top:10px;margin-bottom:20px;font-family:Calibri,Arial,sans-serif;color:#000000;">',
                '<table style="width:100%;border-collapse:collapse;margin-bottom:12px;border:1.5px solid #000;">',
                '<tr>',
                '<td style="width:55%;padding:12px;border:1.5px solid #000;vertical-align:middle;">',
                '<span style="font-size:24px;font-weight:900;color:#004B87;letter-spacing:1px;">🔥 CONSIGAZ</span> &nbsp;&nbsp;',
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
                f'<td style="border:1.5px solid #000;padding:7px;text-align:left;color:#000;">KM Calculado da Rota: <b>{km_maps_folha:,.1f} km</b> ({dossie["Sequência da Rota"]})</td>'
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
            '<th>KM CALCULADO DA ROTA</th>',
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
                f'<td style="text-align:center;font-weight:800;font-size:14px;">{r["OS Cobrada"]}</td>'
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

st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: #64748B; font-size: 14px; font-weight: 700; padding: 18px 0;">
        Consigaz &bull; Auditoria de KM &bull; Desenvolvido por Chiku
    </div>
    """,
    unsafe_allow_html=True
)
