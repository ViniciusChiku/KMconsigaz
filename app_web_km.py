import io
import os
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

raio_franquia = st.sidebar.number_input(
    "Franquia Contratual / Raio (KM):",
    min_value=0.0,
    max_value=200.0,
    value=40.0,
    step=5.0
)

tolerancia_km = st.sidebar.number_input(
    "Tolerância de Diferença de Rota (KM):",
    min_value=0.0,
    max_value=100.0,
    value=20.0,
    step=5.0,
    help="Se o KM cobrado no dia passar do KM Calculado no Mapa mais esta tolerância, a linha será destacada em amarelo/vermelho para glosa."
)

senha_planilha = st.sidebar.text_input(
    "Senha da Planilha (se houver):",
    value="fechamento",
    type="password"
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
                    raio_franquia=float(raio_franquia)
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
    # ABA 1: SOMA DE KMs POR DIA + KM CALCULADO AUTOMATICAMENTE NA TELA
    # =========================================================================
    with tab_dia:
        if not por_dia:
            st.warning("Nenhuma cobrança de KM foi encontrada nas abas desta planilha de fechamento.")
        else:
            st.markdown("#### 📅 Comparativo Automático por Dia: Soma de KM Cobrado vs KM Calculado da Rota")
            st.caption(
                "O sistema já calculou automaticamente a **Quilometragem da Rota do Dia** (estrada entre as cidades + deslocamento entre as OSs do dia) "
                "e comparou com a **Soma de KM Dito no Dia**. O botão **🗺️ Abrir Rota por Cidades** abre o Google Maps direto pelas cidades na ordem certa, **sem risco de falha de endereço**."
            )

            html_dias = [
                '<table class="table-km">',
                '<thead><tr>',
                '<th>Data do Dia</th>',
                '<th>OSs Cobrando KM no Dia</th>',
                '<th>SOMA DE KM DITO NO DIA</th>',
                '<th>KM CALCULADO DA ROTA (MAPA)</th>',
                '<th>VEREDITO / DIFERENÇA</th>',
                '<th>Total de OSs e Rota de Cidades do Dia</th>',
                '<th>Conferir no Google Maps (Sem Erro)</th>',
                '</tr></thead><tbody>'
            ]

            for d in por_dia:
                soma_km = float(d["Soma KM Cobrado no Dia"])
                km_mapa = float(d["KM Calculado no Mapa"])
                km_estrada = float(d["KM Estrada Mapa"])
                km_urb = float(d["KM Urbano Mapa"])
                dif_km = round(soma_km - km_mapa, 1)
                qtd_cobrando = int(d["Qtd OSs Cobrando"])
                qtd_total_dia = int(d["Qtd Total OSs no Dia"])
                qtd_enderecos = int(d["Qtd Endereços Distintos"])

                # Verifica se o KM cobrado passa do KM Calculado no Mapa + tolerância
                is_glosa = dif_km > float(tolerancia_km)
                row_cls = ' class="row-highlight"' if is_glosa else ''

                chips_cobrando = []
                for o_cob in d["Lista OSs Cobrando"]:
                    chips_cobrando.append(
                        f'<span class="badge-cobrando">💰 OS {o_cob["os"]}: {o_cob["km_cobrado"]:,.1f} km ({o_cob["cidade"]})</span>'
                    )
                html_oss_cobrando = "<br>".join(chips_cobrando)

                alerta_multi = f"<br><small style='color:#B91C1C;font-weight:700;'>⚡ {qtd_cobrando} OSs cobrando no mesmo dia!</small>" if qtd_cobrando > 1 else ""
                badge_soma = f'<span style="font-size:15px;font-weight:800;color:#0F172A;">{soma_km:,.1f} km</span>{alerta_multi}'

                badge_mapa = (
                    f'<span class="badge-mapa">🛣️ {km_mapa:,.1f} km</span>'
                    f'<br><small style="color:#475569;">({km_estrada:,.0f} km estrada + {km_urb:,.0f} km nas {qtd_enderecos} OSs)</small>'
                )

                if is_glosa:
                    badge_veredito = (
                        f'<span class="badge-alto">🚨 +{dif_km:,.1f} km ACIMA DO MAPA</span>'
                        f'<br><small style="color:#991B1B;font-weight:700;">Revisar / Glosar excesso</small>'
                    )
                else:
                    badge_veredito = (
                        f'<span class="badge-normal">✅ BATEU COM A ROTA</span>'
                        f'<br><small style="color:#166534;font-weight:600;">Rota cobre os {soma_km:,.1f} km</small>'
                    )

                url_cidades = d["Link Maps (Por Cidades - Sem Erro)"]
                botoes_maps = [
                    f'<a href="{url_cidades}" target="_blank" class="btn-maps">🗺️ Abrir Rota por Cidades (Sem Erro)</a>'
                ]
                todos_trechos = d["Trechos Maps"].get("todos_trechos", [])
                if len(todos_trechos) == 1:
                    botoes_maps.append(f'<br><a href="{todos_trechos[0]["url"]}" target="_blank" class="btn-maps-sec">🏠 Abrir c/ Ruas ({qtd_enderecos})</a>')
                else:
                    links_r = " ".join(
                        f'<a href="{tr["url"]}" target="_blank" class="btn-maps-sec">🏠 Ruas T{idx_t}</a>'
                        for idx_t, tr in enumerate(todos_trechos, start=1)
                    )
                    botoes_maps.append(f'<br>{links_r}')

                html_dias.append(
                    f'<tr{row_cls}>'
                    f'<td style="text-align:center;font-weight:800;font-size:14.5px;">{d["Data"]}</td>'
                    f'<td style="text-align:left;">{html_oss_cobrando}</td>'
                    f'<td style="text-align:center;">{badge_soma}</td>'
                    f'<td style="text-align:center;">{badge_mapa}</td>'
                    f'<td style="text-align:center;">{badge_veredito}</td>'
                    f'<td style="text-align:left;"><b>{qtd_total_dia} OSs no dia:</b><br><span style="font-size:12.5px;color:#1F4E79;font-weight:600;">{d["Sequência da Rota"]}</span></td>'
                    f'<td style="text-align:center;">{"".join(botoes_maps)}</td>'
                    f'</tr>'
                )

            html_dias.append('</tbody></table>')
            st.markdown("\n".join(html_dias), unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("#### 🔍 Ver Todas as OSs e Endereços de um Dia Específico")

            opcoes_dias_det = [
                f"{d['Data']} — Cobrado no Dia: {d['Soma KM Cobrado no Dia']:,.1f} km | Calculado pela Rota: {d['KM Calculado no Mapa']:,.1f} km ({d['Qtd Total OSs no Dia']} OSs)"
                for d in por_dia
            ]
            idx_dia_sel = st.selectbox(
                "Escolha o Dia abaixo para inspecionar todas as OSs do dia juntas:",
                options=range(len(opcoes_dias_det)),
                format_func=lambda i: opcoes_dias_det[i],
                index=0
            )

            dia_atual = por_dia[idx_dia_sel]

            st.markdown(
                f"**🛣️ Rota Calculada ({dia_atual['Data']}):** `{dia_atual['Sequência da Rota']}` — "
                f"**KM Estrada:** `{dia_atual['KM Estrada Mapa']:,.1f} km` + "
                f"**Deslocamento Urbano ({dia_atual['Qtd Endereços Distintos']} endereços):** `{dia_atual['KM Urbano Mapa']:,.1f} km` = "
                f"**Total Calculado: `{dia_atual['KM Calculado no Mapa']:,.1f} km`** (vs **`{dia_atual['Soma KM Cobrado no Dia']:,.1f} km`** cobrados no dia)"
            )

            html_oss_dia = [
                '<table class="table-km">',
                '<thead><tr>',
                '<th>Aba da Planilha</th>',
                '<th>Ordem de Serviço</th>',
                '<th>Status KM</th>',
                '<th>KM Registrado na OS</th>',
                '<th>Cód. Cliente</th>',
                '<th>Cidade</th>',
                '<th>Endereço Simplificado (Sem Bairro)</th>',
                '<th>Serviço Prestado</th>',
                '</tr></thead><tbody>'
            ]

            for o_item in dia_atual["Todas as OSs do Dia"]:
                cob = o_item["km_cobrado"] > 0
                r_cls = ' class="row-highlight"' if cob else ''
                badge_st = '<span class="badge-cobrando">💰 COBRANDO KM</span>' if cob else '<span class="badge-sem-km">Mesmo Dia (0 km)</span>'
                km_txt = f'<b>{o_item["km_cobrado"]:,.1f} km</b>' if cob else '0,0 km'

                html_oss_dia.append(
                    f'<tr{r_cls}>'
                    f'<td style="text-align:center;font-weight:600;">{o_item["aba"]}</td>'
                    f'<td style="text-align:center;font-weight:800;font-size:14px;">{o_item["os"]}</td>'
                    f'<td style="text-align:center;">{badge_st}</td>'
                    f'<td style="text-align:center;font-size:14px;">{km_txt}</td>'
                    f'<td style="text-align:center;">{o_item["cod_cliente"]}</td>'
                    f'<td style="text-align:center;font-weight:700;">{o_item["cidade"]}</td>'
                    f'<td style="text-align:left;font-weight:600;">{o_item["endereco_limpo"]}</td>'
                    f'<td style="text-align:left;">{o_item["servico"]}</td>'
                    f'</tr>'
                )

            html_oss_dia.append(
                '<tr class="row-total">'
                f'<td colspan="3" style="text-align:right;">SOMA DE KM DITO NO DIA ({dia_atual["Data"]}):</td>'
                f'<td style="text-align:center;">{dia_atual["Soma KM Cobrado no Dia"]:,.1f} km</td>'
                f'<td colspan="4" style="text-align:left;">KM Calculado da Rota no Mapa: {dia_atual["KM Calculado no Mapa"]:,.1f} km ({dia_atual["Qtd Total OSs no Dia"]} OSs no dia)</td>'
                '</tr>'
            )
            html_oss_dia.append('</tbody></table>')
            st.markdown("\n".join(html_oss_dia), unsafe_allow_html=True)

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
