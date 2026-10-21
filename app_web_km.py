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
    font-weight: 600;
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
    background-color: #0D9488 !important;
    color: #FFFFFF !important;
    padding: 5px 9px;
    border-radius: 5px;
    text-decoration: none !important;
    font-weight: 600;
    font-size: 11.5px;
    display: inline-block;
    margin: 2px;
}
</style>
""", unsafe_allow_html=True)

# Título Principal
st.markdown('<div class="main-header">🔥 Consigaz — Auditoria de KM por Dia e Rota Completa</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Soma todos os KMs cobrados no mesmo dia, junta todas as OSs da data e joga todos os endereços no Google Maps</div>', unsafe_allow_html=True)

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
    help="Selecione a base correspondente para carregar o endereço oficial de partida e retorno."
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
    step=5.0,
    help="Quilometragem deduzida do percurso (ex: 20 km raio ida e volta = 40 km)."
)

limite_alerta_km = st.sidebar.number_input(
    "Destacar Soma de KM do Dia acima de (KM):",
    min_value=20.0,
    max_value=1000.0,
    value=100.0,
    step=10.0,
    help="Dias cuja soma de KM cobrado atingir este valor (ou que tiverem mais de 1 OS cobrando KM no mesmo dia) serão destacados em amarelo."
)

senha_planilha = st.sidebar.text_input(
    "Senha da Planilha (se houver):",
    value="fechamento",
    type="password",
    help="Senha padrão de fechamento utilizada para desbloquear os arquivos."
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
    if st.button("🚀 Processar Soma de KMs por Dia e Gerar Rotas", use_container_width=True):
        with st.spinner("Cruzando todas as OSs do fechamento com o Totvs, agrupando por dia e montando rotas..."):
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
        "📅 1. Soma de KMs por Dia + Todas as OSs do Dia",
        "📄 2. Relatório de Deslocamento (Estilo PDF Consigaz)",
        "📋 3. Tabela Individual por OS Cobrada",
        "📥 4. Baixar Relatório em Excel"
    ])

    # =========================================================================
    # ABA 1: SOMA DE KMs POR DIA + TODAS AS OSs E ENDEREÇOS DO DIA
    # =========================================================================
    with tab_dia:
        if not por_dia:
            st.warning("Nenhuma cobrança de KM foi encontrada nas abas desta planilha de fechamento.")
        else:
            st.markdown("#### 📅 Resumo Consolidado por Dia (Soma dos KMs Registrados no Dia)")
            st.caption(
                "Cada linha abaixo reúne **todas as OSs feitas no mesmo dia** e **soma todos os KMs cobrados naquela data**. "
                "Clique nos botões azuis para jogar **todos os endereços do dia no Google Maps** e ver se a rota bate com o KM total cobrado no dia."
            )

            html_dias = [
                '<table class="table-km">',
                '<thead><tr>',
                '<th>Data do Dia</th>',
                '<th>OSs que Estão Cobrando KM no Dia (Valor por OS)</th>',
                '<th>SOMA DE KM DITO NO DIA</th>',
                '<th>Total de OSs no Mesmo Dia</th>',
                '<th>Cidades Visitadas no Dia</th>',
                '<th>Jogar Todos os Endereços do Dia no Google Maps</th>',
                '</tr></thead><tbody>'
            ]

            for d in por_dia:
                soma_km = float(d["Soma KM Cobrado no Dia"])
                qtd_cobrando = int(d["Qtd OSs Cobrando"])
                qtd_total_dia = int(d["Qtd Total OSs no Dia"])
                qtd_enderecos = int(d["Qtd Endereços Distintos"])

                # Destaca se a soma do dia for alta OU se tiver mais de 1 OS cobrando KM no mesmo dia
                is_alert = (soma_km >= float(limite_alerta_km)) or (qtd_cobrando > 1)
                row_cls = ' class="row-highlight"' if is_alert else ''

                # Badges das OSs cobrando KM naquele dia
                chips_cobrando = []
                for o_cob in d["Lista OSs Cobrando"]:
                    chips_cobrando.append(
                        f'<span class="badge-cobrando">💰 OS {o_cob["os"]}: {o_cob["km_cobrado"]:,.1f} km ({o_cob["cidade"]} - {o_cob["aba"]})</span>'
                    )
                html_oss_cobrando = "<br>".join(chips_cobrando)

                if is_alert:
                    alerta_multi = f"<br><small style='color:#B91C1C;font-weight:700;'>⚡ {qtd_cobrando} OSs cobrando no mesmo dia!</small>" if qtd_cobrando > 1 else ""
                    badge_soma = f'<span class="badge-alto">⚠️ {soma_km:,.1f} km</span>{alerta_multi}'
                else:
                    badge_soma = f'<span class="badge-normal">{soma_km:,.1f} km</span>'

                # Botões para abrir todos os trechos do dia no Google Maps
                todos_trechos = d["Trechos Maps"].get("todos_trechos", [])
                botoes_maps = []
                if len(todos_trechos) <= 1:
                    url_u = d["Link Maps (Todas do Dia)"]
                    botoes_maps.append(f'<a href="{url_u}" target="_blank" class="btn-maps">🗺️ Abrir Rota Completa ({qtd_enderecos} endereços)</a>')
                else:
                    for idx_t, tr in enumerate(todos_trechos, start=1):
                        botoes_maps.append(f'<a href="{tr["url"]}" target="_blank" class="btn-maps">🗺️ Abrir Trecho {idx_t} ({tr["qtd_paradas"]} paradas)</a>')

                if qtd_total_dia > qtd_cobrando:
                    url_so_km = d["Link Maps (Só OSs com KM)"]
                    botoes_maps.append(f'<br><a href="{url_so_km}" target="_blank" class="btn-maps-sec">🎯 Rota Só OSs Cobradas ({qtd_cobrando})</a>')

                html_dias.append(
                    f'<tr{row_cls}>'
                    f'<td style="text-align:center;font-weight:800;font-size:14.5px;">{d["Data"]}</td>'
                    f'<td style="text-align:left;">{html_oss_cobrando}</td>'
                    f'<td style="text-align:center;">{badge_soma}</td>'
                    f'<td style="text-align:center;font-weight:700;">{qtd_total_dia} OSs no dia<br><small style="color:#475569;">({qtd_cobrando} c/ KM e {qtd_total_dia - qtd_cobrando} s/ KM)</small></td>'
                    f'<td style="text-align:left;font-weight:600;">{d["Cidades do Dia"]}</td>'
                    f'<td style="text-align:center;">{"".join(botoes_maps)}</td>'
                    f'</tr>'
                )

            html_dias.append('</tbody></table>')
            st.markdown("\n".join(html_dias), unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("#### 🔍 Conferir Todas as OSs, Endereços e Quilometragem de um Dia Específico")

            opcoes_dias_det = [
                f"{d['Data']} — Soma Cobrada: {d['Soma KM Cobrado no Dia']:,.1f} km ({d['Qtd OSs Cobrando']} OS cobrando de {d['Qtd Total OSs no Dia']} OSs no dia)"
                for d in por_dia
            ]
            idx_dia_sel = st.selectbox(
                "Escolha o Dia abaixo para ver todas as OSs do dia juntas, os endereços enviados ao Google Maps e simular a soma:",
                options=range(len(opcoes_dias_det)),
                format_func=lambda i: opcoes_dias_det[i],
                index=0
            )

            dia_atual = por_dia[idx_dia_sel]
            trechos_atuais = dia_atual["Trechos Maps"].get("todos_trechos", [])

            # Calculadora rápida de confronto: Soma Cobrada no Dia vs Google Maps
            c_calc1, c_calc2 = st.columns([1.2, 1.8])
            with c_calc1:
                st.markdown(f"**🗺️ Links do Google Maps para o dia {dia_atual['Data']}**")
                st.caption(f"Partida e Retorno: `{res['base_endereco']}`")
                soma_maps_dia = 0.0
                for idx_t, tr in enumerate(trechos_atuais, start=1):
                    st.link_button(
                        f"🌐 Abrir {tr['nome']} no Google Maps",
                        tr["url"],
                        use_container_width=True
                    )
                    val_default = float(dia_atual["Soma KM Cobrado no Dia"] / len(trechos_atuais))
                    km_digitado = st.number_input(
                        f"KM que deu no Google Maps (Trecho {idx_t}):",
                        min_value=0.0,
                        max_value=3000.0,
                        value=val_default,
                        step=1.0,
                        key=f"calc_km_dia_{dia_atual['chave_dia']}_{idx_t}"
                    )
                    soma_maps_dia += km_digitado

            with c_calc2:
                soma_cobrada = float(dia_atual["Soma KM Cobrado no Dia"])
                dif_km = soma_maps_dia - soma_cobrada
                if soma_maps_dia >= soma_cobrada:
                    box_veredito = (
                        '<div style="background-color:#ECFDF5;border:2px solid #10B981;border-radius:8px;padding:18px;margin-top:8px;">'
                        f'<h4 style="color:#065F46 !important;margin:0 0 8px 0;">✅ A SOMA DO GOOGLE MAPS BATEU COM O KM COBRADO NO DIA ({dia_atual["Data"]})</h4>'
                        f'<p style="margin:0;color:#064E3B !important;font-size:14.5px;line-height:1.5;">'
                        f'• <b>Soma dos KMs cobrados pelas OSs no dia:</b> {soma_cobrada:,.1f} km ({dia_atual["OSs Cobrando KM"]})<br>'
                        f'• <b>KM Total da Rota no Google Maps (com todas as {dia_atual["Qtd Total OSs no Dia"]} OSs do dia):</b> {soma_maps_dia:,.1f} km<br>'
                        f'• <b>Conclusão:</b> O trajeto completo do dia no Google Maps ({soma_maps_dia:,.1f} km) cobre os {soma_cobrada:,.1f} km cobrados.'
                        '</p></div>'
                    )
                else:
                    box_veredito = (
                        '<div style="background-color:#FEF2F2;border:2px solid #EF4444;border-radius:8px;padding:18px;margin-top:8px;">'
                        f'<h4 style="color:#991B1B !important;margin:0 0 8px 0;">🚨 A ROTA NÃO BATE! COBRANÇA MAIOR QUE O GOOGLE MAPS ({dia_atual["Data"]})</h4>'
                        f'<p style="margin:0;color:#7F1D1D !important;font-size:14.5px;line-height:1.5;">'
                        f'• <b>Soma dos KMs cobrados pelas OSs no dia:</b> <b>{soma_cobrada:,.1f} km</b> ({dia_atual["OSs Cobrando KM"]})<br>'
                        f'• <b>KM Total no Google Maps (juntando todas as {dia_atual["Qtd Total OSs no Dia"]} OSs do dia):</b> <b>{soma_maps_dia:,.1f} km</b><br>'
                        f'• <b>Diferença Cobrada a Maior (Glosa sugerida):</b> <span style="background:#DC2626;color:#FFF;padding:2px 8px;border-radius:4px;font-weight:800;">{abs(dif_km):,.1f} km a mais sem justificativa</span>'
                        '</p></div>'
                    )
                st.markdown(box_veredito, unsafe_allow_html=True)

                # Resumo rápido da ordem dos endereços enviados ao Google Maps
                paradas_html = [
                    '<div style="background-color:#F8FAFC;border:1px solid #CBD5E1;border-radius:8px;padding:12px 16px;margin-top:12px;">',
                    f'<div style="font-weight:700;color:#1F4E79;margin-bottom:6px;">📍 Sequência de Endereços no Google Maps ({len(dia_atual["Paradas Google Maps"])} paradas entre Partida e Retorno):</div>',
                    f'<div style="font-size:13px;color:#1E293B;margin-bottom:4px;">🟢 <b>Partida (Base):</b> {res["base_endereco"]}</div>'
                ]
                for p in dia_atual["Paradas Google Maps"]:
                    oss_p_txt = ", ".join(
                        f"<b>OS {x['os']} (💰 {x['km_cobrado']:,.1f} km)</b>" if x["km_cobrado"] > 0 else f"OS {x['os']}"
                        for x in p["oss"]
                    )
                    icone_p = "🟠" if p["tem_cobranca"] else "⚪"
                    paradas_html.append(
                        f'<div style="font-size:13px;color:#1E293B;margin-bottom:4px;">'
                        f'{icone_p} <b>Parada {p["ordem"]}:</b> {p["endereco"]} — <span style="color:#0F172A;">[{oss_p_txt}]</span>'
                        f'</div>'
                    )
                paradas_html.append(f'<div style="font-size:13px;color:#1E293B;margin-top:4px;">🔴 <b>Retorno (Base):</b> {res["base_endereco"]}</div>')
                paradas_html.append('</div>')
                st.markdown("\n".join(paradas_html), unsafe_allow_html=True)

            # Tabela com TODAS as OSs daquele dia juntas (mostrando quem está cobrando KM e quem tem 0 km)
            st.markdown(f"##### 📋 Todas as {dia_atual['Qtd Total OSs no Dia']} OSs do Dia {dia_atual['Data']} (Com Endereços e Soma de KM)")
            html_oss_dia = [
                '<table class="table-km">',
                '<thead><tr>',
                '<th>Aba da Planilha</th>',
                '<th>Ordem de Serviço</th>',
                '<th>Status KM</th>',
                '<th>KM Registrado na OS</th>',
                '<th>Cód. Cliente</th>',
                '<th>Cidade</th>',
                '<th>Endereço Completo (Enviado ao Google Maps)</th>',
                '<th>Serviço Prestado</th>',
                '</tr></thead><tbody>'
            ]

            for o_item in dia_atual["Todas as OSs do Dia"]:
                cob = o_item["km_cobrado"] > 0
                r_cls = ' class="row-highlight"' if cob else ''
                badge_st = f'<span class="badge-cobrando">💰 COBRANDO KM</span>' if cob else '<span class="badge-sem-km">Mesmo Dia (0 km)</span>'
                km_txt = f'<b>{o_item["km_cobrado"]:,.1f} km</b>' if cob else '0,0 km'

                html_oss_dia.append(
                    f'<tr{r_cls}>'
                    f'<td style="text-align:center;font-weight:600;">{o_item["aba"]}</td>'
                    f'<td style="text-align:center;font-weight:800;font-size:14px;">{o_item["os"]}</td>'
                    f'<td style="text-align:center;">{badge_st}</td>'
                    f'<td style="text-align:center;font-size:14px;">{km_txt}</td>'
                    f'<td style="text-align:center;">{o_item["cod_cliente"]}</td>'
                    f'<td style="text-align:center;font-weight:600;">{o_item["cidade"]}</td>'
                    f'<td style="text-align:left;font-weight:600;">{o_item["endereco_limpo"]}</td>'
                    f'<td style="text-align:left;">{o_item["servico"]}</td>'
                    f'</tr>'
                )

            html_oss_dia.append(
                '<tr class="row-total">'
                f'<td colspan="3" style="text-align:right;">SOMA TOTAL DE KM REGISTRADO NO DIA ({dia_atual["Data"]}):</td>'
                f'<td style="text-align:center;">{dia_atual["Soma KM Cobrado no Dia"]:,.1f} km</td>'
                f'<td colspan="4" style="text-align:left;">Total de {dia_atual["Qtd Total OSs no Dia"]} OSs no dia ({dia_atual["Qtd OSs Cobrando"]} OSs cobrando KM)</td>'
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
                "Selecione o Dia para gerar a folha de Relatório de Deslocamento (modelo PDF Consigaz):",
                options=lista_opcoes,
                index=0
            )
            dossie = dossies[opcao_sel]
            trechos_d = dossie["Trechos Maps"].get("todos_trechos", [])

            c_up1, c_up2 = st.columns(2)
            with c_up1:
                km_maps_folha = st.number_input(
                    "KM Apurado no Google Maps para exibir na folha:",
                    min_value=0.0,
                    max_value=3000.0,
                    value=float(dossie["Soma KM Cobrado no Dia"]),
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
                '<th style="border:1px solid #000;padding:6px;color:#000;">ENDEREÇO (GOOGLE MAPS):</th>',
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
                f'<td style="border:1.5px solid #000;padding:7px;text-align:left;color:#000;"> KM Apurado Google Maps: <b>{km_maps_folha:,.1f} km</b></td>'
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
        st.caption("Lista individual de cada linha onde foi lançado KM > 0, mostrando ao lado a **Soma Total de KM daquele mesmo dia**.")
        linhas_html = [
            '<table class="table-km">',
            '<thead><tr>',
            '<th>Aba</th>',
            '<th>OS Cobrada</th>',
            '<th>Data</th>',
            '<th>KM Registrado na OS</th>',
            '<th>SOMA DE KM NO MESMO DIA</th>',
            '<th>Total de OSs no Dia</th>',
            '<th>Cidade</th>',
            '<th>Endereço Limpo (Google Maps)</th>',
            '<th>Rota Completa do Dia</th>',
            '</tr></thead><tbody>'
        ]
        for _, r in df_resumo.iterrows():
            km_val = float(r["KM Cobrado"])
            soma_d = float(r["Soma KM do Dia"])
            is_high = soma_d >= float(limite_alerta_km) or km_val >= float(limite_alerta_km)
            row_class = ' class="row-highlight"' if is_high else ''
            badge_os = f'<span class="badge-cobrando">{km_val:,.1f} km</span>'
            badge_dia = f'<span class="badge-alto">⚠️ {soma_d:,.1f} km</span>' if is_high else f'<span class="badge-normal">{soma_d:,.1f} km</span>'
            link_maps = r["Link Maps (Geral do Dia)"]
            btn_link = f'<a href="{link_maps}" target="_blank" class="btn-maps">🗺️ Abrir Rota do Dia</a>' if link_maps else '-'

            linhas_html.append(
                f'<tr{row_class}>'
                f'<td style="text-align:center;font-weight:600;">{r["Aba"]}</td>'
                f'<td style="text-align:center;font-weight:800;font-size:14px;">{r["OS Cobrada"]}</td>'
                f'<td style="text-align:center;font-weight:700;">{r["Data"]}</td>'
                f'<td style="text-align:center;">{badge_os}</td>'
                f'<td style="text-align:center;">{badge_dia}</td>'
                f'<td style="text-align:center;font-weight:700;">{r["Qtd OS Total Dia"]} OSs</td>'
                f'<td style="text-align:center;font-weight:600;">{r["Cidade"]}</td>'
                f'<td>{r["Endereço OS"]}</td>'
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
        st.write("A planilha gerada contém a aba **Soma de KM por Dia** (com os links diretos do Google Maps) e a aba **Relatorio Deslocamento Dia** com todas as OSs do mesmo dia agrupadas e os endereços completos.")

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
    st.info("👆 Arraste a **Planilha de Fechamento** e o **Relatório Totvs** nos campos acima e clique em **Processar Soma de KMs por Dia e Gerar Rotas**.")
