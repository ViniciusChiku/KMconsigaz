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
    page_title="Consigaz - Auditoria de KM de OSs",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS para identidade visual limpa e profissional
st.markdown("""
<style>
    .main-header {
        font-size: 26px;
        font-weight: 700;
        color: #1F4E79;
        margin-bottom: 2px;
    }
    .sub-header {
        font-size: 15px;
        color: #595959;
        margin-bottom: 20px;
    }
    .card-metric {
        background-color: #F8F9FA;
        border: 1px solid #E9ECEF;
        border-radius: 8px;
        padding: 15px;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-value {
        font-size: 24px;
        font-weight: 700;
        color: #1F4E79;
    }
    .metric-label {
        font-size: 13px;
        color: #6C757D;
    }
    .stButton>button {
        background-color: #1F4E79;
        color: white;
        font-weight: 600;
        border-radius: 6px;
        height: 44px;
    }
    .stButton>button:hover {
        background-color: #163857;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# Título Principal
st.markdown('<div class="main-header">🔥 Consigaz — Auditoria de KM e Deslocamento</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Conferência inteligente de rotas com cruzamento do Relatório Totvs e Google Maps</div>', unsafe_allow_html=True)

# Barra Lateral - Configurações
st.sidebar.header("⚙️ Parâmetros da Auditoria")

lista_labels = list(OPCOES_BASES_FORMATADAS.keys())
indice_padrao = 0
for idx, lbl in enumerate(lista_labels):
    if "MAU" in lbl or "Base 8" in lbl:
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

senha_planilha = st.sidebar.text_input(
    "Senha da Planilha (se houver):",
    value="fechamento",
    type="password",
    help="Senha padrão de fechamento utilizada para desbloquear os arquivos."
)

st.sidebar.markdown("---")
st.sidebar.caption("📌 **Dica**: O sistema percorre todas as abas com OS e KM (Moto, Carro, Engenharia, etc.) e ignora pedágio.")

# Área de Upload
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Planilha de Fechamento")
    arquivo_fechamento = st.file_uploader(
        "Arraste a planilha enviada pela base (.xlsx)",
        type=["xlsx", "xlsm"],
        key="upload_fechamento",
        help="Exemplo: G.IGÁS-MAU-OUT-2026-1-CORTE-RELIGUE.xlsx"
    )
    if arquivo_fechamento:
        base_detectada = detectar_base_por_texto(arquivo_fechamento.name)
        st.success(f"Arquivo carregado: **{arquivo_fechamento.name}** (Base detectada: **{base_detectada}**)")

with col2:
    st.subheader("2. Relatório Totvs (ESOS)")
    arquivo_totvs = st.file_uploader(
        "Arraste a planilha de consulta Totvs com as OSs (.xlsx)",
        type=["xlsx", "xlsm"],
        key="upload_totvs",
        help="Exemplo: maua (2).xlsx com a aba ESOS0301..."
    )
    if arquivo_totvs:
        st.success(f"Arquivo carregado: **{arquivo_totvs.name}**")

st.markdown("---")

if arquivo_fechamento and arquivo_totvs:
    if st.button("🚀 Processar e Auditar Quilometragem", use_container_width=True):
        with st.spinner("Desbloqueando arquivos, cruzando OSs e calculando rotas..."):
            try:
                # 1. Carrega fechamento
                wb_fech = carregar_planilha_segura(arquivo_fechamento, senha=senha_planilha)
                
                # 2. Carrega Totvs
                wb_totvs = carregar_planilha_segura(arquivo_totvs, senha=senha_planilha)

                # 3. Executa auditoria
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
    linhas = res["resumo"]
    detalhes = res["detalhes"]

    st.markdown("### 📊 Indicadores Consolidados")
    
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        st.markdown(f"""
        <div class="card-metric">
            <div class="metric-value">{res['total_oss_km']}</div>
            <div class="metric-label">OSs com KM Cobrado</div>
        </div>
        """, unsafe_allow_html=True)
    with m2:
        st.markdown(f"""
        <div class="card-metric">
            <div class="metric-value">{res['total_km_cobrado']:,.1f} km</div>
            <div class="metric-label">KM Total Cobrado</div>
        </div>
        """, unsafe_allow_html=True)
    with m3:
        st.markdown(f"""
        <div class="card-metric">
            <div class="metric-value">{res['total_km_faturavel']:,.1f} km</div>
            <div class="metric-label">KM Faturável (Pós-Raio {res['raio_franquia']:.0f}km)</div>
        </div>
        """, unsafe_allow_html=True)
    with m4:
        st.markdown(f"""
        <div class="card-metric">
            <div class="metric-value">{len(res['abas_auditadas'])}</div>
            <div class="metric-label">Abas de Fechamento Auditadas</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Abas de visualização
    tab_dossie, tab_resumo, tab_detalhes, tab_exportar = st.tabs([
        "📑 Dossiê Visual da Rota (Estilo Print)",
        "📋 Tabela de Auditoria por OS",
        "📍 Detalhamento de Visitas do Dia",
        "📥 Baixar Relatório em Excel"
    ])

    with tab_dossie:
        dossies = res.get("dossies", {})
        if not dossies:
            st.warning("Nenhum dia com cobrança de KM localizado no fechamento.")
        else:
            lista_opcoes = list(dossies.keys())
            opcao_sel = st.selectbox(
                "Selecione o Dia / OS para gerar o Dossiê Visual:",
                options=lista_opcoes,
                index=0,
                help="Escolha o dia para visualizar a folha de conferência formatada."
            )

            dossie = dossies[opcao_sel]
            trechos = dossie["trechos"]
            t1 = trechos["trecho1"]
            t2 = trechos.get("trecho2")

            st.markdown("---")
            st.markdown("#### 🗺️ Conferência dos Trechos no Google Maps")

            c_tr1, c_tr2 = st.columns(2)
            with c_tr1:
                st.markdown(f"**{t1['nome']}**")
                st.link_button(
                    "🌐 Abrir Trecho 1 no Google Maps",
                    t1["url"],
                    use_container_width=True
                )
                km_t1 = st.number_input(
                    "KM Apurado no Maps (Trecho 1):",
                    min_value=0.0,
                    max_value=1500.0,
                    value=float(dossie["km_cobrado"] if not t2 else dossie["km_cobrado"] / 2),
                    step=1.0,
                    key=f"km_t1_{opcao_sel}"
                )
                img_t1 = st.file_uploader(
                    "Anexar print do Trecho 1 (Opcional):",
                    type=["png", "jpg", "jpeg"],
                    key=f"img_t1_{opcao_sel}"
                )

            with c_tr2:
                if t2:
                    st.markdown(f"**{t2['nome']}**")
                    st.link_button(
                        "🌐 Abrir Trecho 2 no Google Maps",
                        t2["url"],
                        use_container_width=True
                    )
                    km_t2 = st.number_input(
                        "KM Apurado no Maps (Trecho 2):",
                        min_value=0.0,
                        max_value=1500.0,
                        value=float(dossie["km_cobrado"] / 2),
                        step=1.0,
                        key=f"km_t2_{opcao_sel}"
                    )
                    img_t2 = st.file_uploader(
                        "Anexar print do Trecho 2 (Opcional):",
                        type=["png", "jpg", "jpeg"],
                        key=f"img_t2_{opcao_sel}"
                    )
                else:
                    km_t2 = 0.0
                    img_t2 = None
                    st.info("ℹ️ Este percurso possui 8 ou menos paradas e coube inteiramente em 1 único trecho.")

            total_km_apurado = km_t1 + km_t2

            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("### 📸 Folha de Conferência da Rota (Pronta para Print)")
            st.caption("Você pode tirar print diretamente desta folha abaixo para arquivar ou anexar no fechamento.")

            # Container formatado idêntico ao modelo da imagem
            st.markdown("""
            <div style="background-color: white; border: 2px solid #1F4E79; border-radius: 8px; padding: 25px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); margin-bottom: 25px;">
                <h2 style="text-align: center; color: black; font-weight: 900; letter-spacing: 2px; margin-bottom: 2px;">ROTA DE KM</h2>
                <h4 style="text-align: center; color: #333; margin-top: 0px; font-weight: 600; margin-bottom: 20px;">""" + dossie['data_formatada'] + """</h4>
            """, unsafe_allow_html=True)

            # Tabela de OSs no estilo da imagem do usuário
            df_folha = pd.DataFrame(dossie["tabela_oss"])[["Ordem Serviço", "Cliente", "Cidade", "Data", "Serviço"]]
            
            html_tabela = """
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; font-family: Calibri, Arial, sans-serif; font-size: 13px; color: black;">
                <thead>
                    <tr style="background-color: #F2F4F7; border: 1.5px solid #000; text-align: center; font-weight: bold;">
                        <th style="border: 1.5px solid #000; padding: 8px;">Ordem Serviço</th>
                        <th style="border: 1.5px solid #000; padding: 8px;">Cliente</th>
                        <th style="border: 1.5px solid #000; padding: 8px;">Cidade</th>
                        <th style="border: 1.5px solid #000; padding: 8px;">Data</th>
                        <th style="border: 1.5px solid #000; padding: 8px;">Serviço</th>
                    </tr>
                </thead>
                <tbody>
            """
            for _, row in df_folha.iterrows():
                html_tabela += f"""
                    <tr style="border: 1px solid #444; text-align: center;">
                        <td style="border: 1px solid #444; padding: 6px; font-weight: bold;">{row['Ordem Serviço']}</td>
                        <td style="border: 1px solid #444; padding: 6px; text-align: left;">{row['Cliente']}</td>
                        <td style="border: 1px solid #444; padding: 6px;">{row['Cidade']}</td>
                        <td style="border: 1px solid #444; padding: 6px;">{row['Data']}</td>
                        <td style="border: 1px solid #444; padding: 6px;">{row['Serviço']}</td>
                    </tr>
                """
            html_tabela += """
                </tbody>
            </table>
            """
            st.markdown(html_tabela, unsafe_allow_html=True)

            # Destaque do TOTAL DE KM em letras garrafais idêntico ao modelo
            st.markdown(f"""
            <div style="text-align: center; font-size: 26px; font-weight: 900; color: black; margin-top: 15px; margin-bottom: 25px; letter-spacing: 1px;">
                TOTAL DE {total_km_apurado:,.0f} KM
            </div>
            """, unsafe_allow_html=True)

            # Imagens dos prints
            if img_t1:
                st.image(img_t1, caption=f"Trecho 1 — {km_t1:,.0f} km", use_container_width=True)
            if img_t2:
                st.image(img_t2, caption=f"Trecho 2 — {km_t2:,.0f} km", use_container_width=True)

            # Quadro de Confronto e Auditoria
            km_cobrado = dossie["km_cobrado"]
            franquia = dossie["raio_franquia"]
            km_fat_cobrado = dossie["km_faturavel"]
            diferenca = total_km_apurado - km_cobrado

            if total_km_apurado >= km_cobrado:
                status_box = f"""
                <div style="background-color: #E8F5E9; border: 2px solid #4CAF50; border-radius: 6px; padding: 15px; margin-top: 20px;">
                    <h4 style="color: #2E7D32; margin: 0 0 5px 0;">✅ ROTA VALIDADA NO GOOGLE MAPS</h4>
                    <p style="margin: 0; color: #1B5E20; font-size: 14px;">
                        A rota apurada no Google Maps (<b>{total_km_apurado:,.1f} km</b>) cobre a quilometragem afirmada na planilha da base (<b>{km_cobrado:,.1f} km</b>).<br>
                        <b>KM Faturável Aprovado:</b> {max(0.0, total_km_apurado - franquia):,.1f} km (com franquia deduzida de {franquia:.0f} km).
                    </p>
                </div>
                """
            else:
                status_box = f"""
                <div style="background-color: #FFEBEE; border: 2px solid #E53935; border-radius: 6px; padding: 15px; margin-top: 20px;">
                    <h4 style="color: #C62828; margin: 0 0 5px 0;">⚠️ DIVERGÊNCIA IDENTIFICADA</h4>
                    <p style="margin: 0; color: #B71C1C; font-size: 14px;">
                        A rota no Google Maps somou apenas <b>{total_km_apurado:,.1f} km</b>, ficando <b>{abs(diferenca):,.1f} km abaixo</b> da quilometragem afirmada pela base (<b>{km_cobrado:,.1f} km</b>).<br>
                        <b>KM Faturável Real do Maps:</b> {max(0.0, total_km_apurado - franquia):,.1f} km vs <b>Cobrado:</b> {km_fat_cobrado:,.1f} km.
                    </p>
                </div>
                """
            st.markdown(status_box, unsafe_allow_html=True)
            st.markdown("</div>", unsafe_allow_html=True)

    with tab_resumo:
        df_resumo = pd.DataFrame(linhas)
        
        # Filtros
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            abas_unicas = ["Todas"] + list(df_resumo["Aba"].unique())
            aba_filtro = st.selectbox("Filtrar por Aba:", abas_unicas)
        with col_f2:
            datas_unicas = ["Todas"] + list(df_resumo["Data"].unique())
            data_filtro = st.selectbox("Filtrar por Data:", datas_unicas)

        df_filtrado = df_resumo.copy()
        if aba_filtro != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Aba"] == aba_filtro]
        if data_filtro != "Todas":
            df_filtrado = df_filtrado[df_filtrado["Data"] == data_filtro]

        # Tabela com Links Clicáveis no Streamlit
        st.dataframe(
            df_filtrado[[
                "Aba", "OS Cobrada", "Data", "Técnico da OS",
                "KM Cobrado", "KM Faturável Cobrado", "Qtd OS Técnico", "Qtd OS Total Dia",
                "Link Maps (Técnico)", "Link Maps (Geral do Dia)", "Endereço OS"
            ]],
            column_config={
                "OS Cobrada": st.column_config.NumberColumn(format="%d"),
                "KM Cobrado": st.column_config.NumberColumn(format="%.1f km"),
                "KM Faturável Cobrado": st.column_config.NumberColumn(format="%.1f km"),
                "Link Maps (Técnico)": st.column_config.LinkColumn(
                    "Rota do Técnico (Maps)",
                    display_text="🗺️ Abrir Rota Técnico"
                ),
                "Link Maps (Geral do Dia)": st.column_config.LinkColumn(
                    "Rota Geral (Maps)",
                    display_text="🗺️ Abrir Rota Geral"
                ),
            },
            hide_index=True,
            use_container_width=True
        )

        st.caption("ℹ️ Clique no botão **🗺️ Abrir Rota Técnico** ou **🗺️ Abrir Rota Geral** para abrir o Google Maps traçado no navegador.")

    with tab_detalhes:
        st.markdown("#### Lista Completa de Atendimentos por Dia")
        if detalhes:
            df_det = pd.DataFrame(detalhes)
            
            # Filtro por OS cobrada
            oss_cobradas = list(df_det["OS Âncora Fechamento"].unique())
            os_sel = st.selectbox("Selecione a OS Âncora do Fechamento:", oss_cobradas)
            
            df_det_filtrado = df_det[df_det["OS Âncora Fechamento"] == os_sel]
            
            st.dataframe(
                df_det_filtrado[[
                    "Ordem", "OS Atendida", "Técnico", "Hora",
                    "Cliente", "Endereço Limpo", "Bairro", "Cidade", "Situação"
                ]],
                column_config={
                    "OS Atendida": st.column_config.NumberColumn(format="%d"),
                    "Ordem": st.column_config.NumberColumn(format="%dº"),
                },
                hide_index=True,
                use_container_width=True
            )
        else:
            st.info("Nenhum detalhe disponível.")

    with tab_exportar:
        st.markdown("#### Exportar Relatório Oficial com Links Hipertexto")
        st.write("Baixe a planilha Excel contendo o resumo com fórmulas `=HYPERLINK` clicáveis e a aba detalhada com cada cliente visitado.")

        buffer = io.BytesIO()
        exportar_relatorio_excel(res, buffer)
        buffer.seek(0)

        nome_arquivo_excel = f"Auditoria_KM_{st.session_state.get('nome_fechamento', 'Consigaz')}.xlsx"
        
        st.download_button(
            label="📥 Baixar Planilha de Auditoria (.xlsx)",
            data=buffer,
            file_name=nome_arquivo_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👆 Por favor, carregue a **Planilha de Fechamento** e a **Planilha Totvs** nos campos acima e clique em **Processar e Auditar Quilometragem**.")
