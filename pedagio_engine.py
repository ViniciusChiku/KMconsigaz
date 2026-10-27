"""
Engine de Auditoria de Pedágios - Consigaz
Cruza os pedágios cobrados na Planilha de Fechamento (.xlsx / .msg / .eml)
com o Relatório de Pedágios em PDF (tabela de consolidação por data + fotos dos comprovantes)
e com as rotas rodoviárias do dia.
"""

import io
import os
import re
import zlib
import base64
import tempfile
import itertools
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from auditar_engine import (
    obter_info_base,
    detectar_base_por_texto,
    carregar_planilha_segura,
    extrair_todas_oss_fechamento,
    carregar_dados_totvs,
    calcular_rota_osrm_por_cidades,
    gerar_url_google_maps
)
from fechamento_engine import (
    encontrar_base,
    formatar_moeda_br
)

try:
    import extract_msg
except ImportError:
    extract_msg = None


def extrair_dados_pdf_pedagios(pdf_bytes: bytes) -> dict:
    """
    Extrai em Python puro (sem depender de bibliotecas externas):
    1. A tabela 'Consolidação por data' da Página 1 (Data, Quantidade, Total R$)
    2. Os contadores gerais (Comprovantes Lançados, Fotos Anexadas, Total Geral)
    3. As imagens JPEG dos comprovantes embutidas nas páginas do PDF
    """
    streams = re.findall(b'stream[\r\n]+(.*?)[\r\n]+endstream', pdf_bytes, re.S)

    cmap = {}
    decompressed_streams = []
    imagens_jpeg = []

    for s in streams:
        s_strip = s.lstrip(b'\r\n ')
        if s_strip.startswith(b'\xff\xd8'):
            imagens_jpeg.append(s_strip)
            continue
        try:
            dec = zlib.decompress(s)
            if dec.lstrip(b'\r\n ').startswith(b'\xff\xd8'):
                imagens_jpeg.append(dec.lstrip(b'\r\n '))
                continue
            txt = dec.decode('latin1', errors='ignore')
            decompressed_streams.append(txt)
            if 'begincmap' in txt:
                for bfc in re.findall(r'beginbfchar(.*?)endbfchar', txt, re.S):
                    for m in re.finditer(r'<([0-9A-Fa-f]{4})>\s*<([0-9A-Fa-f]{4})>', bfc):
                        cid = int(m.group(1), 16)
                        uni = int(m.group(2), 16)
                        cmap[cid] = chr(uni)
                for bfr in re.findall(r'beginbfrange(.*?)endbfrange', txt, re.S):
                    for m in re.finditer(r'<([0-9A-Fa-f]{4})>\s*<([0-9A-Fa-f]{4})>\s*<([0-9A-Fa-f]{4})>', bfr):
                        c_start = int(m.group(1), 16)
                        c_end = int(m.group(2), 16)
                        u_start = int(m.group(3), 16)
                        for offset in range(c_end - c_start + 1):
                            cmap[c_start + offset] = chr(u_start + offset)
        except Exception:
            pass

    def decodificar_hex_pdf(hex_str: str) -> str:
        chars = []
        for i in range(0, len(hex_str), 4):
            bloco = hex_str[i:i+4]
            if len(bloco) == 4:
                val = int(bloco, 16)
                if val in cmap:
                    chars.append(cmap[val])
                elif 3 <= val <= 120:
                    chars.append(chr(val + 29))
                else:
                    chars.append(chr(val))
        return "".join(chars)

    linhas_texto = []
    for txt in decompressed_streams:
        if 'begincmap' in txt:
            continue
        blocos_bt = re.findall(r'BT(.*?)ET', txt, re.S)
        for b in blocos_bt:
            partes = []
            for m in re.finditer(r'<([0-9A-Fa-f]+)>\s*Tj|\((.*?)\)\s*Tj|\[(.*?)\]\s*TJ', b, re.S):
                if m.group(1):
                    partes.append(decodificar_hex_pdf(m.group(1)))
                elif m.group(2):
                    partes.append(m.group(2))
                elif m.group(3):
                    sub = m.group(3)
                    for sm in re.finditer(r'<([0-9A-Fa-f]+)>|\((.*?)\)', sub):
                        if sm.group(1):
                            partes.append(decodificar_hex_pdf(sm.group(1)))
                        elif sm.group(2):
                            partes.append(sm.group(2))
            texto_bloco = "".join(partes).strip().replace("\xa0", " ")
            if texto_bloco:
                linhas_texto.append(texto_bloco)

    por_data = []
    gerado_em = ""
    i = 0
    while i < len(linhas_texto):
        linha = linhas_texto[i].strip()
        if re.match(r'^\d{2}/\d{2}/\d{4},\s*\d{2}:\d{2}', linha) and not gerado_em:
            gerado_em = linha
        if re.match(r'^\d{2}/\d{2}/\d{4}$', linha):
            data_br = linha
            try:
                data_iso = datetime.strptime(data_br, "%d/%m/%Y").strftime("%Y-%m-%d")
            except Exception:
                data_iso = data_br
            qtd = 0
            valor = 0.0
            if i + 1 < len(linhas_texto) and linhas_texto[i+1].strip().isdigit():
                qtd = int(linhas_texto[i+1].strip())
            if i + 2 < len(linhas_texto) and ('R$' in linhas_texto[i+2] or ',' in linhas_texto[i+2]):
                val_txt = re.sub(r'[^\d,]', '', linhas_texto[i+2]).replace(',', '.')
                try:
                    valor = float(val_txt)
                except Exception:
                    valor = 0.0
            por_data.append({
                "data_iso": data_iso,
                "data_br": data_br,
                "quantidade": qtd,
                "valor_pdf": round(valor, 2)
            })
            i += 3
            continue
        i += 1

    comprovantes_imgs = []
    for idx_img, img_bytes in enumerate(imagens_jpeg, start=1):
        comprovantes_imgs.append({
            "numero": idx_img,
            "titulo": f"Comprovante {idx_img}",
            "bytes": img_bytes,
            "b64": base64.b64encode(img_bytes).decode("utf-8")
        })

    return {
        "gerado_em": gerado_em,
        "por_data": por_data,
        "imagens_comprovantes": comprovantes_imgs,
        "qtd_fotos": len(comprovantes_imgs),
        "total_pdf": round(sum(x["valor_pdf"] for x in por_data), 2),
        "total_comprovantes": sum(x["quantidade"] for x in por_data),
        "linhas_brutas": linhas_texto
    }


def desempacotar_arquivos_pedagio(arquivos_uploaded) -> dict:
    """
    Recebe uma lista de arquivos enviados no Streamlit (.msg, .eml, .xlsx, .xlsm, .pdf)
    e separa automaticamente:
    - wb_fechamento (planilha de fechamento da base)
    - nome_fechamento
    - wb_totvs (relatório Totvs ESOS, se enviado)
    - nome_totvs
    - pdf_pedagios_bytes e nome_pdf (Relatório de Pedágios.pdf)
    - base_sugerida (sigla de 3 letras, ex: BRU, CAN, PLN)
    """
    planilhas_candidatas = []
    pdfs_candidatos = []
    textos_contexto = []

    for arq in arquivos_uploaded:
        if hasattr(arq, "seek"):
            arq.seek(0)
        nome = arq.name
        conteudo = arq.getvalue() if hasattr(arq, "getvalue") else arq.read()
        nome_up = nome.upper()
        textos_contexto.append(nome)

        if nome_up.endswith((".XLSX", ".XLSM")):
            planilhas_candidatas.append((nome, conteudo))
        elif nome_up.endswith(".PDF"):
            pdfs_candidatos.append((nome, conteudo))
        elif nome_up.endswith(".MSG") and extract_msg is not None:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
                tmp.write(conteudo)
                tmp_path = tmp.name
            try:
                with extract_msg.openMsg(tmp_path) as msg:
                    if msg.subject:
                        textos_contexto.append(str(msg.subject))
                    for att in msg.attachments:
                        att_nome = att.longFilename or att.shortFilename or ""
                        att_up = att_nome.upper()
                        if att_up.endswith((".XLSX", ".XLSM")) and att.data:
                            planilhas_candidatas.append((att_nome, att.data))
                        elif att_up.endswith(".PDF") and att.data:
                            pdfs_candidatos.append((att_nome, att.data))
            finally:
                if os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except Exception:
                        pass
        elif nome_up.endswith(".EML"):
            import email
            from email import policy
            msg_eml = email.message_from_bytes(conteudo, policy=policy.default)
            if msg_eml.get("subject"):
                textos_contexto.append(str(msg_eml.get("subject")))
            for part in msg_eml.walk():
                fn = part.get_filename()
                if not fn:
                    continue
                fn_up = fn.upper()
                payload = part.get_payload(decode=True)
                if not payload:
                    continue
                if fn_up.endswith((".XLSX", ".XLSM")):
                    planilhas_candidatas.append((fn, payload))
                elif fn_up.endswith(".PDF"):
                    pdfs_candidatos.append((fn, payload))

    wb_fechamento = None
    nome_fechamento = ""
    wb_totvs = None
    nome_totvs = ""

    for nome_pl, bytes_pl in planilhas_candidatas:
        try:
            wb = carregar_planilha_segura(bytes_pl, senha="FECHAMENTO")
            snames_up = [s.upper() for s in wb.sheetnames]
            if any(s.startswith("ESOS") for s in snames_up) or ("CONSOLIDADO" not in snames_up and len(wb.sheetnames) <= 2):
                if wb_totvs is None:
                    wb_totvs = wb
                    nome_totvs = nome_pl
            else:
                if wb_fechamento is None:
                    wb_fechamento = wb
                    nome_fechamento = nome_pl
                    # Tenta detectar base pela célula BT5 da aba CONSOLIDADO
                    if "CONSOLIDADO" in wb.sheetnames:
                        val_bt5 = str(wb["CONSOLIDADO"]["BT5"].value or "").strip()
                        if val_bt5:
                            textos_contexto.insert(0, val_bt5)
        except Exception:
            pass

    pdf_escolhido = None
    nome_pdf = ""
    for n_pdf, b_pdf in pdfs_candidatos:
        n_up = n_pdf.upper()
        if "PEDAGIO" in n_up or "PEDÁGIO" in n_up:
            pdf_escolhido = b_pdf
            nome_pdf = n_pdf
            break
    if pdf_escolhido is None and pdfs_candidatos:
        pdf_escolhido = pdfs_candidatos[0][1]
        nome_pdf = pdfs_candidatos[0][0]

    base_sugerida = None
    for t_ctx in textos_contexto:
        b = encontrar_base(t_ctx)
        if b:
            base_sugerida = b
            break
    if not base_sugerida and nome_fechamento:
        base_sugerida = detectar_base_por_texto(nome_fechamento)

    return {
        "wb_fechamento": wb_fechamento,
        "nome_fechamento": nome_fechamento,
        "wb_totvs": wb_totvs,
        "nome_totvs": nome_totvs,
        "pdf_bytes": pdf_escolhido,
        "nome_pdf": nome_pdf,
        "base_sugerida": base_sugerida or "BRU"
    }


def processar_auditoria_pedagios(wb_fechamento, pdf_bytes=None, wb_totvs=None, base_codigo="BRU", base_endereco_custom=None) -> dict:
    """
    Executa a auditoria completa de Pedágios:
    1. Extrai todas as OSs e pedágios cobrados na Planilha de Fechamento.
    2. Se houver Relatório Totvs, enriquece datas e endereços.
    3. Se houver Relatório de Pedágios (PDF), extrai a consolidação por data e fotos dos comprovantes.
    4. Calcula a rota do dia (cidades e maior KM estimado) para verificar se o deslocamento justifica pedágio.
    5. Executa o detector inteligente de anomalias:
       - Bateu Exato (Planilha == PDF)
       - Alerta de 0 KM cobrado na OS com cobrança de pedágio (possível estacionamento ou cobrança indevida)
       - Troca de Datas (valor do PDF lançado em outro dia na planilha)
       - Soma Acumulada (valor de 2 ou mais dias do PDF somados numa única OS / sem data na planilha)
       - Divergência de Valor / Sem comprovante no PDF
    """
    info_base = obter_info_base(base_codigo)
    uf_base = info_base["uf"]
    cidade_base = info_base["cidade_ibge"]
    base_endereco = base_endereco_custom or info_base["endereco"]

    _, _, abas_auditadas, todas_oss_fechamento = extrair_todas_oss_fechamento(wb_fechamento)
    set_todas_oss = {reg["os"] for reg in todas_oss_fechamento}

    os_dict = {}
    if wb_totvs is not None:
        try:
            os_dict, _ = carregar_dados_totvs(wb_totvs, uf_padrao=uf_base, oss_filtro=set_todas_oss)
        except Exception:
            os_dict = {}

    # Agrupa OSs do fechamento por dia (e separa OSs sem data que têm pedágio)
    fech_por_dia = {}
    oss_com_pedagio_lista = []

    for reg in todas_oss_fechamento:
        os_id = reg["os"]
        totvs_os = os_dict.get(os_id, {})
        dt_iso = reg["data_str"] or totvs_os.get("data_str")

        if dt_iso:
            try:
                data_br = datetime.strptime(dt_iso, "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                data_br = dt_iso
            chave_dia = dt_iso
        else:
            if reg["pedagio"] > 0 or reg["km_cobrado"] > 0:
                data_br = f"Sem Data (OS {os_id})"
                chave_dia = f"SEM_DATA_{os_id}"
            else:
                continue

        cidade_os = (reg["local"] or totvs_os.get("cidade") or "").strip()
        servico_os = (totvs_os.get("servico") or reg["servico_fechamento"] or "ATENDIMENTO").strip()
        end_limpo = totvs_os.get("endereco_limpo") or (f"{cidade_os.title()} - {uf_base}" if cidade_os else "")

        item_os = {
            "aba": reg["aba"],
            "linha": reg["linha"],
            "os": os_id,
            "data_iso": dt_iso or "",
            "data_br": data_br,
            "chave_dia": chave_dia,
            "cod_cliente": reg["cod_cliente"] or totvs_os.get("cod_cliente", ""),
            "cidade": cidade_os.title() if cidade_os else "-",
            "endereco": end_limpo,
            "km_cobrado": float(reg["km_cobrado"] or 0.0),
            "pedagio": round(float(reg["pedagio"] or 0.0), 2),
            "servico": servico_os
        }

        if item_os["pedagio"] > 0:
            oss_com_pedagio_lista.append(item_os)

        if chave_dia not in fech_por_dia:
            fech_por_dia[chave_dia] = {
                "chave_dia": chave_dia,
                "data_br": data_br,
                "tem_data_valida": bool(dt_iso),
                "todas_oss": [],
                "oss_pedagio": [],
                "soma_pedagio_planilha": 0.0,
                "soma_km_cobrado": 0.0,
                "cidades_dia": []
            }

        grp = fech_por_dia[chave_dia]
        grp["todas_oss"].append(item_os)
        grp["soma_km_cobrado"] = round(grp["soma_km_cobrado"] + item_os["km_cobrado"], 1)
        if item_os["pedagio"] > 0:
            grp["oss_pedagio"].append(item_os)
            grp["soma_pedagio_planilha"] = round(grp["soma_pedagio_planilha"] + item_os["pedagio"], 2)
        if item_os["cidade"] and item_os["cidade"] != "-" and item_os["cidade"] not in grp["cidades_dia"]:
            grp["cidades_dia"].append(item_os["cidade"])

    # Extrai dados do PDF se fornecido
    dados_pdf = None
    pdf_por_dia = {}
    if pdf_bytes:
        dados_pdf = extrair_dados_pdf_pedagios(pdf_bytes)
        for item_p in dados_pdf.get("por_data", []):
            d_iso = item_p["data_iso"]
            pdf_por_dia[d_iso] = {
                "data_iso": d_iso,
                "data_br": item_p["data_br"],
                "quantidade": item_p["quantidade"],
                "valor_pdf": item_p["valor_pdf"]
            }

    # Une todas as chaves de dias relevantes (dias com pedágio na planilha OU dias presentes no PDF)
    todas_chaves = sorted(
        set(k for k, v in fech_por_dia.items() if v["soma_pedagio_planilha"] > 0)
        | set(pdf_por_dia.keys())
    )

    # Inteligência de Conciliação:
    # 1. Identifica dias que já batem direto na mesma data
    dias_exatos = set()
    fech_pendentes = {}
    pdf_pendentes = {}

    for k in todas_chaves:
        val_f = fech_por_dia[k]["soma_pedagio_planilha"] if k in fech_por_dia else 0.0
        val_p = pdf_por_dia[k]["valor_pdf"] if k in pdf_por_dia else 0.0
        if val_f > 0 and val_p > 0 and abs(val_f - val_p) <= 0.05:
            dias_exatos.add(k)
        else:
            if val_f > 0 and val_p == 0:
                fech_pendentes[k] = val_f
            elif val_p > 0 and val_f == 0:
                pdf_pendentes[k] = val_p

    # 2. Detecta Troca de Data (1 para 1 entre fech_pendentes e pdf_pendentes)
    vinculos_data_trocada = {}  # chave_fech -> chave_pdf
    vinculos_pdf_trocado = {}   # chave_pdf -> chave_fech

    for k_f, v_f in list(fech_pendentes.items()):
        for k_p, v_p in list(pdf_pendentes.items()):
            if abs(v_f - v_p) <= 0.05:
                vinculos_data_trocada[k_f] = k_p
                vinculos_pdf_trocado[k_p] = k_f
                del fech_pendentes[k_f]
                del pdf_pendentes[k_p]
                break

    # 3. Detecta Soma Acumulada (1 lançamento na planilha = soma de 2 ou mais dias pendentes do PDF)
    vinculos_soma_acumulada = {}  # chave_fech -> [chaves_pdf]
    vinculos_pdf_acumulado = {}   # chave_pdf -> chave_fech

    for k_f, v_f in list(fech_pendentes.items()):
        chaves_p_disp = list(pdf_pendentes.keys())
        encontrou_comb = False
        for tam in range(2, min(len(chaves_p_disp) + 1, 7)):
            for comb in itertools.combinations(chaves_p_disp, tam):
                soma_comb = round(sum(pdf_pendentes[c] for c in comb), 2)
                if abs(soma_comb - v_f) <= 0.05:
                    vinculos_soma_acumulada[k_f] = list(comb)
                    for c in comb:
                        vinculos_pdf_acumulado[c] = k_f
                        del pdf_pendentes[c]
                    del fech_pendentes[k_f]
                    encontrou_comb = True
                    break
            if encontrou_comb:
                break

    # Constrói a lista final de análise dia a dia
    conciliacao_dias = []
    for k in todas_chaves:
        grp_f = fech_por_dia.get(k)
        grp_p = pdf_por_dia.get(k)

        data_exibicao = grp_f["data_br"] if grp_f else grp_p["data_br"]
        val_planilha = grp_f["soma_pedagio_planilha"] if grp_f else 0.0
        val_pdf = grp_p["valor_pdf"] if grp_p else 0.0
        qtd_recibos_pdf = grp_p["quantidade"] if grp_p else 0
        dif_valor = round(val_planilha - val_pdf, 2)

        oss_ped = grp_f["oss_pedagio"] if grp_f else []
        todas_oss_dia = grp_f["todas_oss"] if grp_f else []
        cidades_dia = grp_f["cidades_dia"] if grp_f else []
        km_cobrado_dia = grp_f["soma_km_cobrado"] if grp_f else 0.0
        km_cobrado_nas_oss_ped = sum(o["km_cobrado"] for o in oss_ped)

        # Calcula rota e link do Google Maps para as cidades atendidas no dia
        if cidades_dia:
            info_rota = calcular_rota_osrm_por_cidades(
                cidade_base=cidade_base,
                uf_base=uf_base,
                lista_cidades_dia=cidades_dia,
                qtd_enderecos_distintos=max(1, len(todas_oss_dia))
            )
            km_rota_mapa = info_rota["km_total_estimado"]
            seq_rota = info_rota["sequencia_formatada"]
            cidades_ord_maps = info_rota["cidades_ordenadas"]
            link_maps = gerar_url_google_maps(base_endereco, cidades_ord_maps, base_endereco)
        else:
            km_rota_mapa = 0.0
            seq_rota = "Sem deslocamento registrado na planilha para esta data"
            link_maps = ""

        # Determina diagnóstico e nível de alerta
        alertas_dia = []
        if not pdf_bytes:
            if not grp_f["tem_data_valida"]:
                status_codigo = "ALERTA"
                badge_status = "🚨 OS SEM DATA NA PLANILHA"
                explicacao = f"Foram cobrados {formatar_moeda_br(val_planilha)} em OS sem data preenchida ({', '.join('OS ' + str(o['os']) for o in oss_ped)})."
            elif km_cobrado_nas_oss_ped == 0:
                status_codigo = "ATENCAO"
                badge_status = "⚠️ PEDÁGIO EM OS COM 0 KM"
                explicacao = f"A OS que cobra {formatar_moeda_br(val_planilha)} está com 0 km de deslocamento na planilha. Anexe o PDF de pedágios para conferir o comprovante."
            else:
                status_codigo = "INFO"
                badge_status = "📋 REGISTRADO NA PLANILHA"
                explicacao = "Anexe o Relatório de Pedágios (.pdf) para cruzar automaticamente os valores e comprovantes."
        else:
            if k in dias_exatos:
                if km_cobrado_nas_oss_ped == 0:
                    status_codigo = "ATENCAO"
                    badge_status = "⚠️ VALOR BATE, MAS OS COM 0 KM"
                    explicacao = (
                        f"O valor ({formatar_moeda_br(val_planilha)}) bate com o PDF ({qtd_recibos_pdf} recibo(s)), "
                        f"porém a OS não cobrou KM (0 km). Confira as fotos dos comprovantes abaixo (ex: risco de recibo de estacionamento)."
                    )
                else:
                    status_codigo = "APROVADO"
                    badge_status = "✅ BATEU EXATO COM O PDF"
                    explicacao = f"Valor da planilha ({formatar_moeda_br(val_planilha)}) confere exatamente com o Relatório PDF na mesma data ({qtd_recibos_pdf} comprovante(s))."
            elif k in vinculos_data_trocada:
                k_pdf_orig = vinculos_data_trocada[k]
                dt_pdf_br = pdf_por_dia[k_pdf_orig]["data_br"]
                qtd_p_orig = pdf_por_dia[k_pdf_orig]["quantidade"]
                status_codigo = "ATENCAO"
                badge_status = f"🔄 DATA TROCADA (PDF: {dt_pdf_br})"
                explicacao = (
                    f"Na planilha foi lançado em {data_exibicao} ({formatar_moeda_br(val_planilha)}), mas no Relatório PDF "
                    f"esse valor exato ({qtd_p_orig} recibo(s)) pertence ao dia {dt_pdf_br}!"
                )
            elif k in vinculos_pdf_trocado:
                k_fech_dest = vinculos_pdf_trocado[k]
                dt_fech_br = fech_por_dia[k_fech_dest]["data_br"]
                oss_dest = ", ".join(f"OS {o['os']}" for o in fech_por_dia[k_fech_dest]["oss_pedagio"])
                status_codigo = "ATENCAO"
                badge_status = f"🔄 LANÇADO EM {dt_fech_br}"
                explicacao = f"Consta no PDF em {data_exibicao} ({formatar_moeda_br(val_pdf)}), mas foi cobrado na planilha no dia {dt_fech_br} ({oss_dest})."
            elif k in vinculos_soma_acumulada:
                lista_k_pdf = vinculos_soma_acumulada[k]
                partes_soma = [f"{pdf_por_dia[kp]['data_br']} ({formatar_moeda_br(pdf_por_dia[kp]['valor_pdf'])})" for kp in lista_k_pdf]
                status_codigo = "ALERTA"
                badge_status = "🧩 SOMA ACUMULADA DE OUTROS DIAS"
                explicacao = (
                    f"Cobrança acumulada de {formatar_moeda_br(val_planilha)} lançada de uma só vez! "
                    f"Equivale exatamente à soma de {len(lista_k_pdf)} dias do PDF: {' + '.join(partes_soma)}."
                )
                if not (grp_f and grp_f["tem_data_valida"]):
                    explicacao += " Além disso, a OS está SEM DATA preenchida na planilha!"
                if km_cobrado_nas_oss_ped == 0:
                    explicacao += " E a OS está com 0 KM cobrado!"
            elif k in vinculos_pdf_acumulado:
                k_fech_dest = vinculos_pdf_acumulado[k]
                dt_fech_br = fech_por_dia[k_fech_dest]["data_br"]
                oss_dest = ", ".join(f"OS {o['os']}" for o in fech_por_dia[k_fech_dest]["oss_pedagio"])
                val_tot_acum = fech_por_dia[k_fech_dest]["soma_pedagio_planilha"]
                status_codigo = "ATENCAO"
                badge_status = f"🧩 EMBUTIDO EM '{dt_fech_br}'"
                explicacao = (
                    f"O valor deste dia no PDF ({formatar_moeda_br(val_pdf)}, {qtd_recibos_pdf} recibo(s)) não foi lançado na data {data_exibicao}: "
                    f"ele foi somado e embutido em '{dt_fech_br}' ({oss_dest}, total {formatar_moeda_br(val_tot_acum)})."
                )
            elif val_planilha > 0 and val_pdf == 0:
                status_codigo = "ALERTA"
                badge_status = "🚨 SEM COMPROVANTE NO PDF"
                explicacao = f"Cobrado {formatar_moeda_br(val_planilha)} na planilha, mas não há nenhum valor correspondente para esta data no Relatório PDF."
            elif val_pdf > 0 and val_planilha == 0:
                status_codigo = "INFO"
                badge_status = "📄 CONSTA APENAS NO PDF"
                explicacao = f"Consta {formatar_moeda_br(val_pdf)} ({qtd_recibos_pdf} recibo(s)) no PDF nesta data, sem lançamento direto nesta data na planilha."
            else:
                status_codigo = "ALERTA"
                badge_status = f"🚨 DIVERGÊNCIA ({dif_valor:+,.2f})"
                explicacao = f"Planilha cobra {formatar_moeda_br(val_planilha)}, mas o PDF comprova {formatar_moeda_br(val_pdf)} (diferença de {formatar_moeda_br(dif_valor)})."

        conciliacao_dias.append({
            "chave_dia": k,
            "Data": data_exibicao,
            "tem_data_valida": grp_f["tem_data_valida"] if grp_f else True,
            "OSs com Pedágio": ", ".join(f"OS {o['os']} ({o['aba']})" for o in oss_ped) if oss_ped else "-",
            "Lista OSs Pedágio": oss_ped,
            "Todas as OSs do Dia": todas_oss_dia,
            "Cidades do Dia": ", ".join(cidades_dia) if cidades_dia else "-",
            "Sequência da Rota": seq_rota,
            "KM Cobrado no Dia": km_cobrado_dia,
            "KM Cobrado nas OSs c/ Pedágio": km_cobrado_nas_oss_ped,
            "KM Calculado da Rota": km_rota_mapa,
            "Link Maps": link_maps,
            "Pedágio na Planilha": round(val_planilha, 2),
            "Pedágio no PDF": round(val_pdf, 2),
            "Qtd Recibos PDF": qtd_recibos_pdf,
            "Diferença (Planilha - PDF)": dif_valor,
            "Status Código": status_codigo,
            "Badge Status": badge_status,
            "Explicação": explicacao
        })

    total_planilha = round(sum(o["pedagio"] for o in oss_com_pedagio_lista), 2)
    total_pdf = dados_pdf["total_pdf"] if dados_pdf else 0.0
    diferenca_total = round(total_planilha - total_pdf, 2) if dados_pdf else 0.0

    qtd_aprovados = sum(1 for d in conciliacao_dias if d["Status Código"] == "APROVADO")
    qtd_alertas = sum(1 for d in conciliacao_dias if d["Status Código"] in ["ALERTA", "ATENCAO"] and d["Pedágio na Planilha"] > 0)

    return {
        "base_codigo": base_codigo,
        "base_endereco": base_endereco,
        "total_planilha": total_planilha,
        "total_pdf": total_pdf,
        "diferenca_total": diferenca_total,
        "qtd_oss_com_pedagio": len(oss_com_pedagio_lista),
        "qtd_comprovantes_pdf": dados_pdf["total_comprovantes"] if dados_pdf else 0,
        "qtd_fotos_pdf": dados_pdf["qtd_fotos"] if dados_pdf else 0,
        "imagens_comprovantes": dados_pdf["imagens_comprovantes"] if dados_pdf else [],
        "por_data_pdf": dados_pdf["por_data"] if dados_pdf else [],
        "gerado_em_pdf": dados_pdf["gerado_em"] if dados_pdf else "",
        "qtd_dias_aprovados": qtd_aprovados,
        "qtd_dias_com_alerta": qtd_alertas,
        "conciliacao_dias": conciliacao_dias,
        "oss_com_pedagio": oss_com_pedagio_lista,
        "tem_pdf": bool(dados_pdf)
    }


def exportar_relatorio_pedagios_excel(dados_pedagio: dict, output_stream):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Auditoria de Pedágios"

    cor_azul = "1F4E79"
    cor_branco = "FFFFFF"
    cor_alerta = "FFF2CC"

    fonte_cab = Font(name="Calibri", size=11, bold=True, color=cor_branco)
    fill_cab = PatternFill(start_color=cor_azul, end_color=cor_azul, fill_type="solid")
    fill_alerta = PatternFill(start_color=cor_alerta, end_color=cor_alerta, fill_type="solid")
    borda = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    headers = [
        "Data", "OSs Cobrando Pedágio", "Cidades Visitadas",
        "KM Cobrado na Planilha", "Maior KM da Rota (Maps)",
        "Pedágio na Planilha (R$)", "Pedágio no PDF (R$)",
        "Qtd. Recibos PDF", "Diferença (R$)", "Status", "Diagnóstico / Observação"
    ]
    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        c = ws.cell(1, col_idx)
        c.font = fonte_cab
        c.fill = fill_cab
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for r_idx, d in enumerate(dados_pedagio.get("conciliacao_dias", []), start=2):
        ws.cell(r_idx, 1, d["Data"])
        ws.cell(r_idx, 2, d["OSs com Pedágio"])
        ws.cell(r_idx, 3, d["Cidades do Dia"])
        ws.cell(r_idx, 4, d["KM Cobrado no Dia"])
        ws.cell(r_idx, 5, d["KM Calculado da Rota"])
        ws.cell(r_idx, 6, d["Pedágio na Planilha"])
        ws.cell(r_idx, 7, d["Pedágio no PDF"])
        ws.cell(r_idx, 8, d["Qtd Recibos PDF"])
        ws.cell(r_idx, 9, d["Diferença (Planilha - PDF)"])
        ws.cell(r_idx, 10, d["Badge Status"])
        ws.cell(r_idx, 11, d["Explicação"])

        is_alerta = d["Status Código"] in ["ALERTA", "ATENCAO"]
        for c_i in range(1, len(headers) + 1):
            cell = ws.cell(r_idx, c_i)
            cell.border = borda
            if is_alerta:
                cell.fill = fill_alerta
            if c_i in [6, 7, 9]:
                cell.number_format = 'R$ #,##0.00'

    larguras = [16, 26, 24, 16, 16, 18, 18, 14, 16, 28, 65]
    for i, w in enumerate(larguras, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    wb.save(output_stream)

