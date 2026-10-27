import io
import os
import re
import json
import tempfile
import unicodedata
from datetime import datetime
import openpyxl
import pandas as pd

from auditar_engine import carregar_planilha_segura

try:
    import extract_msg
except ImportError:
    extract_msg = None


# Tabela Oficial de Bases Consigaz (Código -> Abreviação Oficial de 3 letras)
# 01 - Barueri: BRU
# 02 - Paulínia: PLN
# 03 - São Vicente: SVC
# 04 - São José dos Campos: SJC
# 05 - Cariacica: CAR
# 06 - Canoas: CAN
# 07 - Senador Canedo: SNC
# 08 - Mauá: MAU
# 09 - Duque de Caxias: DQX
# 10 - Araucária: ARA
# 11 - Ribeirão Preto: RIB
# 13 - Uberlândia: UBR
# 14 - Brasília: BRA
# 15 - Betim: BTM
# 17 - Porto Belo: PTB
# 18 - Jandaia do Sul: JDS

BASES = [
    "BRU", "PLN", "SVC", "SJC", "CAR", "CAN", "SNC", "MAU",
    "DQX", "ARA", "RIB", "UBR", "BRA", "BTM", "PTB", "JDS", "CORP"
]

SIGLAS_BASES = {
    "BRU": "BRU",
    "PLN": "PLN",
    "SVC": "SVC",
    "SJC": "SJC",
    "CAR": "CAR",
    "CAN": "CAN",
    "SNC": "SNC",
    "MAU": "MAU",
    "DQX": "DQX",
    "ARA": "ARA",
    "RIB": "RIB",
    "RBP": "RIB",
    "UBR": "UBR",
    "BRA": "BRA",
    "BSB": "BRA",
    "BTM": "BTM",
    "PTB": "PTB",
    "JDS": "JDS",
    "CORP": "CORP",
}

NOMES_CIDADES_BASES = [
    ("SAO JOSE DOS CAMPOS", "SJC"),
    ("S J DOS CAMPOS", "SJC"),
    ("SJCAMPOS", "SJC"),
    ("DUQUE DE CAXIAS", "DQX"),
    ("SENADOR CANEDO", "SNC"),
    ("RIBEIRAO PRETO", "RIB"),
    ("JANDAIA DO SUL", "JDS"),
    ("SAO VICENTE", "SVC"),
    ("PORTO BELO", "PTB"),
    ("UBERLANDIA", "UBR"),
    ("ARAUCARIA", "ARA"),
    ("ARAUCARA", "ARA"),
    ("CARIACICA", "CAR"),
    ("PAULINIA", "PLN"),
    ("PAULINA", "PLN"),
    ("BRASILIA", "BRA"),
    ("BARUERI", "BRU"),
    ("JANDAIA", "JDS"),
    ("CANEDO", "SNC"),
    ("CAXIAS", "DQX"),
    ("CANOAS", "CAN"),
    ("BETIM", "BTM"),
    ("MAUA", "MAU"),
    ("CORPORATIVO", "CORP"),
]

NUMEROS_BASES = {
    1: "BRU",
    2: "PLN",
    3: "SVC",
    4: "SJC",
    5: "CAR",
    6: "CAN",
    7: "SNC",
    8: "MAU",
    9: "DQX",
    10: "ARA",
    11: "RIB",
    13: "UBR",
    14: "BRA",
    15: "BTM",
    17: "PTB",
    18: "JDS",
}

MAPA_BASES = {**SIGLAS_BASES, **dict(NOMES_CIDADES_BASES)}

MESES = [
    "JAN", "FEV", "MAR", "ABR", "MAI", "JUN",
    "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"
]

MAPA_MESES = {
    "JAN": "JAN", "JANEIRO": "JAN",
    "FEV": "FEV", "FEVEREIRO": "FEV",
    "MAR": "MAR", "MARCO": "MAR", "MARÇO": "MAR",
    "ABR": "ABR", "ABRIL": "ABR",
    "MAI": "MAI", "MAIO": "MAI",
    "JUN": "JUN", "JUNHO": "JUN",
    "JUL": "JUL", "JULHO": "JUL",
    "AGO": "AGO", "AGOSTO": "AGO",
    "SET": "SET", "SETEMBRO": "SET",
    "OUT": "OUT", "OUTUBRO": "OUT",
    "NOV": "NOV", "NOVEMBRO": "NOV",
    "DEZ": "DEZ", "DEZEMBRO": "DEZ"
}


def _normalizar_sem_acento(texto: str) -> str:
    if not texto:
        return ""
    return "".join(
        c for c in unicodedata.normalize("NFD", str(texto).upper())
        if unicodedata.category(c) != "Mn"
    )


def limpar_assunto(texto: str) -> str:
    """Remove prefixos como ENC:, RE:, RES:, FWD: e espaços extras."""
    if not texto:
        return ""
    padrao = r'^(?:(?:ENC|RE|RES|FWD|FW|AN[AÁ]LISE)\s*[:_]\s*)+'
    limpo = re.sub(padrao, '', str(texto).strip(), flags=re.IGNORECASE).strip()
    return limpo


def extrair_partes(texto: str):
    """Divide o texto em palavras isoladas na ordem em que aparecem."""
    if not texto:
        return []
    return [p for p in re.split(r'[-_/\\\s\[\]().,;:!ºª]+', str(texto).upper()) if p]


def encontrar_sigla_exata(texto: str):
    """Procura especificamente pela abreviação de 3 letras (ex: BRU, PLN, SVC, SJC...) como palavra isolada."""
    if not texto:
        return None
    txt_norm = _normalizar_sem_acento(texto)
    for palavra in re.findall(r'[A-Z]+', txt_norm):
        if palavra in SIGLAS_BASES:
            return SIGLAS_BASES[palavra]
    return None


def encontrar_base(texto: str, aceitar_numero_puro: bool = False):
    """
    Identifica a Base sempre retornando a abreviação oficial de 3 letras (ex: BRU, PLN, SVC, SJC...).
    1. Procura primeiro pela abreviação exata (token isolado, sem confundir com pedaços de palavras como OUT[UBR]O).
    2. Procura pelo nome da cidade (ex: Barueri -> BRU, São José dos Campos -> SJC).
    3. Procura por 'Base XX' ou código numérico da base (quando na célula da planilha).
    """
    if not texto:
        return None

    sigla = encontrar_sigla_exata(texto)
    if sigla:
        return sigla

    txt_norm = _normalizar_sem_acento(texto)
    palavras = re.findall(r'[A-Z]+', txt_norm)
    frase = " " + " ".join(palavras) + " "
    for nome_cid, s_base in NOMES_CIDADES_BASES:
        if f" {nome_cid} " in frase:
            return s_base

    m_base = re.search(r'\bBASE\s*0*(\d{1,2})\b', txt_norm)
    if m_base:
        num = int(m_base.group(1))
        if num in NUMEROS_BASES:
            return NUMEROS_BASES[num]

    if aceitar_numero_puro:
        m_num = re.match(r'^\s*0*(\d{1,2})(?:\b|\D|$)', txt_norm)
        if m_num:
            num = int(m_num.group(1))
            if num in NUMEROS_BASES:
                return NUMEROS_BASES[num]

    return None


def encontrar_mes(texto: str):
    if not texto:
        return None
    partes = extrair_partes(texto)
    for p in partes:
        if p in MAPA_MESES:
            return MAPA_MESES[p]
    return None


def normalizar_valor(v):
    if v is None:
        return None
    try:
        if isinstance(v, (int, float)):
            return round(float(v), 2)
        txt = str(v).strip().replace("R$", "").replace(" ", "")
        if "." in txt and "," in txt:
            txt = txt.replace(".", "").replace(",", ".")
        elif "," in txt:
            txt = txt.replace(",", ".")
        return round(float(txt), 2)
    except Exception:
        return None


def formatar_moeda_br(v):
    if v is None:
        return "R$ 0,00"
    try:
        val = float(v)
        texto = f"{val:,.2f}"
        return "R$ " + texto.replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)


def extrair_dados_de_planilha(conteudo_ou_stream, nome_arquivo="", texto_email=""):
    """
    Abre o workbook Excel (suportando senha com carregar_planilha_segura)
    e lê as células da aba CONSOLIDADO:
      - BT5: Base
      - I5: Fornecedor (Razão Social / Nome)
      - AQ5: Código / CNPJ do Fornecedor
      - BJ16: Valor Total
      - BD5: Período
    Sempre retorna a Base como abreviação de 3 letras (ex: BRU, PLN, SVC, SJC...),
    olhando por abreviação no e-mail ou no Excel (nome do arquivo e aba CONSOLIDADO).
    """
    wb = carregar_planilha_segura(conteudo_ou_stream, senha="FECHAMENTO")
    if not wb:
        raise ValueError("Não foi possível carregar a planilha (formato inválido ou corrompido).")

    if "CONSOLIDADO" not in wb.sheetnames:
        wb.close()
        raise ValueError("A aba 'CONSOLIDADO' não foi encontrada nesta planilha.")

    ws = wb["CONSOLIDADO"]

    # 1. Base: prioriza abreviação exata no e-mail, no nome do Excel ou na célula BT5;
    #    caso contrário, converte nome da cidade / código da base (01..18) para a abreviação de 3 letras.
    base_celula = str(ws["BT5"].value or "").strip()
    base_detectada = (
        encontrar_sigla_exata(texto_email)
        or encontrar_sigla_exata(nome_arquivo)
        or encontrar_sigla_exata(base_celula)
        or encontrar_base(texto_email, aceitar_numero_puro=False)
        or encontrar_base(nome_arquivo, aceitar_numero_puro=False)
        or encontrar_base(base_celula, aceitar_numero_puro=True)
    )

    # Se BT5 estiver mesclada ou deslocada na aba CONSOLIDADO, verifica células vizinhas do cabeçalho (linhas 3 a 7, colunas BP a BZ)
    if not base_detectada:
        for r_idx in range(3, 8):
            for c_idx in range(68, 79):
                val_viz = ws.cell(r_idx, c_idx).value
                if val_viz is not None:
                    b_viz = encontrar_base(str(val_viz), aceitar_numero_puro=False)
                    if b_viz:
                        base_detectada = b_viz
                        break
            if base_detectada:
                break

    # 2. Código do Fornecedor (AQ5)
    cod_celula = ws["AQ5"].value
    codigo_fornecedor = str(cod_celula).strip() if cod_celula is not None else ""

    # 3. Fornecedor (I5 ou fallback pelo nome do arquivo)
    fornec_celula = ws["I5"].value
    fornecedor = str(fornec_celula).strip() if fornec_celula is not None else ""

    # 4. Valor (BJ16)
    valor_celula = ws["BJ16"].value
    valor_num = normalizar_valor(valor_celula)

    # 5. Período (BD5)
    periodo_celula = str(ws["BD5"].value or "").strip()
    mes_detectado = encontrar_mes(periodo_celula) or encontrar_mes(nome_arquivo) or encontrar_mes(texto_email)

    wb.close()

    return {
        "base": base_detectada or "N/D",
        "codigo_fornecedor": codigo_fornecedor,
        "fornecedor": fornecedor or codigo_fornecedor or "N/D",
        "valor": valor_num,
        "periodo": periodo_celula or mes_detectado or "",
        "mes": mes_detectado or ""
    }


def _filtrar_registros_anexos_email(registros_email):
    """
    Quando um mesmo e-mail possui mais de um anexo Excel (ex: planilha de fechamento + relatório Totvs sem aba CONSOLIDADO),
    se pelo menos uma planilha válida (com aba CONSOLIDADO) foi lida com sucesso, ignora os anexos auxiliares sem CONSOLIDADO.
    """
    if len(registros_email) <= 1:
        return registros_email
    tem_valido = any(not str(r.get("status", "")).startswith("ERRO") for r in registros_email)
    if tem_valido:
        filtrados = [
            r for r in registros_email
            if "CONSOLIDADO" not in str(r.get("detalhe_erro", ""))
        ]
        if filtrados:
            return filtrados
    return registros_email


def processar_arquivo_uploaded(uploaded_file):
    """
    Processa um arquivo carregado via Streamlit (podendo ser .xlsx/.xlsm, .msg ou .eml).
    Suporta tanto o objeto UploadedFile do Streamlit quanto um dicionário com {'nome': ..., 'bytes': ...}.
    Retorna uma lista de dicionários com os registros extraídos.
    """
    if isinstance(uploaded_file, dict):
        nome_orig = uploaded_file.get("nome") or uploaded_file.get("name", "arquivo.xlsx")
        conteudo_bytes = uploaded_file.get("bytes", b"")
    else:
        if hasattr(uploaded_file, "seek"):
            uploaded_file.seek(0)
        nome_orig = uploaded_file.name
        if hasattr(uploaded_file, "getvalue"):
            conteudo_bytes = uploaded_file.getvalue()
        elif hasattr(uploaded_file, "getbuffer"):
            conteudo_bytes = bytes(uploaded_file.getbuffer())
        else:
            conteudo_bytes = uploaded_file.read()

    nome_upper = nome_orig.upper()
    registros = []

    if nome_upper.endswith(".MSG"):
        if extract_msg is None:
            raise ImportError("Biblioteca extract_msg não está instalada no ambiente.")

        # Salva em arquivo temporário para leitura segura pelo extract_msg
        with tempfile.NamedTemporaryFile(delete=False, suffix=".msg") as tmp:
            tmp.write(conteudo_bytes)
            tmp_path = tmp.name

        try:
            with extract_msg.openMsg(tmp_path) as msg:
                assunto_original = str(msg.subject or "")
                assunto_limpo = limpar_assunto(assunto_original)
                remetente = str(msg.sender or "")
                data_msg = msg.date
                texto_email = f"{assunto_original} {nome_orig}"

                data_str = ""
                if data_msg:
                    try:
                        if isinstance(data_msg, datetime):
                            data_str = data_msg.strftime("%d/%m/%Y")
                        else:
                            data_str = str(data_msg)[:10]
                    except Exception:
                        data_str = str(data_msg)
                if not data_str:
                    data_str = datetime.now().strftime("%d/%m/%Y")

                # Localiza anexos Excel
                anexos_excel = []
                for att in msg.attachments:
                    att_nome = att.longFilename or att.shortFilename or ""
                    if att_nome.upper().endswith((".XLSX", ".XLSM")):
                        anexos_excel.append((att_nome, att.data))

                if not anexos_excel:
                    base_assunto = encontrar_base(texto_email) or "N/D"
                    registros.append({
                        "base": base_assunto,
                        "identificacao": assunto_limpo or nome_orig,
                        "codigo_fornecedor": "",
                        "fornecedor": remetente or "Nenhum Fornecedor",
                        "valor": 0.0,
                        "status": "ERRO: Sem anexo Excel",
                        "data": data_str,
                        "arquivo": nome_orig,
                        "origem": "E-mail (.msg)",
                        "detalhe_erro": "E-mail recebido sem planilha Excel anexa.",
                        "arquivo_chave": nome_orig
                    })
                    return registros

                # Processa cada anexo Excel encontrado no e-mail
                registros_email = []
                for att_nome, att_bytes in anexos_excel:
                    try:
                        dados = extrair_dados_de_planilha(att_bytes, att_nome, texto_email=texto_email)
                        base_final = dados["base"]
                        if base_final == "N/D" or not base_final:
                            base_final = encontrar_base(texto_email) or encontrar_base(att_nome) or "N/D"

                        identificacao = assunto_limpo if assunto_limpo else os.path.splitext(att_nome)[0]

                        registros_email.append({
                            "base": base_final,
                            "identificacao": identificacao,
                            "codigo_fornecedor": dados["codigo_fornecedor"],
                            "fornecedor": dados["fornecedor"] or remetente,
                            "valor": dados["valor"],
                            "status": "PENDENTE",
                            "data": data_str,
                            "arquivo": att_nome,
                            "origem": f"E-mail: {nome_orig}",
                            "detalhe_erro": "",
                            "arquivo_chave": nome_orig
                        })
                    except Exception as err_att:
                        registros_email.append({
                            "base": encontrar_base(texto_email) or encontrar_base(att_nome) or "N/D",
                            "identificacao": assunto_limpo or att_nome,
                            "codigo_fornecedor": "",
                            "fornecedor": remetente or "N/D",
                            "valor": 0.0,
                            "status": f"ERRO: {str(err_att)[:40]}",
                            "data": data_str,
                            "arquivo": att_nome,
                            "origem": f"E-mail: {nome_orig}",
                            "detalhe_erro": str(err_att),
                            "arquivo_chave": nome_orig
                        })
                registros.extend(_filtrar_registros_anexos_email(registros_email))
        finally:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

    elif nome_upper.endswith(".EML"):
        import email
        from email import policy
        import email.utils

        try:
            msg = email.message_from_bytes(conteudo_bytes, policy=policy.default)
            assunto_original = str(msg.get("subject") or "")
            assunto_limpo = limpar_assunto(assunto_original)
            remetente = str(msg.get("from") or "")
            data_msg_header = msg.get("date")
            texto_email = f"{assunto_original} {nome_orig}"

            data_str = ""
            if data_msg_header:
                try:
                    dt = email.utils.parsedate_to_datetime(data_msg_header)
                    if dt:
                        data_str = dt.strftime("%d/%m/%Y")
                except Exception:
                    pass
            if not data_str:
                data_str = datetime.now().strftime("%d/%m/%Y")

            anexos_excel = []
            for part in msg.iter_attachments():
                att_nome = part.get_filename() or ""
                if att_nome.upper().endswith((".XLSX", ".XLSM")):
                    anexos_excel.append((att_nome, part.get_content()))

            if not anexos_excel:
                base_assunto = encontrar_base(texto_email) or "N/D"
                registros.append({
                    "base": base_assunto,
                    "identificacao": assunto_limpo or nome_orig,
                    "codigo_fornecedor": "",
                    "fornecedor": remetente or "Nenhum Fornecedor",
                    "valor": 0.0,
                    "status": "ERRO: Sem anexo Excel",
                    "data": data_str,
                    "arquivo": nome_orig,
                    "origem": "E-mail (.eml)",
                    "detalhe_erro": "E-mail recebido sem planilha Excel anexa.",
                    "arquivo_chave": nome_orig
                })
                return registros

            registros_email = []
            for att_nome, att_bytes in anexos_excel:
                try:
                    dados = extrair_dados_de_planilha(att_bytes, att_nome, texto_email=texto_email)
                    base_final = dados["base"]
                    if base_final == "N/D" or not base_final:
                        base_final = encontrar_base(texto_email) or encontrar_base(att_nome) or "N/D"

                    identificacao = assunto_limpo if assunto_limpo else os.path.splitext(att_nome)[0]

                    registros_email.append({
                        "base": base_final,
                        "identificacao": identificacao,
                        "codigo_fornecedor": dados["codigo_fornecedor"],
                        "fornecedor": dados["fornecedor"] or remetente,
                        "valor": dados["valor"],
                        "status": "PENDENTE",
                        "data": data_str,
                        "arquivo": att_nome,
                        "origem": f"E-mail: {nome_orig}",
                        "detalhe_erro": "",
                        "arquivo_chave": nome_orig
                    })
                except Exception as err_att:
                    registros_email.append({
                        "base": encontrar_base(texto_email) or encontrar_base(att_nome) or "N/D",
                        "identificacao": assunto_limpo or att_nome,
                        "codigo_fornecedor": "",
                        "fornecedor": remetente or "N/D",
                        "valor": 0.0,
                        "status": f"ERRO: {str(err_att)[:40]}",
                        "data": data_str,
                        "arquivo": att_nome,
                        "origem": f"E-mail: {nome_orig}",
                        "detalhe_erro": str(err_att),
                        "arquivo_chave": nome_orig
                    })
            registros.extend(_filtrar_registros_anexos_email(registros_email))
        except Exception as e:
            registros.append({
                "base": encontrar_base(nome_orig) or "N/D",
                "identificacao": nome_orig,
                "codigo_fornecedor": "",
                "fornecedor": "ERRO",
                "valor": 0.0,
                "status": f"ERRO: {str(e)[:40]}",
                "data": datetime.now().strftime("%d/%m/%Y"),
                "arquivo": nome_orig,
                "origem": "E-mail (.eml)",
                "detalhe_erro": str(e),
                "arquivo_chave": nome_orig
            })

    elif nome_upper.endswith((".XLSX", ".XLSM")):
        try:
            dados = extrair_dados_de_planilha(conteudo_bytes, nome_orig)
            nome_sem_ext = os.path.splitext(nome_orig)[0]
            data_hoje = datetime.now().strftime("%d/%m/%Y")

            registros.append({
                "base": dados["base"],
                "identificacao": nome_sem_ext,
                "codigo_fornecedor": dados["codigo_fornecedor"],
                "fornecedor": dados["fornecedor"],
                "valor": dados["valor"],
                "status": "PENDENTE",
                "data": data_hoje,
                "arquivo": nome_orig,
                "origem": "Upload Direto (.xlsx)",
                "detalhe_erro": "",
                "arquivo_chave": nome_orig
            })
        except Exception as err_plan:
            registros.append({
                "base": encontrar_base(nome_orig) or "N/D",
                "identificacao": os.path.splitext(nome_orig)[0],
                "codigo_fornecedor": "",
                "fornecedor": "N/D",
                "valor": 0.0,
                "status": f"ERRO: {str(err_plan)[:40]}",
                "data": datetime.now().strftime("%d/%m/%Y"),
                "arquivo": nome_orig,
                "origem": "Upload Direto (.xlsx)",
                "detalhe_erro": str(err_plan),
                "arquivo_chave": nome_orig
            })

    return registros


def extrair_chaves_historico(arquivo_controle):
    """
    Lê uma planilha de controle existente (se o usuário fizer upload para conferência)
    e extrai chaves (Base, Código/Fornecedor, Valor) para checar duplicidade prévia.
    """
    if not arquivo_controle:
        return set()
    try:
        wb = carregar_planilha_segura(arquivo_controle)
        ws = wb.active
        chaves = set()
        for linha in range(2, ws.max_row + 1):
            b = ws.cell(linha, 1).value
            c = ws.cell(linha, 3).value  # Código ou Fornecedor
            f = ws.cell(linha, 4).value  # Fornecedor
            v = ws.cell(linha, 5).value  # Valor
            b_str = str(b).strip().upper() if b else ""
            v_norm = normalizar_valor(v)
            if b_str and v_norm is not None:
                if c:
                    chaves.add((b_str, str(c).strip().upper(), v_norm))
                if f:
                    chaves.add((b_str, str(f).strip().upper(), v_norm))
        wb.close()
        return chaves
    except Exception:
        return set()


def processar_lote(arquivos_upload, chaves_historico=None):
    """
    Processa um conjunto de arquivos carregados e faz análise de duplicidades
    contra histórico e entre os próprios itens do lote.
    """
    if chaves_historico is None:
        chaves_historico = set()

    todos_registros = []
    chaves_vistas = set()

    for arq in arquivos_upload:
        try:
            itens = processar_arquivo_uploaded(arq)
            todos_registros.extend(itens)
        except Exception as e:
            nome_arq = arq.get("nome") if isinstance(arq, dict) else getattr(arq, "name", "Arquivo")
            todos_registros.append({
                "base": "N/D",
                "identificacao": nome_arq,
                "codigo_fornecedor": "",
                "fornecedor": "ERRO",
                "valor": 0.0,
                "status": f"ERRO: {str(e)[:40]}",
                "data": datetime.now().strftime("%d/%m/%Y"),
                "arquivo": nome_arq,
                "origem": "Falha de Leitura",
                "detalhe_erro": str(e),
                "arquivo_chave": nome_arq
            })

    # Verificação de duplicados
    for item in todos_registros:
        if str(item.get("status", "")).startswith("ERRO"):
            continue

        b = str(item.get("base") or "").strip().upper()
        f = str(item.get("fornecedor") or "").strip().upper()
        c = str(item.get("codigo_fornecedor") or "").strip().upper()
        v = normalizar_valor(item.get("valor"))

        if v is None or v <= 0:
            continue

        # 1. Verifica contra histórico existente
        if (b, c, v) in chaves_historico or (b, f, v) in chaves_historico:
            item["status"] = "DUPLICADO (Histórico)"
            item["detalhe_erro"] = f"Fechamento já consta no arquivo de controle (Base: {b}, Valor: {v})"
            continue

        # 2. Verifica dentro do próprio lote
        chave = (b, c or f, v)
        if chave in chaves_vistas:
            item["status"] = "DUPLICADO (No Lote)"
            item["detalhe_erro"] = f"Fechamento repetido dentro dos arquivos enviados (Base: {b}, Valor: {v})"
        else:
            chaves_vistas.add(chave)

    return todos_registros


def gerar_tsv_colar_excel(registros, incluir_cabecalho=False):
    """
    Gera string tabulada (TSV) pronta para copiar e colar diretamente no Excel:
    Colunas: Base, Identificação, Código / CNPJ, Fornecedor, Valor (R$), Status, Data
    """
    linhas = []
    if incluir_cabecalho:
        linhas.append("\t".join([
            "Base",
            "Identificação",
            "Código / CNPJ",
            "Fornecedor",
            "Valor",
            "Status",
            "Data Recebimento"
        ]))

    for r in registros:
        val = r.get("valor")
        val_str = f"{val:.2f}".replace(".", ",") if (isinstance(val, (int, float)) and val is not None) else str(val or "")
        linha = "\t".join([
            str(r.get("base", "")),
            str(r.get("identificacao", "")),
            str(r.get("codigo_fornecedor", "")),
            str(r.get("fornecedor", "")),
            val_str,
            str(r.get("status", "PENDENTE")),
            str(r.get("data", ""))
        ])
        linhas.append(linha)

    return "\n".join(linhas)


def gerar_excel_fechamento(registros):
    """Gera um arquivo Excel formatado com os dados dos fechamentos processados."""
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fechamento"

    headers = [
        "Base", "Identificação / Assunto", "Cód. Fornecedor",
        "Fornecedor", "Valor (R$)", "Status", "Data Recebimento",
        "Arquivo Original", "Origem"
    ]
    ws.append(headers)

    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    fill_header = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")
    font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    thin = Side(border_style="thin", color="CCCCCC")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)

    for col_idx, col_name in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = align_center

    for r_idx, r in enumerate(registros, 2):
        val = r.get("valor", 0.0)
        try:
            val_num = float(val) if val is not None else 0.0
        except Exception:
            val_num = 0.0

        row_data = [
            r.get("base", ""),
            r.get("identificacao", ""),
            r.get("codigo_fornecedor", ""),
            r.get("fornecedor", ""),
            val_num,
            r.get("status", "PENDENTE"),
            r.get("data", ""),
            r.get("arquivo", ""),
            r.get("origem", "")
        ]
        ws.append(row_data)

        ws.cell(row=r_idx, column=1).alignment = align_center
        ws.cell(row=r_idx, column=2).alignment = align_left
        ws.cell(row=r_idx, column=3).alignment = align_center
        ws.cell(row=r_idx, column=4).alignment = align_left
        
        cell_val = ws.cell(row=r_idx, column=5)
        cell_val.number_format = 'R$ #,##0.00'
        cell_val.alignment = align_right

        status_cell = ws.cell(row=r_idx, column=6)
        status_cell.alignment = align_center
        st_val = str(r.get("status", "")).upper()
        if "DUPLICADO" in st_val:
            status_cell.fill = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
        elif "ERRO" in st_val:
            status_cell.fill = PatternFill(start_color="FECACA", end_color="FECACA", fill_type="solid")
        else:
            status_cell.fill = PatternFill(start_color="BBF7D0", end_color="BBF7D0", fill_type="solid")

        ws.cell(row=r_idx, column=7).alignment = align_center

        for col_idx in range(1, 10):
            ws.cell(row=r_idx, column=col_idx).border = border

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    wb.save(output)
    output.seek(0)
    return output.getvalue()
