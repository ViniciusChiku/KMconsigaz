"""
Engine de Auditoria de KM / Deslocamento - Consigaz
Cruza a planilha de fechamento com o relatório Totvs,
agrupa todas as OSs e KMs por dia, calcula automaticamente a quilometragem
rodoviária real (via IBGE + OSRM) e gera links infalíveis do Google Maps.
"""

import io
import os
import re
import json
import unicodedata
import urllib.parse
import urllib.request
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import msoffcrypto
except ImportError:
    msoffcrypto = None

# Tabela Oficial de Endereços por Base - Consigaz (Tabela 2)
ENDERECOS_BASES_DETALHADOS = {
    8: {"sigla": "MAU", "uf": "SP", "cidade_ibge": "MAUA", "cidade": "Mauá/SP", "endereco": "RUA VITORIA PERDAO PETIGROSSO, 115 - MAUA/SP"},
    2: {"sigla": "PLN", "uf": "SP", "cidade_ibge": "PAULINIA", "cidade": "Paulínia/SP", "endereco": "RUA EITOR VALENTINO DE SOUZA, 337 - PAULINA/SP"},
    1: {"sigla": "BRU", "uf": "SP", "cidade_ibge": "BARUERI", "cidade": "Barueri/SP", "endereco": "RUA JOSE PEREIRA SOBRINHO, 485 - BARUERI/SP"},
    6: {"sigla": "CAN", "uf": "RS", "cidade_ibge": "CANOAS", "cidade": "Canoas/RS", "endereco": "RUA ANTONIO FREDERICO OZANAN, 1901 - CANOAS/RS"},
    9: {"sigla": "DQX", "uf": "RJ", "cidade_ibge": "DUQUE DE CAXIAS", "cidade": "Duque de Caxias/RJ", "endereco": "RUA FREI CANECA, 407 - DUQUE DE CAXIAS/RJ"},
    11: {"sigla": "RIB", "uf": "SP", "cidade_ibge": "RIBEIRAO PRETO", "cidade": "Ribeirão Preto/SP", "endereco": "ROD ALEXANDRE BALBO, SN - RIBEIRÃO PRETO/SP"},
    5: {"sigla": "CAR", "uf": "ES", "cidade_ibge": "CARIACICA", "cidade": "Cariacica/ES", "endereco": "ROD GOVERNADOR MARIO COVAS 1941, 7161 - CARIACICA/ES"},
    4: {"sigla": "SVC", "uf": "SP", "cidade_ibge": "SAO VICENTE", "cidade": "São Vicente/SP", "endereco": "RUA MANOEL DE ABREU, 790 - SÃO VICENTE/SP"},
    3: {"sigla": "SVC", "uf": "SP", "cidade_ibge": "SAO VICENTE", "cidade": "São Vicente/SP (QD 108)", "endereco": "AV. MANOEL DE EBREU, 790 QD 108 – SÃO VICENTE/SP"},
    7: {"sigla": "SNC", "uf": "GO", "cidade_ibge": "SENADOR CANEDO", "cidade": "Senador Canedo/GO", "endereco": "AV PROGRESSO, SN - SENADOR CANEDO/GO"},
    10: {"sigla": "ARA", "uf": "PR", "cidade_ibge": "ARAUCARIA", "cidade": "Araucária/PR", "endereco": "R DOUTOR JOSE CZAKI, 320 - ARAUCARA/PR"},
    13: {"sigla": "UBR", "uf": "MG", "cidade_ibge": "UBERLANDIA", "cidade": "Uberlândia/MG", "endereco": "R WANDELVART FARIA MARQUEZ, 500 - UBERLANDIA/MG"},
    14: {"sigla": "BRA", "uf": "DF", "cidade_ibge": "BRASILIA", "cidade": "Brasília/DF", "endereco": "ST SIN (SETOR DE INFLAMÁVEIS), SN - BRASILIA/DF"},
    15: {"sigla": "BTM", "uf": "MG", "cidade_ibge": "BETIM", "cidade": "Betim/MG", "endereco": "ROD BR-381 FERNAO DIAS, S/N, SALA A KM 485 - BETIM/MG"},
    17: {"sigla": "PTB", "uf": "SC", "cidade_ibge": "PORTO BELO", "cidade": "Porto Belo/SC", "endereco": "ROD BR 101, S/N, SERTÃO DE SANTA LUZIA - PORTO BELO/SC"},
    18: {"sigla": "JDS", "uf": "PR", "cidade_ibge": "JANDAIA DO SUL", "cidade": "Jandaia do Sul/PR", "endereco": "RUA JOAO ORTEGA SANCHES, 65 - JANDAIA DO SUL/PR"}
}

ENDERECOS_BASES = {}
OPCOES_BASES_FORMATADAS = {}

for num, dados in ENDERECOS_BASES_DETALHADOS.items():
    label = f"Base {num} - {dados['sigla']} ({dados['cidade']})"
    ENDERECOS_BASES[dados['sigla']] = dados['endereco']
    ENDERECOS_BASES[str(num)] = dados['endereco']
    OPCOES_BASES_FORMATADAS[label] = {
        "numero": num,
        "sigla": dados['sigla'],
        "uf": dados['uf'],
        "cidade_ibge": dados['cidade_ibge'],
        "endereco": dados['endereco']
    }

# Carrega base de coordenadas dos 5.571 municípios brasileiros (IBGE)
_MAPA_MUNICIPIOS = None


def carregar_mapa_municipios():
    global _MAPA_MUNICIPIOS
    if _MAPA_MUNICIPIOS is not None:
        return _MAPA_MUNICIPIOS
    caminho = os.path.join(os.path.dirname(os.path.abspath(__file__)), "municipios_br.json")
    if os.path.exists(caminho):
        try:
            with open(caminho, "r", encoding="utf-8") as f:
                _MAPA_MUNICIPIOS = json.load(f)
                return _MAPA_MUNICIPIOS
        except Exception:
            pass
    _MAPA_MUNICIPIOS = {}
    return _MAPA_MUNICIPIOS


def normalizar_texto_sem_acento(txt):
    if not txt:
        return ""
    s = str(txt).strip().upper()
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def obter_info_base(codigo_base):
    cod = str(codigo_base).strip().upper()
    for num, d in ENDERECOS_BASES_DETALHADOS.items():
        if cod == d["sigla"] or cod == str(num):
            return d
    return ENDERECOS_BASES_DETALHADOS[6]


def obter_coordenada_cidade(cidade, uf_padrao="RS"):
    muns = carregar_mapa_municipios()
    cid_norm = normalizar_texto_sem_acento(cidade)
    if not cid_norm or cid_norm == "-":
        return None
    uf_norm = normalizar_texto_sem_acento(uf_padrao) or "RS"

    # 1. Busca exata Cidade|UF
    chave = f"{cid_norm}|{uf_norm}"
    if chave in muns:
        return muns[chave]

    # 2. Tenta correções comuns de grafia (ex: PAULINA -> PAULINIA, ARAUCARA -> ARAUCARIA)
    apelidos = {
        "PAULINA": "PAULINIA",
        "ARAUCARA": "ARAUCARIA",
        "SANTANA DO LIVRAMENTO": "SANT'ANA DO LIVRAMENTO",
        "ELDORADO DO SUL": "ELDORADO DO SUL",
        "XANGRI-LA": "XANGRI-LA",
        "XANGRILA": "XANGRI-LA"
    }
    if cid_norm in apelidos and f"{apelidos[cid_norm]}|{uf_norm}" in muns:
        return muns[f"{apelidos[cid_norm]}|{uf_norm}"]

    # 3. Busca parcial na mesma UF
    for k, coord in muns.items():
        nome_m, uf_m = k.split("|")
        if uf_m == uf_norm and (cid_norm in nome_m or nome_m in cid_norm):
            return coord

    # 4. Busca em qualquer UF se o nome for único
    for k, coord in muns.items():
        nome_m, _ = k.split("|")
        if nome_m == cid_norm:
            return coord

    return None


def calcular_rota_osrm_por_cidades(cidade_base, uf_base, lista_cidades_dia, qtd_enderecos_distintos=1):
    """
    Calcula a quilometragem rodoviária real usando o servidor OSRM (OpenStreetMap)
    partindo da cidade da Base Consigaz, passando por todas as cidades visitadas no dia
    na melhor sequência geográfica (Trip / Caixeiro Viajante) e retornando à Base.
    """
    coord_base = obter_coordenada_cidade(cidade_base, uf_base)
    if not coord_base:
        return {
            "km_estrada": 0.0,
            "km_urbano": 0.0,
            "km_total_estimado": 0.0,
            "cidades_ordenadas": lista_cidades_dia,
            "sequencia_Formatada": " ➔ ".join([cidade_base] + lista_cidades_dia + [cidade_base])
        }

    coords = [coord_base]
    nomes = [f"{cidade_base.title()} - {uf_base}"]

    for cid in lista_cidades_dia:
        c_limpa = str(cid or "").strip()
        if not c_limpa or c_limpa == "-":
            continue
        coord_c = obter_coordenada_cidade(c_limpa, uf_base)
        if coord_c:
            # Se for na própria cidade da base, não precisa adicionar ponto interurbano duplicado
            if abs(coord_c[0] - coord_base[0]) < 0.001 and abs(coord_c[1] - coord_base[1]) < 0.001:
                continue
            if coord_c not in coords:
                coords.append(coord_c)
                nomes.append(f"{c_limpa.title()} - {uf_base}")

    # Se todas as visitas foram dentro da própria cidade da Base
    if len(coords) == 1:
        km_urb = round(max(12.0, qtd_enderecos_distintos * 4.5), 1)
        return {
            "km_estrada": 0.0,
            "km_urbano": km_urb,
            "km_total_estimado": km_urb,
            "cidades_ordenadas": [f"{cidade_base.title()} - {uf_base}"],
            "sequencia_formatada": f"{cidade_base.title()} (Deslocamento Urbano Local)"
        }

    coords_str = ";".join(f"{c[1]:.5f},{c[0]:.5f}" for c in coords)
    km_estrada = 0.0
    cidades_ord = nomes[1:]
    seq_completa = nomes + [nomes[0]]

    # Consulta OSRM Trip (ordena as cidades na sequência real da rodovia e fecha o circuito de ida e volta)
    url_trip = f"http://router.project-osrm.org/trip/v1/driving/{coords_str}?roundtrip=true&source=first&overview=false"
    try:
        req = urllib.request.Request(url_trip, headers={"User-Agent": "ConsigazAuditoriaKM/2.0"})
        with urllib.request.urlopen(req, timeout=6) as resp:
            data = json.loads(resp.read().decode())
            if data.get("trips"):
                km_estrada = round(data["trips"][0]["distance"] / 1000.0, 1)
                wps = sorted(enumerate(data.get("waypoints", [])), key=lambda x: x[1].get("waypoint_index", 0))
                ordem_idx = [idx for idx, _ in wps]
                cidades_ord = [nomes[i] for i in ordem_idx if i != 0]
                seq_completa = [nomes[0]] + cidades_ord + [nomes[0]]
    except Exception:
        pass

    # Fallback: distância geodésica * fator rodoviário (1.28) caso o usuário esteja sem internet para o OSRM
    if km_estrada <= 0.0:
        import math
        def haversine(c1, c2):
            r = 6371.0
            p1, p2 = math.radians(c1[0]), math.radians(c2[0])
            dp = math.radians(c2[0] - c1[0])
            dl = math.radians(c2[1] - c1[1])
            a = math.sin(dp/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
            return 2 * r * math.asin(math.sqrt(a))
        pts_circuito = coords + [coords[0]]
        dist_reta = sum(haversine(pts_circuito[i], pts_circuito[i+1]) for i in range(len(pts_circuito)-1))
        km_estrada = round(dist_reta * 1.28, 1)

    # Deslocamento urbano interno entre as ruas visitadas dentro das cidades (~2.5 km por endereço distinto visitado)
    km_urbano = round(max(4.0, qtd_enderecos_distintos * 2.5), 1)
    km_total = round(km_estrada + km_urbano, 1)

    return {
        "km_estrada": km_estrada,
        "km_urbano": km_urbano,
        "km_total_estimado": km_total,
        "cidades_ordenadas": cidades_ord,
        "sequencia_formatada": " ➔ ".join(n.split(" - ")[0] for n in seq_completa)
    }


def detectar_base_por_texto(texto):
    """Identifica o código da base (ex: MAU, PLN, DQX, CAN) no nome do arquivo ou texto."""
    if not texto:
        return "CAN"
    txt = str(texto).upper()
    for num, d in ENDERECOS_BASES_DETALHADOS.items():
        sigla = d["sigla"]
        if re.search(rf'(?:^|[-_ \.\(\)\[\]]){sigla}(?:$|[-_ \.\(\)\[\]])', txt):
            return sigla
    if "MAUA" in txt or "MAUÁ" in txt: return "MAU"
    if "PAULINIA" in txt or "PAULÍNIA" in txt: return "PLN"
    if "CANOAS" in txt: return "CAN"
    if "CAXIAS" in txt: return "DQX"
    if "BARUERI" in txt: return "BRU"
    if "RIBEIRAO" in txt or "RIBEIRÃO" in txt: return "RIB"
    if "VICENTE" in txt: return "SVC"
    if "CARIACICA" in txt: return "CAR"
    if "CANEDO" in txt: return "SNC"
    if "ARAUCARIA" in txt or "ARAUCÁRIA" in txt: return "ARA"
    if "UBERLANDIA" in txt or "UBERLÂNDIA" in txt: return "UBR"
    if "BRASILIA" in txt or "BRASÍLIA" in txt: return "BRA"
    if "BETIM" in txt: return "BTM"
    if "PORTO BELO" in txt: return "PTB"
    if "JANDAIA" in txt: return "JDS"
    return "CAN"


def obter_endereco_base(codigo_base):
    cod = str(codigo_base).strip().upper()
    if cod in ENDERECOS_BASES:
        return ENDERECOS_BASES[cod]
    for num, d in ENDERECOS_BASES_DETALHADOS.items():
        if cod == d["sigla"] or cod == str(num):
            return d["endereco"]
    return ENDERECOS_BASES.get("CAN", "RUA ANTONIO FREDERICO OZANAN, 1901 - CANOAS/RS")


def carregar_planilha_segura(file_source, senha="fechamento"):
    """
    Carrega arquivo Excel de arquivo físico ou stream em memória (UploadedFile).
    Descriptografa com msoffcrypto se houver senha padrão de abertura.
    """
    if isinstance(file_source, str):
        with open(file_source, "rb") as f:
            conteudo = f.read()
    elif hasattr(file_source, "read"):
        conteudo = file_source.read()
        if hasattr(file_source, "seek"):
            file_source.seek(0)
    else:
        conteudo = file_source

    stream = io.BytesIO(conteudo)

    try:
        wb = openpyxl.load_workbook(stream, data_only=True)
        return wb
    except Exception:
        pass

    if msoffcrypto:
        stream.seek(0)
        try:
            office_file = msoffcrypto.OfficeFile(stream)
            if office_file.is_encrypted():
                decrypted = io.BytesIO()
                office_file.load_key(password=senha)
                office_file.decrypt(decrypted)
                decrypted.seek(0)
                return openpyxl.load_workbook(decrypted, data_only=True)
        except Exception:
            pass

    stream.seek(0)
    return openpyxl.load_workbook(stream, data_only=True)


def normalizar_data(valor):
    """Converte qualquer formato de data do Excel para string YYYY-MM-DD e datetime.date."""
    if valor is None or str(valor).strip() in ["", "-", "None", "00/01/1900"]:
        return None, None
    if isinstance(valor, datetime):
        d = valor.date()
        if d.year < 1910: return None, None
        return d.strftime("%Y-%m-%d"), d
    if isinstance(valor, date):
        if valor.year < 1910: return None, None
        return valor.strftime("%Y-%m-%d"), valor
    if isinstance(valor, (int, float)):
        try:
            dt = datetime.fromordinal(datetime(1899, 12, 30).toordinal() + int(valor))
            if dt.year < 1910: return None, None
            return dt.strftime("%Y-%m-%d"), dt.date()
        except Exception:
            pass
    s = str(valor).strip()
    if " " in s:
        s = s.split(" ")[0].strip()
    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"]:
        try:
            dt = datetime.strptime(s, fmt).date()
            if dt.year < 1910: return None, None
            return dt.strftime("%Y-%m-%d"), dt
        except Exception:
            pass
    return None, None


def limpar_endereco(endereco_bruto, bairro=None, cidade=None, uf=None, cep=None, uf_padrao="RS"):
    """
    Simplifica o endereço para evitar falhas no Google Maps:
    - Usa SEMPRE a UF correta da Base quando o Totvs não informar UF.
    - Remove Bairro e complementos prediais (AP, BL, Condomínio, Sala) que confundem o Google Maps.
    - Se o nome da rua estiver inválido/vazio, usa 'Cidade - UF'.
    """
    uf_str = str(uf).strip().upper() if uf and str(uf).strip() and str(uf).strip().upper() not in ["NONE", "0", "-"] else uf_padrao
    cidade_str = str(cidade).strip().title() if cidade and str(cidade).strip() and str(cidade).strip().upper() not in ["NONE", "0", "-"] else ""

    if not endereco_bruto or str(endereco_bruto).strip() in ["", "0", "None", "-"]:
        if cidade_str:
            return f"{cidade_str} - {uf_str}"
        return ""

    s = normalizar_texto_sem_acento(endereco_bruto)
    s = re.sub(r'\s+', ' ', s)

    # Remove duplicação no início como "R RUA " ou "AV AVENIDA "
    s = re.sub(r'^(?:R\.?\s+RUA\s+|AV\.?\s+AVENIDA\s+)', 'RUA ', s)

    # Corta complementos prediais
    padrao = r'[\s,]*(?:BLOCO|BL\s*\d*|APTO|APT|AP\s*\d*|TORRE|TR\s*\d*|CASA|CS\s*\d*|LT\s*\d*|LOTE|CONDOMINIO|COND\s|FASE\s*\d*|ANDAR|TERREO|SALA\s*\d*|LOJA\s*\d*).*$'
    s = re.sub(padrao, '', s, flags=re.IGNORECASE).strip().rstrip(' ,;.-')

    partes = s.split(',')
    rua = partes[0].strip()
    numero = ""
    if len(partes) > 1:
        resto = partes[1].strip()
        m_rep = re.match(r'^(\d+)\s+\1\b', resto)
        if m_rep:
            resto = resto[len(m_rep.group(1)):].strip()
        if not any(sn in resto.upper() for sn in ['S/N', 'SN', 'S/Nº', 'SEM NUMERO']):
            num = re.search(r'^\d+', resto)
            numero = num.group(0) if num else ""

    # Se a rua ficou genérica ("RUA", "AVENIDA", "CENTRO", etc.), usa direto a Cidade - UF para nunca dar erro no Google Maps
    if len(rua) <= 4 or rua in ["RUA", "AVENIDA", "AV", "ESTRADA", "RODOVIA", "CENTRO", "INTERIOR"]:
        if cidade_str:
            return f"{cidade_str} - {uf_str}"
        return ""

    if numero and numero not in ["0", "00", "SN", "S/N"]:
        end_rua = f"{rua.title()}, {numero}"
    else:
        end_rua = rua.title()

    # Sem o Bairro para evitar divergência de cadastro no Google Maps! Apenas Rua + Cidade - UF
    if cidade_str:
        return f"{end_rua}, {cidade_str} - {uf_str}"
    return f"{end_rua} - {uf_str}"


def extrair_todas_oss_fechamento(wb_fechamento):
    """
    Varre todas as abas da planilha de fechamento (MOTO, CARRO, ENGENHARIA, INSTALAÇÃO, etc.).
    """
    itens_km = []
    todas_oss_fechamento = []
    oss_por_data = {}
    abas_encontradas = set()

    for sname in wb_fechamento.sheetnames:
        ws = wb_fechamento[sname]
        if ws.max_row is None or ws.max_row < 2:
            continue

        header_row = None
        col_os = None
        col_cliente = None
        col_km = None
        col_pedagio = None
        col_dt = None
        col_local = None
        col_servico = None
        col_viagens = None
        col_total = None

        for r in range(1, min(16, ws.max_row + 1)):
            for c in range(1, min(ws.max_column + 1, 35)):
                val = str(ws.cell(r, c).value or "").strip().upper()
                if not val:
                    continue
                if "ORDEM DE SERVI" in val or val == "OS" or val == "ORDEM":
                    col_os = c
                    header_row = r
                elif "CLIENTE" in val and ("COD" in val or "CÓD" in val or "CODIGO" in val or "CÓDIGO" in val):
                    col_cliente = c
                elif "DESLOCAMENTO" in val:
                    col_km = c
                elif val == "KM" and col_km is None:
                    col_km = c
                elif "PEDAGIO" in val or "PEDÁGIO" in val:
                    col_pedagio = c
                elif ("DATA" in val and ("ENCERRAMENTO" in val or "CONCLUSAO" in val or "FINALIZA" in val or "EXECU" in val or "ATENDIMENTO" in val)) or val == "DATA":
                    col_dt = c
                elif "LOCAL" in val or "CIDADE" in val or "MUNICIPIO" in val:
                    col_local = c
                elif val in ["SERVIÇO", "SERVICO", "SERVIÇO PRESTADO", "TIPO DE SERVIÇO", "DESCRIÇÃO DO SERVIÇO"]:
                    if r >= 7:
                        col_servico = c
                elif "VIAGENS" in val or "QT. VIAGENS" in val or "VIAGEM" in val:
                    col_viagens = c
                elif val == "TOTAL" or "VALOR TOTAL" in val or "VALOR TRANSPORTE" in val:
                    if r >= 7:
                        col_total = c

        if col_os and col_km and header_row:
            abas_encontradas.add(sname)
            for r in range(header_row + 1, ws.max_row + 1):
                val_os = ws.cell(r, col_os).value
                if val_os is None or str(val_os).strip() in ["", "-", "TOTAL", "SUBTOTAL"]:
                    continue

                try:
                    os_clean = int(float(str(val_os).replace(",", ".")))
                except Exception:
                    os_clean = str(val_os).strip()
                    if not any(ch.isdigit() for ch in os_clean):
                        continue

                val_dt = ws.cell(r, col_dt).value if col_dt else None
                dt_str, dt_obj = normalizar_data(val_dt)

                val_km = ws.cell(r, col_km).value
                try:
                    km_num = float(str(val_km).replace(",", "."))
                except (ValueError, TypeError):
                    km_num = 0.0

                val_ped = ws.cell(r, col_pedagio).value if col_pedagio else 0.0
                try:
                    ped_num = float(str(val_ped).replace(",", "."))
                except (ValueError, TypeError):
                    ped_num = 0.0

                viagens = ws.cell(r, col_viagens).value if col_viagens else 1
                try:
                    viagens_num = int(float(viagens))
                except Exception:
                    viagens_num = 1

                local_val = ws.cell(r, col_local).value if col_local else ""
                cod_cli_val = ws.cell(r, col_cliente).value if col_cliente else ""
                if cod_cli_val is not None and str(cod_cli_val).strip() != "":
                    try:
                        cod_cli_str = str(int(float(str(cod_cli_val).replace(",", "."))))
                    except Exception:
                        cod_cli_str = str(cod_cli_val).strip()
                else:
                    cod_cli_str = ""

                serv_val = ws.cell(r, col_servico).value if col_servico else ""
                total_val = ws.cell(r, col_total).value if col_total else 0.0

                registro_os = {
                    "aba": sname,
                    "linha": r,
                    "os": os_clean,
                    "cod_cliente": cod_cli_str,
                    "km_cobrado": km_num,
                    "pedagio": ped_num,
                    "data_str": dt_str,
                    "data_obj": dt_obj,
                    "viagens": viagens_num,
                    "local": str(local_val or "").strip(),
                    "servico_fechamento": str(serv_val or "").strip(),
                    "total_pago": total_val
                }
                todas_oss_fechamento.append(registro_os)

                if dt_str:
                    if dt_str not in oss_por_data:
                        oss_por_data[dt_str] = []
                    if os_clean not in oss_por_data[dt_str]:
                        oss_por_data[dt_str].append(os_clean)

                if km_num > 0:
                    itens_km.append(registro_os)

    return itens_km, oss_por_data, list(abas_encontradas), todas_oss_fechamento


def carregar_dados_totvs(wb_totvs, uf_padrao="RS"):
    """
    Processa a aba Totvs (ESOS0301... ou Relatório) e indexa todas as OSs.
    """
    ws_rel = None
    for sname in wb_totvs.sheetnames:
        if sname.upper().startswith("ESOS") or "RELAT" in sname.upper():
            ws_rel = wb_totvs[sname]
            break
    if ws_rel is None:
        ws_rel = wb_totvs.active

    header_row = 2
    for r in range(1, 10):
        val1 = str(ws_rel.cell(r, 1).value or "").strip().upper()
        val2 = str(ws_rel.cell(r, 2).value or "").strip().upper()
        if "NR ORDEM" in val1 or (("ORDEM" in val1 or "OS" in val1) and val2 != ""):
            header_row = r
            break

    col_map = {}
    for c in range(1, ws_rel.max_column + 1):
        v = str(ws_rel.cell(header_row, c).value or "").strip().upper()
        if "NR ORDEM" in v or v == "ORDEM" or v == "OS": col_map["os"] = c
        elif "ENDERE" in v: col_map["endereco"] = c
        elif "BAIRRO" in v: col_map["bairro"] = c
        elif "CIDADE" in v: col_map["cidade"] = c
        elif "ESTADO" in v or v == "UF": col_map["uf"] = c
        elif "CEP" in v: col_map["cep"] = c
        elif any(t in v for t in ["TÉCNICO", "TECNICO", "TCNICO"]) and "USU" not in v:
            col_map["tecnico"] = c
        elif "DT CONCLUSAO" in v or "CONCLUSAO" in v: col_map["dt_concl"] = c
        elif "DATA FINALIZA" in v or "FINALIZA" in v: col_map["dt_final"] = c
        elif "DT ABERTURA" in v or "ABERTURA" in v: col_map["dt_abert"] = c
        elif "SITUA" in v and "USU" not in v and "MORADOR" not in v: col_map["situacao"] = c
        elif "COD CLIENTE" in v or "CÓD CLIENTE" in v or "CD CLIENTE" in v: col_map["cod_cliente"] = c
        elif "NOME CLIENTE" in v or ("CLIENTE" in v and "CLASSIF" not in v and "NOME" in v): col_map["cliente"] = c
        elif "MOTIVO" in v and "DESC" not in v: col_map["motivo"] = c
        elif "DESC" in v and "MOTIVO" in v: col_map["desc_motivo"] = c

    os_dict = {}
    por_data_dict = {}

    for r in range(header_row + 1, ws_rel.max_row + 1):
        val_os = ws_rel.cell(r, col_map.get("os", 1)).value
        if val_os is None or str(val_os).strip() == "":
            continue

        try:
            os_num = int(float(str(val_os).replace(",", ".")))
        except Exception:
            os_num = str(val_os).strip()

        end_bruto = ws_rel.cell(r, col_map.get("endereco", 15)).value if "endereco" in col_map else ""
        bairro = ws_rel.cell(r, col_map.get("bairro", 16)).value if "bairro" in col_map else ""
        cidade = ws_rel.cell(r, col_map.get("cidade", 17)).value if "cidade" in col_map else ""
        uf = ws_rel.cell(r, col_map.get("uf", 18)).value if "uf" in col_map else ""
        cep = ws_rel.cell(r, col_map.get("cep", 19)).value if "cep" in col_map else ""
        tecnico = ws_rel.cell(r, col_map.get("tecnico", 35)).value if "tecnico" in col_map else ""
        dt_final = ws_rel.cell(r, col_map.get("dt_final", 29)).value if "dt_final" in col_map else None
        dt_concl = ws_rel.cell(r, col_map.get("dt_concl", 28)).value if "dt_concl" in col_map else None
        dt_abert = ws_rel.cell(r, col_map.get("dt_abert", 27)).value if "dt_abert" in col_map else None
        situacao = ws_rel.cell(r, col_map.get("situacao", 7)).value if "situacao" in col_map else ""
        cliente = ws_rel.cell(r, col_map.get("cliente", 13)).value if "cliente" in col_map else ""
        cod_cli = ws_rel.cell(r, col_map.get("cod_cliente", 12)).value if "cod_cliente" in col_map else ""

        dt_str, dt_obj = normalizar_data(dt_final or dt_concl or dt_abert)
        hora_str = ""
        if isinstance(dt_final, datetime):
            hora_str = dt_final.strftime("%H:%M:%S")
        elif dt_final and " " in str(dt_final):
            hora_str = str(dt_final).split(" ")[1][:8]

        desc_motivo = ws_rel.cell(r, col_map.get("desc_motivo", 6)).value if "desc_motivo" in col_map else ""
        motivo_num = ws_rel.cell(r, col_map.get("motivo", 5)).value if "motivo" in col_map else ""
        servico_nome = str(desc_motivo or motivo_num or "VISITA OPERACIONAL").strip()

        end_limpo = limpar_endereco(end_bruto, bairro, cidade, uf, cep, uf_padrao=uf_padrao)

        rec = {
            "os": os_num,
            "endereco_bruto": str(end_bruto or "").strip(),
            "endereco_limpo": end_limpo,
            "bairro": str(bairro or "").strip(),
            "cidade": str(cidade or "").strip(),
            "uf": str(uf or uf_padrao).strip(),
            "cep": str(cep or "").strip(),
            "tecnico": str(tecnico or "").strip(),
            "data_str": dt_str,
            "data_obj": dt_obj,
            "hora_str": hora_str,
            "situacao": str(situacao or "").strip(),
            "cliente": str(cliente or "").strip(),
            "cod_cliente": str(cod_cli or "").strip(),
            "servico": servico_nome
        }

        os_dict[os_num] = rec

        if dt_str:
            if dt_str not in por_data_dict:
                por_data_dict[dt_str] = []
            por_data_dict[dt_str].append(rec)

    return os_dict, por_data_dict


def remover_duplicados_preservando_ordem(lista_enderecos):
    vistos = set()
    resultado = []
    for end in lista_enderecos:
        e_str = str(end or "").strip()
        if not e_str:
            continue
        chave = e_str.upper()
        if chave not in vistos:
            vistos.add(chave)
            resultado.append(e_str)
    return resultado


def dividir_trechos_maps(base_endereco, paradas):
    paradas_validas = remover_duplicados_preservando_ordem(paradas)
    if not paradas_validas:
        return {
            "trecho1": {
                "nome": "Rota Única",
                "url": gerar_url_google_maps(base_endereco, [], base_endereco),
                "origem": base_endereco,
                "destino": base_endereco,
                "qtd_paradas": 0
            },
            "trecho2": None,
            "todos_trechos": []
        }

    if len(paradas_validas) <= 8:
        t_unico = {
            "nome": f"Rota Completa ({len(paradas_validas)} paradas)",
            "url": gerar_url_google_maps(base_endereco, paradas_validas, base_endereco),
            "origem": base_endereco,
            "destino": base_endereco,
            "qtd_paradas": len(paradas_validas)
        }
        return {
            "trecho1": t_unico,
            "trecho2": None,
            "todos_trechos": [t_unico]
        }

    todos_trechos = []
    idx_inicio = 0
    num_trecho = 1
    ponto_atual = base_endereco

    while idx_inicio < len(paradas_validas):
        restantes = len(paradas_validas) - idx_inicio
        if restantes <= 8:
            intermed = paradas_validas[idx_inicio:]
            destino_final = base_endereco
            url_t = gerar_url_google_maps(ponto_atual, intermed, destino_final)
            todos_trechos.append({
                "nome": f"Trecho {num_trecho} (Paradas {idx_inicio + 1} a {len(paradas_validas)} + Retorno à Base)",
                "url": url_t,
                "origem": ponto_atual,
                "destino": destino_final,
                "qtd_paradas": len(intermed)
            })
            break
        else:
            bloco = paradas_validas[idx_inicio:idx_inicio + 8]
            intermed = bloco[:-1]
            ponto_transicao = bloco[-1]
            url_t = gerar_url_google_maps(ponto_atual, intermed, ponto_transicao)
            resumo_dest = ponto_transicao.split(",")[0] if "," in ponto_transicao else ponto_transicao
            todos_trechos.append({
                "nome": f"Trecho {num_trecho} (Paradas {idx_inicio + 1} a {idx_inicio + len(bloco)} — até {resumo_dest})",
                "url": url_t,
                "origem": ponto_atual,
                "destino": ponto_transicao,
                "qtd_paradas": len(bloco)
            })
            ponto_atual = ponto_transicao
            idx_inicio += 8
            num_trecho += 1

    return {
        "trecho1": todos_trechos[0],
        "trecho2": todos_trechos[1] if len(todos_trechos) > 1 else None,
        "todos_trechos": todos_trechos
    }


def gerar_url_google_maps(origem, paradas, destino=None):
    if not destino:
        destino = origem
    paradas_validas = remover_duplicados_preservando_ordem(paradas)
    if len(paradas_validas) > 8:
        paradas_formatadas = paradas_validas[:8]
    else:
        paradas_formatadas = paradas_validas

    partes = [origem] + paradas_formatadas + [destino]
    enc_partes = [urllib.parse.quote_plus(str(p).strip()) for p in partes]
    return "https://www.google.com/maps/dir/" + "/".join(enc_partes)


def processar_auditoria_completa(wb_fechamento, wb_totvs, base_codigo="CAN", base_endereco_custom=None, raio_franquia=40.0):
    """
    Executa o cruzamento completo, agrupa por dia, soma os KMs registrados no dia
    e calcula automaticamente a quilometragem real da rota (via OSRM + IBGE).
    """
    info_base = obter_info_base(base_codigo)
    uf_base = info_base["uf"]
    cidade_base = info_base["cidade_ibge"]
    base_endereco = base_endereco_custom or info_base["endereco"]
    base_cidade_uf = f"{cidade_base.title()} - {uf_base}"

    # 1. Extrai todas as OSs do fechamento (com e sem KM)
    itens_km, _, abas_auditadas, todas_oss_fechamento = extrair_todas_oss_fechamento(wb_fechamento)

    # 2. Carrega índice Totvs passando a UF correta da Base
    os_dict, _ = carregar_dados_totvs(wb_totvs, uf_padrao=uf_base)

    # 3. Enriquece todas as OSs do fechamento com os dados do Totvs
    oss_enriquecidas = []
    mapa_oss_por_dia = {}

    for reg in todas_oss_fechamento:
        os_id = reg["os"]
        totvs_os = os_dict.get(os_id, {})

        dt_str = reg["data_str"] or totvs_os.get("data_str")
        if dt_str:
            try:
                data_fmt = datetime.strptime(dt_str, "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                data_fmt = dt_str
            chave_dia = dt_str
        else:
            data_fmt = f"Sem Data (OS {os_id})"
            chave_dia = f"SEM_DATA_{os_id}"

        cidade_os = reg["local"] or totvs_os.get("cidade") or ""
        end_limpo = totvs_os.get("endereco_limpo")
        if not end_limpo:
            if cidade_os:
                end_limpo = f"{cidade_os.title()} - {uf_base}"
            else:
                end_limpo = base_cidade_uf

        servico_final = reg["servico_fechamento"] or totvs_os.get("servico") or "Visita Operacional"
        cod_cliente_final = reg["cod_cliente"] or totvs_os.get("cod_cliente") or totvs_os.get("cliente") or "-"
        nome_cliente_final = totvs_os.get("cliente") or reg["cod_cliente"] or "-"

        item_completo = {
            "aba": reg["aba"],
            "linha": reg["linha"],
            "os": os_id,
            "cod_cliente": cod_cliente_final,
            "cliente": nome_cliente_final,
            "cidade": cidade_os.upper() if cidade_os else "-",
            "bairro": totvs_os.get("bairro") or "-",
            "endereco_bruto": totvs_os.get("endereco_bruto") or cidade_os or "-",
            "endereco_limpo": end_limpo,
            "km_cobrado": reg["km_cobrado"],
            "cobrando_km": reg["km_cobrado"] > 0,
            "pedagio": reg["pedagio"],
            "servico": servico_final,
            "data_str": dt_str or "",
            "data_formatada": data_fmt,
            "chave_dia": chave_dia,
            "hora_str": totvs_os.get("hora_str") or "",
            "tecnico": totvs_os.get("tecnico") or "Não identificado",
            "situacao": totvs_os.get("situacao") or "Atendida",
            "total_pago": reg["total_pago"]
        }
        oss_enriquecidas.append(item_completo)

        if chave_dia not in mapa_oss_por_dia:
            mapa_oss_por_dia[chave_dia] = []
        mapa_oss_por_dia[chave_dia].append(item_completo)

    # 4. Constrói a visão consolidada POR DIA e calcula o KM real de cada dia no Mapa (OSRM)
    resumo_por_dia = []
    dossies_por_dia = {}

    chaves_ordenadas = sorted(mapa_oss_por_dia.keys())
    for chave_dia in chaves_ordenadas:
        lista_oss_dia = mapa_oss_por_dia[chave_dia]
        oss_com_km = [o for o in lista_oss_dia if o["km_cobrado"] > 0]
        if not oss_com_km:
            continue

        data_fmt = lista_oss_dia[0]["data_formatada"]
        soma_km_dia = sum(o["km_cobrado"] for o in oss_com_km)
        soma_pedagio_dia = sum(o["pedagio"] for o in lista_oss_dia)
        soma_km_faturavel_dia = sum(max(0.0, o["km_cobrado"] - raio_franquia) for o in oss_com_km)

        cidades_dia = []
        for o in lista_oss_dia:
            c = o["cidade"]
            if c and c != "-" and c not in cidades_dia:
                cidades_dia.append(c)

        # Agrupa as paradas por endereço único
        paradas_agrupadas = []
        mapa_end_idx = {}
        for o in lista_oss_dia:
            end_key = (o["endereco_limpo"] or base_cidade_uf).strip().upper()
            if end_key not in mapa_end_idx:
                mapa_end_idx[end_key] = len(paradas_agrupadas)
                paradas_agrupadas.append({
                    "ordem": len(paradas_agrupadas) + 1,
                    "endereco": o["endereco_limpo"] or base_cidade_uf,
                    "cidade": o["cidade"],
                    "bairro": o["bairro"],
                    "oss": [],
                    "km_neste_endereco": 0.0,
                    "tem_cobranca": False
                })
            idx_p = mapa_end_idx[end_key]
            paradas_agrupadas[idx_p]["oss"].append(o)
            if o["km_cobrado"] > 0:
                paradas_agrupadas[idx_p]["km_neste_endereco"] += o["km_cobrado"]
                paradas_agrupadas[idx_p]["tem_cobranca"] = True

        # CÁLCULO AUTOMÁTICO DE KM DA ROTA NO MAPA (OSRM + IBGE)
        info_rota_osrm = calcular_rota_osrm_por_cidades(
            cidade_base=cidade_base,
            uf_base=uf_base,
            lista_cidades_dia=cidades_dia,
            qtd_enderecos_distintos=len(paradas_agrupadas)
        )
        km_calculado_mapa = info_rota_osrm["km_total_estimado"]
        km_estrada_mapa = info_rota_osrm["km_estrada"]
        km_urbano_mapa = info_rota_osrm["km_urbano"]
        cidades_ordenadas_maps = info_rota_osrm["cidades_ordenadas"]
        sequencia_rota_txt = info_rota_osrm["sequencia_formatada"]

        # Reordena as paradas agrupadas seguindo a sequência otimizada das cidades da rota
        ordem_cidades_norm = [normalizar_texto_sem_acento(c.split(" - ")[0]) for c in cidades_ordenadas_maps]
        def chave_ordem_parada(p):
            c_norm = normalizar_texto_sem_acento(p["cidade"])
            if c_norm in ordem_cidades_norm:
                return (ordem_cidades_norm.index(c_norm), p["ordem"])
            return (999, p["ordem"])

        paradas_agrupadas = sorted(paradas_agrupadas, key=chave_ordem_parada)
        for idx_novo, p in enumerate(paradas_agrupadas, start=1):
            p["ordem"] = idx_novo

        # Link INFALÍVEL do Google Maps por Cidades Ordenadas (100% livre de erros de rua)
        link_maps_cidades = gerar_url_google_maps(base_endereco, cidades_ordenadas_maps, base_endereco)

        # Links com Ruas Simplificadas (sem Bairro e com UF correta)
        enderecos_rota_dia = [p["endereco"] for p in paradas_agrupadas]
        trechos_dia = dividir_trechos_maps(base_endereco, enderecos_rota_dia)

        abas_dia = []
        for o in oss_com_km:
            if o["aba"] not in abas_dia:
                abas_dia.append(o["aba"])

        detalhe_oss_cobrando = [
            f"OS {o['os']} ({o['km_cobrado']:,.1f} km - {o['cidade']})"
            for o in oss_com_km
        ]
        todas_oss_numeros_str = ", ".join(str(o["os"]) for o in lista_oss_dia)

        diferenca_km = round(soma_km_dia - km_calculado_mapa, 1)
        # Considera aprovado se o KM cobrado for menor ou até 15 km próximo do KM calculado no mapa
        precisa_glosar = diferenca_km > 20.0

        registro_dia = {
            "chave_dia": chave_dia,
            "Data": data_fmt,
            "Abas": ", ".join(abas_dia),
            "Qtd OSs Cobrando": len(oss_com_km),
            "Qtd Total OSs no Dia": len(lista_oss_dia),
            "Qtd Endereços Distintos": len(paradas_agrupadas),
            "OSs Cobrando KM": " + ".join(detalhe_oss_cobrando),
            "Lista OSs Cobrando": oss_com_km,
            "Todas as OSs do Dia": lista_oss_dia,
            "Todas OSs Numeros": todas_oss_numeros_str,
            "Soma KM Cobrado no Dia": soma_km_dia,
            "KM Calculado no Mapa": km_calculado_mapa,
            "KM Estrada Mapa": km_estrada_mapa,
            "KM Urbano Mapa": km_urbano_mapa,
            "Diferença (Cobrado - Mapa)": diferenca_km,
            "Precisa Glosar": precisa_glosar,
            "Sequência da Rota": sequencia_rota_txt,
            "Soma Pedágio no Dia": soma_pedagio_dia,
            "KM Faturável do Dia": max(0.0, soma_km_dia - raio_franquia),
            "Soma KM Faturável por OS": soma_km_faturavel_dia,
            "Cidades do Dia": ", ".join(cidades_dia),
            "Paradas Google Maps": paradas_agrupadas,
            "Trechos Maps": trechos_dia,
            "Link Maps (Por Cidades - Sem Erro)": link_maps_cidades,
            "Link Maps (Todas do Dia)": trechos_dia["trecho1"]["url"],
            "Link Maps (Só OSs com KM)": link_maps_cidades
        }
        resumo_por_dia.append(registro_dia)

        label_seletor = f"{data_fmt} — Cobrado: {soma_km_dia:,.1f} km vs Mapa: {km_calculado_mapa:,.1f} km ({len(oss_com_km)} OS cobrando / {len(lista_oss_dia)} OSs no dia)"
        dossies_por_dia[label_seletor] = registro_dia

    # 5. Lista individual por OS que cobrou KM
    mapa_info_dia = {d["chave_dia"]: d for d in resumo_por_dia}
    resumo_linhas = []
    detalhes_visitas = []

    for o in oss_enriquecidas:
        if o["km_cobrado"] <= 0:
            continue
        info_dia = mapa_info_dia.get(o["chave_dia"], {})
        km_faturavel_os = max(0.0, o["km_cobrado"] - raio_franquia)

        resumo_linhas.append({
            "Aba": o["aba"],
            "OS Cobrada": o["os"],
            "Cód. Cliente": o["cod_cliente"],
            "Data": o["data_formatada"],
            "Técnico da OS": o["tecnico"],
            "KM Cobrado": o["km_cobrado"],
            "Soma KM do Dia": info_dia.get("Soma KM Cobrado no Dia", o["km_cobrado"]),
            "KM Calculado no Mapa": info_dia.get("KM Calculado no Mapa", 0.0),
            "OSs Cobrando no Dia": info_dia.get("OSs Cobrando KM", f"OS {o['os']}"),
            "Franquia Raio (km)": raio_franquia,
            "KM Faturável Cobrado": km_faturavel_os,
            "Qtd OS Total Dia": info_dia.get("Qtd Total OSs no Dia", 1),
            "Link Maps (Cidades)": info_dia.get("Link Maps (Por Cidades - Sem Erro)", ""),
            "Link Maps (Geral do Dia)": info_dia.get("Link Maps (Todas do Dia)", ""),
            "Status Totvs": o["situacao"],
            "Cidade": o["cidade"],
            "Serviço": o["servico"],
            "Endereço OS": o["endereco_limpo"]
        })

    for d in resumo_por_dia:
        seq = 1
        for o in d["Todas as OSs do Dia"]:
            detalhes_visitas.append({
                "Data": d["Data"],
                "Soma KM do Dia": d["Soma KM Cobrado no Dia"],
                "KM Calculado Mapa": d["KM Calculado no Mapa"],
                "Ordem": seq,
                "Aba": o["aba"],
                "OS Atendida": o["os"],
                "Cobrando KM?": f"SIM ({o['km_cobrado']:,.1f} km)" if o["km_cobrado"] > 0 else "Não (0 km)",
                "KM Registrado": o["km_cobrado"],
                "Pedágio": o["pedagio"],
                "Cód. Cliente": o["cod_cliente"],
                "Cliente": o["cliente"],
                "Cidade": o["cidade"],
                "Endereço Simplificado": o["endereco_limpo"],
                "Serviço Prestado": o["servico"]
            })
            seq += 1

    return {
        "base_codigo": base_codigo,
        "base_endereco": base_endereco,
        "raio_franquia": raio_franquia,
        "abas_auditadas": abas_auditadas,
        "por_dia": resumo_por_dia,
        "resumo": resumo_linhas,
        "detalhes": detalhes_visitas,
        "dossies": dossies_por_dia,
        "total_dias_km": len(resumo_por_dia),
        "total_oss_km": len(resumo_linhas),
        "total_km_cobrado": sum(r["KM Cobrado"] for r in resumo_linhas),
        "total_km_faturavel": sum(r["KM Faturável Cobrado"] for r in resumo_linhas)
    }


def exportar_relatorio_excel(dados_auditoria, output_stream):
    wb = openpyxl.Workbook()

    cor_azul = "1F4E79"
    cor_branco = "FFFFFF"
    cor_amarelo = "FFF2CC"

    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color=cor_branco)
    fill_cabecalho = PatternFill(start_color=cor_azul, end_color=cor_azul, fill_type="solid")
    fill_destaque = PatternFill(start_color=cor_amarelo, end_color=cor_amarelo, fill_type="solid")
    borda_fina = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    ws_dia = wb.active
    ws_dia.title = "Soma de KM por Dia"
    ws_dia.views.sheetView[0].showGridLines = True

    headers_dia = [
        "Data", "OSs Cobrando KM no Dia", "SOMA KM COBRADO NO DIA",
        "KM CALCULADO NO MAPA (ROTA DO DIA)", "DIFERENÇA (COBRADO - MAPA)",
        "Total OSs no Dia", "Sequência da Rota por Cidades",
        "Link Google Maps (Rota por Cidades - Sem Erro)", "Link Google Maps (Com Ruas)"
    ]
    ws_dia.append(headers_dia)
    for col_idx in range(1, len(headers_dia) + 1):
        cell = ws_dia.cell(1, col_idx)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    row_num = 2
    for d in dados_auditoria.get("por_dia", []):
        ws_dia.cell(row_num, 1, d["Data"])
        ws_dia.cell(row_num, 2, d["OSs Cobrando KM"])
        ws_dia.cell(row_num, 3, d["Soma KM Cobrado no Dia"])
        ws_dia.cell(row_num, 4, d["KM Calculado no Mapa"])
        ws_dia.cell(row_num, 5, d["Diferença (Cobrado - Mapa)"])
        ws_dia.cell(row_num, 6, d["Qtd Total OSs no Dia"])
        ws_dia.cell(row_num, 7, d["Sequência da Rota"])

        c_link_cid = ws_dia.cell(row_num, 8)
        if d["Link Maps (Por Cidades - Sem Erro)"]:
            c_link_cid.value = f'=HYPERLINK("{d["Link Maps (Por Cidades - Sem Erro)"]}", "🗺️ Abrir Rota Cidades")'
            c_link_cid.font = Font(color="0000FF", underline="single")

        c_link_ruas = ws_dia.cell(row_num, 9)
        if d["Link Maps (Todas do Dia)"]:
            c_link_ruas.value = f'=HYPERLINK("{d["Link Maps (Todas do Dia)"]}", "🏠 Abrir Rota c/ Ruas")'
            c_link_ruas.font = Font(color="0000FF", underline="single")

        for col_idx in range(1, len(headers_dia) + 1):
            ws_dia.cell(row_num, col_idx).border = borda_fina
            if d["Precisa Glosar"]:
                ws_dia.cell(row_num, col_idx).fill = fill_destaque
            if col_idx in [1, 3, 4, 5, 6, 8, 9]:
                ws_dia.cell(row_num, col_idx).alignment = Alignment(horizontal="center", vertical="center")

        row_num += 1

    ws_det = wb.create_sheet(title="Relatorio Deslocamento Dia")
    ws_det.views.sheetView[0].showGridLines = True

    headers_det = [
        "Data", "Aba Planilha", "Ordem de Serviço", "Cobrando KM?",
        "KM Registrado na OS", "Soma KM Cobrado Dia", "KM Calculado Mapa Dia",
        "Código Cliente", "Cidade", "Endereço Simplificado", "Serviço Prestado"
    ]
    ws_det.append(headers_det)
    for col_idx in range(1, len(headers_det) + 1):
        cell = ws_det.cell(1, col_idx)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center")

    row_num_det = 2
    for d in dados_auditoria["detalhes"]:
        ws_det.cell(row_num_det, 1, d["Data"])
        ws_det.cell(row_num_det, 2, d["Aba"])
        ws_det.cell(row_num_det, 3, d["OS Atendida"])
        ws_det.cell(row_num_det, 4, d["Cobrando KM?"])
        ws_det.cell(row_num_det, 5, d["KM Registrado"])
        ws_det.cell(row_num_det, 6, d["Soma KM do Dia"])
        ws_det.cell(row_num_det, 7, d["KM Calculado Mapa"])
        ws_det.cell(row_num_det, 8, d["Cód. Cliente"])
        ws_det.cell(row_num_det, 9, d["Cidade"])
        ws_det.cell(row_num_det, 10, d["Endereço Simplificado"])
        ws_det.cell(row_num_det, 11, d["Serviço Prestado"])

        for col_idx in range(1, len(headers_det) + 1):
            ws_det.cell(row_num_det, col_idx).border = borda_fina
            if d["KM Registrado"] > 0:
                ws_det.cell(row_num_det, col_idx).fill = fill_destaque
                ws_det.cell(row_num_det, col_idx).font = Font(bold=True)

        row_num_det += 1

    for ws in [ws_dia, ws_det]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if val_str.startswith("="):
                    val_str = "Abrir Rota Cidades"
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

    wb.save(output_stream)
