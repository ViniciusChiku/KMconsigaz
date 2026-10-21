"""
Engine de Auditoria de KM / Deslocamento - Consigaz
Cruza a planilha de fechamento com o relatório Totvs,
higieniza endereços, monta rotas do dia e gera links oficiais do Google Maps.
"""

import io
import os
import re
import urllib.parse
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
    8: {"sigla": "MAU", "cidade": "Mauá/SP", "endereco": "RUA VITORIA PERDAO PETIGROSSO, 115 - MAUA/SP"},
    2: {"sigla": "PLN", "cidade": "Paulínia/SP", "endereco": "RUA EITOR VALENTINO DE SOUZA, 337 - PAULINA/SP"},
    1: {"sigla": "BRU", "cidade": "Barueri/SP", "endereco": "RUA JOSE PEREIRA SOBRINHO, 485 - BARUERI/SP"},
    6: {"sigla": "CAN", "cidade": "Canoas/RS", "endereco": "RUA ANTONIO FREDERICO OZANAN, 1901 - CANOAS/RS"},
    9: {"sigla": "DQX", "cidade": "Duque de Caxias/RJ", "endereco": "RUA FREI CANECA, 407 - DUQUE DE CAXIAS/RJ"},
    11: {"sigla": "RIB", "cidade": "Ribeirão Preto/SP", "endereco": "ROD ALEXANDRE BALBO, SN - RIBEIRÃO PRETO/SP"},
    5: {"sigla": "CAR", "cidade": "Cariacica/ES", "endereco": "ROD GOVERNADOR MARIO COVAS 1941, 7161 - CARIACICA/ES"},
    4: {"sigla": "SVC", "cidade": "São Vicente/SP", "endereco": "RUA MANOEL DE ABREU, 790 - SÃO VICENTE/SP"},
    3: {"sigla": "SVC", "cidade": "São Vicente/SP (QD 108)", "endereco": "AV. MANOEL DE EBREU, 790 QD 108 – SÃO VICENTE/SP"},
    7: {"sigla": "SNC", "cidade": "Senador Canedo/GO", "endereco": "AV PROGRESSO, SN - SENADOR CANEDO/GO"},
    10: {"sigla": "ARA", "cidade": "Araucária/PR", "endereco": "R DOUTOR JOSE CZAKI, 320 - ARAUCARA/PR"},
    13: {"sigla": "UBR", "cidade": "Uberlândia/MG", "endereco": "R WANDELVART FARIA MARQUEZ, 500 - UBERLANDIA/MG"},
    14: {"sigla": "BRA", "cidade": "Brasília/DF", "endereco": "ST SIN (SETOR DE INFLAMÁVEIS), SN - BRASILIA/DF"},
    15: {"sigla": "BTM", "cidade": "Betim/MG", "endereco": "ROD BR-381 FERNAO DIAS, S/N, SALA A KM 485 - BETIM/MG"},
    17: {"sigla": "PTB", "cidade": "Porto Belo/SC", "endereco": "ROD BR 101, S/N, SERTÃO DE SANTA LUZIA - PORTO BELO/SC"},
    18: {"sigla": "JDS", "cidade": "Jandaia do Sul/PR", "endereco": "RUA JOAO ORTEGA SANCHES, 65 - JANDAIA DO SUL/PR"}
}

# Mapa de endereços indexado por Sigla e por Número de Base
ENDERECOS_BASES = {}
OPCOES_BASES_FORMATADAS = {}

for num, dados in ENDERECOS_BASES_DETALHADOS.items():
    label = f"Base {num} - {dados['sigla']} ({dados['cidade']})"
    ENDERECOS_BASES[dados['sigla']] = dados['endereco']
    ENDERECOS_BASES[str(num)] = dados['endereco']
    OPCOES_BASES_FORMATADAS[label] = {
        "numero": num,
        "sigla": dados['sigla'],
        "endereco": dados['endereco']
    }


def detectar_base_por_texto(texto):
    """Identifica o código da base (ex: MAU, PLN, DQX) no nome do arquivo ou texto."""
    if not texto:
        return "MAU"
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
    return "MAU"


def obter_endereco_base(codigo_base):
    cod = str(codigo_base).strip().upper()
    if cod in ENDERECOS_BASES:
        return ENDERECOS_BASES[cod]
    for num, d in ENDERECOS_BASES_DETALHADOS.items():
        if cod == d["sigla"] or cod == str(num):
            return d["endereco"]
    return ENDERECOS_BASES.get("MAU", "RUA VITORIA PERDAO PETIGROSSO, 115 - MAUA/SP")


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

    # 1. Tenta abrir direto com openpyxl
    try:
        wb = openpyxl.load_workbook(stream, data_only=True)
        return wb
    except Exception:
        pass

    # 2. Se falhar, tenta descriptografar com msoffcrypto
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


def limpar_endereco(endereco_bruto, bairro=None, cidade=None, uf=None, cep=None):
    """
    Higieniza o endereço cadastrado no Totvs para o Google Maps encontrar com exatidão.
    Remove blocos, apartamentos, fases, torres e complementos prediais.
    """
    if not endereco_bruto or str(endereco_bruto).strip() in ["", "0", "None"]:
        return ""

    s = str(endereco_bruto).strip().upper()
    substituicoes = {
        'Á': 'A', 'À': 'A', 'Â': 'A', 'Ã': 'A',
        'É': 'E', 'Ê': 'E', 'Í': 'I',
        'Ó': 'O', 'Ô': 'O', 'Õ': 'O',
        'Ú': 'U', 'Ç': 'C'
    }
    for o, d in substituicoes.items():
        s = s.replace(o, d)

    s = re.sub(r'\s+', ' ', s)

    # Corta complementos prediais
    padrao = r'[\s,]*(?:BLOCO|BL\s*\d*|APTO|APT|AP\s*\d*|TORRE|TR\s*\d*|CASA|CS\s*\d*|LT\s*\d*|LOTE|CONDOMINIO|COND\s|FASE\s*\d*|ANDAR|TERREO|SALA\s*\d*).*$'
    s = re.sub(padrao, '', s, flags=re.IGNORECASE).strip().rstrip(' ,;.-')

    # Trata formato Rua, Número
    partes = s.split(',')
    if len(partes) > 1:
        rua = partes[0].strip()
        resto = partes[1].strip()
        m_rep = re.match(r'^(\d+)\s+\1\b', resto)
        if m_rep:
            resto = resto[len(m_rep.group(1)):].strip()
        if any(sn in resto.upper() for sn in ['S/N', 'SN', 'S/Nº', 'SEM NUMERO']):
            numero = 'SN'
        else:
            num = re.search(r'^\d+', resto)
            numero = num.group(0) if num else resto.split()[0] if resto else ''
        s = f"{rua}, {numero}".strip(' ,')

    elementos = [s.title()]
    if bairro and str(bairro).strip() and str(bairro).strip().upper() != "NONE":
        elementos.append(str(bairro).strip().title())
    if cidade and str(cidade).strip() and str(cidade).strip().upper() != "NONE":
        cidade_str = str(cidade).strip().title()
        uf_str = str(uf).strip().upper() if uf and str(uf).strip() and str(uf).strip().upper() != "NONE" else "SP"
        elementos.append(f"{cidade_str} - {uf_str}")
    elif uf and str(uf).strip() and str(uf).strip().upper() != "NONE":
        elementos.append(str(uf).strip().upper())

    return ", ".join(elementos)


def extrair_todas_oss_fechamento(wb_fechamento):
    """
    Varre todas as abas da planilha de fechamento (MOTO, CARRO, ENGENHARIA, INSTALAÇÃO, etc.).
    Retorna:
      - itens_km: linhas com KM > 0 (ignorando pedágio)
      - oss_por_data: mapa { data_str: [list of OS numbers na data naquele fechamento] }
      - abas_encontradas: resumo de abas auditadas
    """
    itens_km = []
    oss_por_data = {}
    abas_encontradas = set()

    for sname in wb_fechamento.sheetnames:
        ws = wb_fechamento[sname]
        if ws.max_row is None or ws.max_row < 2:
            continue

        header_row = None
        col_os = None
        col_km = None
        col_dt = None
        col_local = None
        col_viagens = None
        col_total = None

        # Identifica cabeçalhos nas primeiras 15 linhas
        for r in range(1, min(16, ws.max_row + 1)):
            for c in range(1, min(ws.max_column + 1, 35)):
                val = str(ws.cell(r, c).value or "").strip().upper()
                if "ORDEM DE SERVI" in val or val == "OS" or val == "ORDEM":
                    col_os = c
                    header_row = r
                elif "DESLOCAMENTO" in val or "KM" in val:
                    if "PEDAGIO" not in val and "PEDÁGIO" not in val:
                        col_km = c
                elif ("DATA" in val and ("ENCERRAMENTO" in val or "CONCLUSAO" in val or "FINALIZA" in val)) or val == "DATA":
                    col_dt = c
                elif "LOCAL" in val or "CIDADE" in val or "MUNICIPIO" in val:
                    col_local = c
                elif "VIAGENS" in val or "QT. VIAGENS" in val or "VIAGEM" in val:
                    col_viagens = c
                elif val == "TOTAL" or "VALOR TOTAL" in val or "VALOR TRANSPORTE" in val:
                    col_total = c

        if col_os and col_km and header_row:
            abas_encontradas.add(sname)
            for r in range(header_row + 1, ws.max_row + 1):
                val_os = ws.cell(r, col_os).value
                if val_os is None or str(val_os).strip() == "":
                    continue

                try:
                    os_clean = int(float(str(val_os).replace(",", ".")))
                except Exception:
                    os_clean = str(val_os).strip()

                val_dt = ws.cell(r, col_dt).value if col_dt else None
                dt_str, dt_obj = normalizar_data(val_dt)

                val_km = ws.cell(r, col_km).value
                try:
                    km_num = float(str(val_km).replace(",", "."))
                except (ValueError, TypeError):
                    km_num = 0.0

                # Registra no mapa geral de OSs da data
                if dt_str:
                    if dt_str not in oss_por_data:
                        oss_por_data[dt_str] = []
                    if os_clean not in oss_por_data[dt_str]:
                        oss_por_data[dt_str].append(os_clean)

                # Se houver cobrança de KM nesta OS
                if km_num > 0:
                    viagens = ws.cell(r, col_viagens).value if col_viagens else 1
                    try:
                        viagens_num = int(float(viagens))
                    except Exception:
                        viagens_num = 1

                    local_val = ws.cell(r, col_local).value if col_local else ""
                    total_val = ws.cell(r, col_total).value if col_total else 0.0

                    itens_km.append({
                        "aba": sname,
                        "linha": r,
                        "os": os_clean,
                        "km_cobrado": km_num,
                        "data_str": dt_str,
                        "data_obj": dt_obj,
                        "viagens": viagens_num,
                        "local": str(local_val or "").strip(),
                        "total_pago": total_val
                    })

    return itens_km, oss_por_data, list(abas_encontradas)


def carregar_dados_totvs(wb_totvs):
    """
    Processa a aba Totvs (ESOS0301... ou Relatório) e indexa todas as OSs
    com endereços higienizados, técnicos e datas/horários de finalização.
    """
    ws_rel = None
    for sname in wb_totvs.sheetnames:
        if sname.upper().startswith("ESOS") or "RELAT" in sname.upper():
            ws_rel = wb_totvs[sname]
            break
    if ws_rel is None:
        ws_rel = wb_totvs.active

    # Localiza a linha correta do cabeçalho da tabela
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
        elif "SITUA" in v and "USU" not in v and "MORADOR" not in v: col_map["situacao"] = c
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
        situacao = ws_rel.cell(r, col_map.get("situacao", 7)).value if "situacao" in col_map else ""
        cliente = ws_rel.cell(r, col_map.get("cliente", 13)).value if "cliente" in col_map else ""

        dt_str, dt_obj = normalizar_data(dt_final or dt_concl)
        hora_str = ""
        if isinstance(dt_final, datetime):
            hora_str = dt_final.strftime("%H:%M:%S")
        elif dt_final and " " in str(dt_final):
            hora_str = str(dt_final).split(" ")[1][:8]

        desc_motivo = ws_rel.cell(r, col_map.get("desc_motivo", 6)).value if "desc_motivo" in col_map else ""
        motivo_num = ws_rel.cell(r, col_map.get("motivo", 5)).value if "motivo" in col_map else ""
        servico_nome = str(desc_motivo or motivo_num or "VISITA OPERACIONAL").strip()

        end_limpo = limpar_endereco(end_bruto, bairro, cidade, uf, cep)

        rec = {
            "os": os_num,
            "endereco_bruto": str(end_bruto or "").strip(),
            "endereco_limpo": end_limpo,
            "bairro": str(bairro or "").strip(),
            "cidade": str(cidade or "").strip(),
            "uf": str(uf or "").strip(),
            "cep": str(cep or "").strip(),
            "tecnico": str(tecnico or "").strip(),
            "data_str": dt_str,
            "data_obj": dt_obj,
            "hora_str": hora_str,
            "situacao": str(situacao or "").strip(),
            "cliente": str(cliente or "").strip(),
            "servico": servico_nome.upper()
        }

        os_dict[os_num] = rec

        if dt_str:
            if dt_str not in por_data_dict:
                por_data_dict[dt_str] = []
            por_data_dict[dt_str].append(rec)

    return os_dict, por_data_dict


def dividir_trechos_maps(base_endereco, paradas):
    """
    Divide a rota em Trecho 1 e Trecho 2 para caber no limite do Google Maps
    e permitir somar os KMs (ex: 265 km no Trecho 1 + 145 km no Trecho 2 = 410 km).
    """
    paradas_validas = [p for p in paradas if p and str(p).strip()]
    if not paradas_validas:
        return {
            "trecho1": {
                "nome": "Trecho Único",
                "url": gerar_url_google_maps(base_endereco, [], base_endereco),
                "origem": base_endereco,
                "destino": base_endereco,
                "qtd_paradas": 0
            },
            "trecho2": None
        }

    # Se couber em 1 trecho (até 8 paradas intermediárias + Origem + Retorno)
    if len(paradas_validas) <= 8:
        return {
            "trecho1": {
                "nome": "Trecho Único (Circuito Completo)",
                "url": gerar_url_google_maps(base_endereco, paradas_validas, base_endereco),
                "origem": base_endereco,
                "destino": base_endereco,
                "qtd_paradas": len(paradas_validas)
            },
            "trecho2": None
        }

    # Se houver mais de 8 paradas, divide em 2 trechos
    # Trecho 1: Base -> paradas 0..corte-1 -> parada corte
    # Trecho 2: parada corte -> paradas corte+1..fim -> Base
    corte = min(8, len(paradas_validas) - 1)
    ponto_transicao = paradas_validas[corte]

    intermed_t1 = paradas_validas[:corte]
    url_t1 = gerar_url_google_maps(base_endereco, intermed_t1, ponto_transicao)

    intermed_t2 = paradas_validas[corte + 1:]
    url_t2 = gerar_url_google_maps(ponto_transicao, intermed_t2, base_endereco)

    ponto_t_resumo = ponto_transicao.split(",")[0] if "," in ponto_transicao else ponto_transicao

    return {
        "trecho1": {
            "nome": f"Trecho 1 (Ida e Paradas 1 a {corte + 1}): Base -> {ponto_t_resumo}",
            "url": url_t1,
            "origem": base_endereco,
            "destino": ponto_transicao,
            "qtd_paradas": len(intermed_t1) + 1
        },
        "trecho2": {
            "nome": f"Trecho 2 (Paradas {corte + 2} a {len(paradas_validas)} e Retorno): {ponto_t_resumo} -> Base",
            "url": url_t2,
            "origem": ponto_transicao,
            "destino": base_endereco,
            "qtd_paradas": len(intermed_t2) + 1
        }
    }


def gerar_url_google_maps(origem, paradas, destino=None):
    """
    Gera link direto de rotas para o Google Maps (/maps/dir/Origem/Dest1/.../DestN/Destino).
    Se houver mais de 9 paradas intermediárias, preserva as principais paradas no link web.
    """
    if not destino:
        destino = origem

    paradas_validas = [p for p in paradas if p and str(p).strip()]
    
    # O Google Maps Web suporta até 10 pontos na URL (Origem + 8 a 9 destinos + Retorno)
    if len(paradas_validas) > 8:
        # Pega as primeiras paradas e a última para caber perfeitamente no link web
        paradas_formatadas = paradas_validas[:8]
    else:
        paradas_formatadas = paradas_validas

    partes = [origem] + paradas_formatadas + [destino]
    enc_partes = [urllib.parse.quote_plus(str(p).strip()) for p in partes]
    return "https://www.google.com/maps/dir/" + "/".join(enc_partes)


def processar_auditoria_completa(wb_fechamento, wb_totvs, base_codigo="MAU", base_endereco_custom=None, raio_franquia=40.0):
    """
    Executa o cruzamento completo e gera os dados consolidados para o dashboard e relatório.
    """
    base_endereco = base_endereco_custom or obter_endereco_base(base_codigo)

    # 1. Extrai todas as OSs do fechamento e mapeia por data
    itens_km, fechamento_oss_por_data, abas_auditadas = extrair_todas_oss_fechamento(wb_fechamento)

    # 2. Carrega índice Totvs
    os_dict, por_data_dict = carregar_dados_totvs(wb_totvs)

    resumo_linhas = []
    detalhes_visitas = []
    dossies_por_dia = {}

    for item in itens_km:
        os_id = item["os"]
        km_cobrado = item["km_cobrado"]
        aba_origem = item["aba"]

        totvs_os = os_dict.get(os_id, {})
        dt_str = item["data_str"] or totvs_os.get("data_str")
        tecnico_os = totvs_os.get("tecnico", "Não identificado")
        situacao_os = totvs_os.get("situacao", "Atendida")

        # 1. Pega todas as OSs do fechamento que foram feitas nesta data
        oss_fechamento_dia = fechamento_oss_por_data.get(dt_str, [])
        if not oss_fechamento_dia:
            oss_fechamento_dia = [os_id]

        # 2. Resgata informações do Totvs para cada uma dessas OSs
        dados_oss_dia = []
        for o_num in oss_fechamento_dia:
            if o_num in os_dict:
                dados_oss_dia.append(os_dict[o_num])
            else:
                dados_oss_dia.append({
                    "os": o_num,
                    "endereco_limpo": item.get("local", ""),
                    "tecnico": tecnico_os,
                    "hora_str": "12:00:00",
                    "cliente": "",
                    "situacao": "Fechamento"
                })

        # 3. Cenário 1: Apenas as OSs do mesmo técnico da OS cobrada
        oss_mesmo_tecnico = [
            o for o in dados_oss_dia
            if tecnico_os and tecnico_os != "Não identificado" and o.get("tecnico") == tecnico_os
        ]
        if not oss_mesmo_tecnico:
            oss_mesmo_tecnico = dados_oss_dia

        # Ordena cronologicamente por horário de finalização
        oss_mesmo_tecnico = sorted(oss_mesmo_tecnico, key=lambda x: x.get("hora_str", "") or "99:99")
        dados_oss_dia_ord = sorted(dados_oss_dia, key=lambda x: x.get("hora_str", "") or "99:99")

        # Endereços para Google Maps (se a OS não trouxer endereço do cliente, usa o endereço da Base como uso excepcional)
        paradas_tecnico = [
            o.get("endereco_limpo") if o.get("endereco_limpo") else base_endereco
            for o in oss_mesmo_tecnico
        ]
        paradas_todas = [
            o.get("endereco_limpo") if o.get("endereco_limpo") else base_endereco
            for o in dados_oss_dia_ord
        ]

        link_maps_tecnico = gerar_url_google_maps(base_endereco, paradas_tecnico, base_endereco) if paradas_tecnico else ""
        link_maps_geral = gerar_url_google_maps(base_endereco, paradas_todas, base_endereco) if paradas_todas else ""

        km_faturavel_fechamento = max(0.0, km_cobrado - raio_franquia)

        data_formatada = ""
        if dt_str:
            try:
                data_formatada = datetime.strptime(dt_str, "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                data_formatada = dt_str

        linha_resumo = {
            "Aba": aba_origem,
            "OS Cobrada": os_id,
            "Data": data_formatada,
            "Técnico da OS": tecnico_os,
            "KM Cobrado": km_cobrado,
            "Franquia Raio (km)": raio_franquia,
            "KM Faturável Cobrado": km_faturavel_fechamento,
            "Qtd OS Técnico": len(oss_mesmo_tecnico),
            "Qtd OS Total Dia": len(dados_oss_dia_ord),
            "Link Maps (Técnico)": link_maps_tecnico,
            "Link Maps (Geral do Dia)": link_maps_geral,
            "Status Totvs": situacao_os,
            "Endereço OS": totvs_os.get("endereco_limpo", item.get("local", ""))
        }
        resumo_linhas.append(linha_resumo)

        # Monta estrutura pronta para o Dossiê Visual de Rota (estilo folha de conferência)
        trechos_maps = dividir_trechos_maps(base_endereco, paradas_todas)
        chave_dossie = f"{data_formatada} — OS {os_id} ({km_cobrado:,.1f} km)"
        dossies_por_dia[chave_dossie] = {
            "os_ancora": os_id,
            "data_formatada": data_formatada,
            "data_str": dt_str,
            "km_cobrado": km_cobrado,
            "raio_franquia": raio_franquia,
            "km_faturavel": km_faturavel_fechamento,
            "tecnico_os": tecnico_os,
            "aba": aba_origem,
            "tabela_oss": [
                {
                    "Ordem Serviço": o.get("os"),
                    "Cliente": o.get("cliente") or "-",
                    "Cidade": o.get("cidade") or "-",
                    "Data": data_formatada,
                    "Serviço": o.get("servico", "VISITA OPERACIONAL")
                }
                for o in dados_oss_dia_ord
            ],
            "trechos": trechos_maps
        }

        seq = 1
        for o in dados_oss_dia_ord:
            detalhes_visitas.append({
                "Data": data_formatada,
                "OS Âncora Fechamento": os_id,
                "Ordem": seq,
                "OS Atendida": o.get("os"),
                "Técnico": o.get("tecnico"),
                "Hora": o.get("hora_str"),
                "Cliente": o.get("cliente"),
                "Endereço Limpo": o.get("endereco_limpo"),
                "Bairro": o.get("bairro"),
                "Cidade": o.get("cidade"),
                "Situação": o.get("situacao")
            })
            seq += 1

    return {
        "base_codigo": base_codigo,
        "base_endereco": base_endereco,
        "raio_franquia": raio_franquia,
        "abas_auditadas": abas_auditadas,
        "resumo": resumo_linhas,
        "detalhes": detalhes_visitas,
        "dossies": dossies_por_dia,
        "total_oss_km": len(resumo_linhas),
        "total_km_cobrado": sum(r["KM Cobrado"] for r in resumo_linhas),
        "total_km_faturavel": sum(r["KM Faturável Cobrado"] for r in resumo_linhas)
    }


def exportar_relatorio_excel(dados_auditoria, output_stream):
    """
    Gera planilha Excel profissional e formatada com links clicáveis do Google Maps.
    """
    wb = openpyxl.Workbook()

    ws_resumo = wb.active
    ws_resumo.title = "Resumo Auditoria KM"
    ws_resumo.views.sheetView[0].showGridLines = True

    cor_azul = "1F4E79"
    cor_branco = "FFFFFF"

    fonte_cabecalho = Font(name="Calibri", size=11, bold=True, color=cor_branco)
    fill_cabecalho = PatternFill(start_color=cor_azul, end_color=cor_azul, fill_type="solid")
    borda_fina = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    headers_resumo = [
        "Aba Fechamento", "OS Cobrada", "Data", "Técnico da OS",
        "KM Cobrado", "Raio Franquia (km)", "KM Faturável Cobrado",
        "Qtd OS Técnico", "Qtd OS Total Dia",
        "Link Google Maps (Rota Técnico)", "Link Google Maps (Rota Geral Dia)",
        "Endereço OS Cobrada"
    ]

    ws_resumo.append(headers_resumo)
    for col_idx in range(1, len(headers_resumo) + 1):
        cell = ws_resumo.cell(1, col_idx)
        cell.font = fonte_cabecalho
        cell.fill = fill_cabecalho
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    row_num = 2
    for r in dados_auditoria["resumo"]:
        ws_resumo.cell(row_num, 1, r["Aba"])
        ws_resumo.cell(row_num, 2, r["OS Cobrada"])
        ws_resumo.cell(row_num, 3, r["Data"])
        ws_resumo.cell(row_num, 4, r["Técnico da OS"])
        ws_resumo.cell(row_num, 5, r["KM Cobrado"])
        ws_resumo.cell(row_num, 6, r["Franquia Raio (km)"])
        ws_resumo.cell(row_num, 7, r["KM Faturável Cobrado"])
        ws_resumo.cell(row_num, 8, r["Qtd OS Técnico"])
        ws_resumo.cell(row_num, 9, r["Qtd OS Total Dia"])

        c_link_tec = ws_resumo.cell(row_num, 10)
        if r["Link Maps (Técnico)"]:
            c_link_tec.value = f'=HYPERLINK("{r["Link Maps (Técnico)"]}", "🗺️ Ver Rota Técnico")'
            c_link_tec.font = Font(color="0000FF", underline="single")
        else:
            c_link_tec.value = "-"

        c_link_geral = ws_resumo.cell(row_num, 11)
        if r["Link Maps (Geral do Dia)"]:
            c_link_geral.value = f'=HYPERLINK("{r["Link Maps (Geral do Dia)"]}", "🗺️ Ver Rota Geral")'
            c_link_geral.font = Font(color="0000FF", underline="single")
        else:
            c_link_geral.value = "-"

        ws_resumo.cell(row_num, 12, r["Endereço OS"])

        for col_idx in range(1, len(headers_resumo) + 1):
            ws_resumo.cell(row_num, col_idx).border = borda_fina
            if col_idx in [2, 3, 5, 6, 7, 8, 9, 10, 11]:
                ws_resumo.cell(row_num, col_idx).alignment = Alignment(horizontal="center", vertical="center")

        row_num += 1

    # Aba Detalhes
    ws_det = wb.create_sheet(title="Detalhamento das Visitas")
    ws_det.views.sheetView[0].showGridLines = True

    headers_det = [
        "Data", "OS Âncora Fechamento", "Ordem Visita", "OS Atendida",
        "Técnico", "Horário Finalização", "Cliente",
        "Endereço Limpo", "Bairro", "Cidade", "Situação"
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
        ws_det.cell(row_num_det, 2, d["OS Âncora Fechamento"])
        ws_det.cell(row_num_det, 3, d["Ordem"])
        ws_det.cell(row_num_det, 4, d["OS Atendida"])
        ws_det.cell(row_num_det, 5, d["Técnico"])
        ws_det.cell(row_num_det, 6, d["Hora"])
        ws_det.cell(row_num_det, 7, d["Cliente"])
        ws_det.cell(row_num_det, 8, d["Endereço Limpo"])
        ws_det.cell(row_num_det, 9, d["Bairro"])
        ws_det.cell(row_num_det, 10, d["Cidade"])
        ws_det.cell(row_num_det, 11, d["Situação"])

        for col_idx in range(1, len(headers_det) + 1):
            ws_det.cell(row_num_det, col_idx).border = borda_fina

        row_num_det += 1

    # Auto-ajuste de colunas
    for ws in [ws_resumo, ws_det]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if val_str.startswith("="):
                    val_str = "Ver Rota Maps"
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    wb.save(output_stream)
