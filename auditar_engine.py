"""
Engine de Auditoria de KM / Deslocamento - Consigaz
Cruza a planilha de fechamento com o relatório Totvs,
higieniza endereços, agrupa todas as OSs e KMs por dia,
calcula a soma diária de KM cobrado e gera links do Google Maps.
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
        if cidade and str(cidade).strip() and str(cidade).strip().upper() != "NONE":
            uf_str = str(uf).strip().upper() if uf and str(uf).strip() and str(uf).strip().upper() != "NONE" else ""
            return f"{str(cidade).strip().title()}" + (f" - {uf_str}" if uf_str else "")
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
            numero = ''
        else:
            num = re.search(r'^\d+', resto)
            numero = num.group(0) if num else resto.split()[0] if resto else ''
        if numero in ['0', '00', 'SN', 'S/N', '']:
            s = rua
        else:
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
      - todas_oss_fechamento: lista completa de todas as linhas de OS do fechamento (com ou sem KM)
      - oss_por_data: mapa { data_str: [list of OS numbers na data naquele fechamento] }
      - abas_encontradas: resumo de abas auditadas
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

        # Identifica cabeçalhos nas primeiras 15 linhas
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


def carregar_dados_totvs(wb_totvs):
    """
    Processa a aba Totvs (ESOS0301... ou Relatório) e indexa todas as OSs
    com endereços higienizados, código de cliente, técnicos e datas/horários de finalização.
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
    """Remove endereços idênticos repetidos mantendo a ordem de visita."""
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
    """
    Divide a rota em múltiplos trechos caso ultrapasse o limite do Google Maps Web (até 8 paradas intermediárias por link)
    para permitir somar os KMs de todos os endereços do dia sem perder nenhuma parada.
    """
    paradas_validas = remover_duplicados_preservando_ordem(paradas)
    if not paradas_validas:
        return {
            "trecho1": {
                "nome": "Trecho Único",
                "url": gerar_url_google_maps(base_endereco, [], base_endereco),
                "origem": base_endereco,
                "destino": base_endereco,
                "qtd_paradas": 0
            },
            "trecho2": None,
            "todos_trechos": []
        }

    # Se couber em 1 trecho (até 8 paradas intermediárias + Origem + Retorno = 10 pontos)
    if len(paradas_validas) <= 8:
        t_unico = {
            "nome": f"Rota Completa ({len(paradas_validas)} endereços): Base ➔ Visitas ➔ Base",
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

    # Se houver mais de 8 paradas, divide em blocos de até 8 paradas conectando um trecho no outro
    todos_trechos = []
    idx_inicio = 0
    num_trecho = 1
    ponto_atual = base_endereco

    while idx_inicio < len(paradas_validas):
        restantes = len(paradas_validas) - idx_inicio
        if restantes <= 8:
            # Último trecho: passa pelas restantes e retorna para a Base
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
            # Trecho intermediário: pega 8 paradas (7 intermediárias + 8ª como destino deste trecho)
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
    """
    Gera link direto de rotas para o Google Maps (/maps/dir/Origem/Dest1/.../DestN/Destino).
    """
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


def processar_auditoria_completa(wb_fechamento, wb_totvs, base_codigo="MAU", base_endereco_custom=None, raio_franquia=40.0):
    """
    Executa o cruzamento completo e gera:
      1. Consolidação por DIA (`por_dia`): junta todas as OSs do mesmo dia, soma todos os KMs cobrados no dia,
         destaca quais OSs estão cobrando e monta a rota completa do Google Maps com todos os endereços do dia.
      2. Resumo por OS Cobrada (`resumo`).
    """
    base_endereco = base_endereco_custom or obter_endereco_base(base_codigo)

    # 1. Extrai todas as OSs do fechamento (com e sem KM)
    itens_km, _, abas_auditadas, todas_oss_fechamento = extrair_todas_oss_fechamento(wb_fechamento)

    # 2. Carrega índice Totvs
    os_dict, _ = carregar_dados_totvs(wb_totvs)

    # 3. Enriquece TODAS as OSs do fechamento com os dados do Totvs (Data, Endereço, Cidade, Serviço, Cliente)
    oss_enriquecidas = []
    mapa_oss_por_dia = {}

    for reg in todas_oss_fechamento:
        os_id = reg["os"]
        totvs_os = os_dict.get(os_id, {})

        dt_str = reg["data_str"] or totvs_os.get("data_str")
        dt_obj = reg["data_obj"] or totvs_os.get("data_obj")

        if dt_str:
            try:
                data_fmt = datetime.strptime(dt_str, "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                data_fmt = dt_str
            chave_dia = dt_str
        else:
            # Se a OS não tiver data nem no fechamento nem no Totvs, agrupa individualmente se tiver KM
            data_fmt = f"Sem Data (OS {os_id})"
            chave_dia = f"SEM_DATA_{os_id}"

        cidade_os = reg["local"] or totvs_os.get("cidade") or ""
        uf_os = totvs_os.get("uf") or ""
        end_limpo = totvs_os.get("endereco_limpo")
        if not end_limpo:
            if cidade_os:
                end_limpo = f"{cidade_os.title()}" + (f" - {uf_os}" if uf_os else "")
            else:
                end_limpo = base_endereco

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

    # 4. Constrói a visão consolidada POR DIA (apenas para os dias que têm pelo menos 1 OS cobrando KM)
    resumo_por_dia = []
    dossies_por_dia = {}

    chaves_ordenadas = sorted(mapa_oss_por_dia.keys())
    for chave_dia in chaves_ordenadas:
        lista_oss_dia = mapa_oss_por_dia[chave_dia]
        oss_com_km = [o for o in lista_oss_dia if o["km_cobrado"] > 0]
        if not oss_com_km:
            continue

        data_fmt = lista_oss_dia[0]["data_formatada"]

        # Soma de TODOS os KMs registrados nas OSs daquele dia
        soma_km_dia = sum(o["km_cobrado"] for o in oss_com_km)
        soma_pedagio_dia = sum(o["pedagio"] for o in lista_oss_dia)
        soma_km_faturavel_dia = sum(max(0.0, o["km_cobrado"] - raio_franquia) for o in oss_com_km)

        # Ordena as OSs do dia: primeiro mantemos a ordem das visitas/planilha, mas também criamos lista com destaque
        # Para a rota no Google Maps, usamos a ordem cronológica (se houver hora no Totvs) ou a ordem da planilha
        lista_oss_rota = sorted(
            lista_oss_dia,
            key=lambda x: (x["hora_str"] if x["hora_str"] else "99:99:99", x["linha"])
        )

        # Agrupa as paradas por endereço único para mostrar claramente no Google Maps:
        # Qual endereço -> Quais OSs do dia estão naquele endereço -> Quanto KM foi cobrado
        paradas_agrupadas = []
        mapa_end_idx = {}
        for o in lista_oss_rota:
            end_key = (o["endereco_limpo"] or base_endereco).strip().upper()
            if end_key not in mapa_end_idx:
                mapa_end_idx[end_key] = len(paradas_agrupadas)
                paradas_agrupadas.append({
                    "ordem": len(paradas_agrupadas) + 1,
                    "endereco": o["endereco_limpo"] or base_endereco,
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

        enderecos_rota_dia = [p["endereco"] for p in paradas_agrupadas]
        enderecos_apenas_cobradas = remover_duplicados_preservando_ordem([
            o["endereco_limpo"] or base_endereco for o in oss_com_km
        ])

        trechos_dia = dividir_trechos_maps(base_endereco, enderecos_rota_dia)
        link_maps_todas_dia = trechos_dia["trecho1"]["url"]
        link_maps_apenas_cobradas = gerar_url_google_maps(base_endereco, enderecos_apenas_cobradas, base_endereco)

        cidades_dia = []
        for o in lista_oss_dia:
            c = o["cidade"]
            if c and c != "-" and c not in cidades_dia:
                cidades_dia.append(c)

        abas_dia = []
        for o in oss_com_km:
            if o["aba"] not in abas_dia:
                abas_dia.append(o["aba"])

        # Texto claro das OSs que estão cobrando KM no dia
        detalhe_oss_cobrando = [
            f"OS {o['os']} ({o['km_cobrado']:,.1f} km - {o['cidade']})"
            for o in oss_com_km
        ]
        todas_oss_numeros_str = ", ".join(str(o["os"]) for o in lista_oss_dia)

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
            "Soma Pedágio no Dia": soma_pedagio_dia,
            "KM Faturável do Dia": max(0.0, soma_km_dia - raio_franquia),
            "Soma KM Faturável por OS": soma_km_faturavel_dia,
            "Cidades do Dia": ", ".join(cidades_dia),
            "Paradas Google Maps": paradas_agrupadas,
            "Trechos Maps": trechos_dia,
            "Link Maps (Todas do Dia)": link_maps_todas_dia,
            "Link Maps (Só OSs com KM)": link_maps_apenas_cobradas
        }
        resumo_por_dia.append(registro_dia)

        label_seletor = f"{data_fmt} — Soma do Dia: {soma_km_dia:,.1f} km ({len(oss_com_km)} OS cobrando / {len(lista_oss_dia)} OSs no dia)"
        dossies_por_dia[label_seletor] = registro_dia

    # 5. Constrói também a lista individual por OS que cobrou KM (enriquecida com a Soma do Dia)
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
            "KM Cobrado na OS": o["km_cobrado"],
            "KM Cobrado": o["km_cobrado"],
            "Soma KM do Dia": info_dia.get("Soma KM Cobrado no Dia", o["km_cobrado"]),
            "OSs Cobrando no Dia": info_dia.get("OSs Cobrando KM", f"OS {o['os']}"),
            "Franquia Raio (km)": raio_franquia,
            "KM Faturável Cobrado": km_faturavel_os,
            "Qtd OS Técnico": info_dia.get("Qtd Total OSs no Dia", 1),
            "Qtd OS Total Dia": info_dia.get("Qtd Total OSs no Dia", 1),
            "Link Maps (Técnico)": info_dia.get("Link Maps (Só OSs com KM)", ""),
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
                "Ordem": seq,
                "Aba": o["aba"],
                "OS Atendida": o["os"],
                "Cobrando KM?": f"SIM ({o['km_cobrado']:,.1f} km)" if o["km_cobrado"] > 0 else "Não (0 km)",
                "KM Registrado": o["km_cobrado"],
                "Pedágio": o["pedagio"],
                "Cód. Cliente": o["cod_cliente"],
                "Cliente": o["cliente"],
                "Cidade": o["cidade"],
                "Endereço p/ Google Maps": o["endereco_limpo"],
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
    """
    Gera planilha Excel formatada com:
      1. Resumo por Dia (Soma de KMs do Dia + todas as OSs + links do Google Maps)
      2. Relatório de Deslocamento (Todas as OSs dos dias com KM, destacando quem cobra KM e os endereços)
      3. Auditoria por OS Cobrada
    """
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

    # ABA 1: Resumo por Dia (Soma de KMs)
    ws_dia = wb.active
    ws_dia.title = "Soma de KM por Dia"
    ws_dia.views.sheetView[0].showGridLines = True

    headers_dia = [
        "Data", "Qtd OSs Cobrando KM", "Qtd Total OSs no Dia",
        "OSs Cobrando KM no Dia (Detalhe)", "SOMA KM COBRADO NO DIA",
        "Cidades Visitadas no Dia", "Todas as OSs do Dia",
        "Link Google Maps (Todas as OSs do Dia)", "Link Google Maps (Só OSs com KM)"
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
        ws_dia.cell(row_num, 2, d["Qtd OSs Cobrando"])
        ws_dia.cell(row_num, 3, d["Qtd Total OSs no Dia"])
        ws_dia.cell(row_num, 4, d["OSs Cobrando KM"])
        ws_dia.cell(row_num, 5, d["Soma KM Cobrado no Dia"])
        ws_dia.cell(row_num, 6, d["Cidades do Dia"])
        ws_dia.cell(row_num, 7, d["Todas OSs Numeros"])

        c_link_todas = ws_dia.cell(row_num, 8)
        if d["Link Maps (Todas do Dia)"]:
            c_link_todas.value = f'=HYPERLINK("{d["Link Maps (Todas do Dia)"]}", "🗺️ Abrir Rota Completa do Dia")'
            c_link_todas.font = Font(color="0000FF", underline="single")
        else:
            c_link_todas.value = "-"

        c_link_km = ws_dia.cell(row_num, 9)
        if d["Link Maps (Só OSs com KM)"]:
            c_link_km.value = f'=HYPERLINK("{d["Link Maps (Só OSs com KM)"]}", "🗺️ Rota Só OSs com KM")'
            c_link_km.font = Font(color="0000FF", underline="single")
        else:
            c_link_km.value = "-"

        for col_idx in range(1, len(headers_dia) + 1):
            ws_dia.cell(row_num, col_idx).border = borda_fina
            if d["Qtd OSs Cobrando"] > 1 or d["Soma KM Cobrado no Dia"] >= 100:
                ws_dia.cell(row_num, col_idx).fill = fill_destaque
            if col_idx in [1, 2, 3, 5, 8, 9]:
                ws_dia.cell(row_num, col_idx).alignment = Alignment(horizontal="center", vertical="center")

        row_num += 1

    # ABA 2: Relatório de Deslocamento (Todas as OSs do Dia + Endereços)
    ws_det = wb.create_sheet(title="Relatorio Deslocamento Dia")
    ws_det.views.sheetView[0].showGridLines = True

    headers_det = [
        "Data", "Aba Planilha", "Ordem de Serviço", "Cobrando KM?",
        "KM Registrado na OS", "Soma KM do Dia", "Código Cliente",
        "Cidade", "Endereço Completo (Google Maps)", "Pedágio", "Serviço Prestado"
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
        ws_det.cell(row_num_det, 7, d["Cód. Cliente"])
        ws_det.cell(row_num_det, 8, d["Cidade"])
        ws_det.cell(row_num_det, 9, d["Endereço p/ Google Maps"])
        ws_det.cell(row_num_det, 10, d["Pedágio"])
        ws_det.cell(row_num_det, 11, d["Serviço Prestado"])

        for col_idx in range(1, len(headers_det) + 1):
            ws_det.cell(row_num_det, col_idx).border = borda_fina
            if d["KM Registrado"] > 0:
                ws_det.cell(row_num_det, col_idx).fill = fill_destaque
                ws_det.cell(row_num_det, col_idx).font = Font(bold=True)

        row_num_det += 1

    # Auto-ajuste de colunas
    for ws in [ws_dia, ws_det]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if val_str.startswith("="):
                    val_str = "Abrir Rota Completa do Dia"
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 48)

    wb.save(output_stream)
