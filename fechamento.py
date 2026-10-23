import os
import re
import sys
import tempfile
import time
from datetime import datetime
import openpyxl
import win32com.client

# Garante compatibilidade de caracteres especiais no terminal do Windows
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass


# =====================================================
# CONFIGURAÇÕES
# =====================================================

PLANILHA_DESTINO = r"C:\Users\vchiku\Desktop\Controle_Fechamento_TESTE.xlsx"
# Quando for para o OneDrive, informe o caminho local sincronizado:
# PLANILHA_DESTINO = r"C:\Users\vchiku\OneDrive - Empresa\Controle_Fechamento.xlsx"

NOME_PASTA_ORIGEM = "_Teste_Fechamento"
NOME_PASTA_PROCESSADOS = "_Processados"

# Arquivo de log na mesma pasta deste script
CAMINHO_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "log_execucao.txt")

# Categorias visuais no Outlook
CATEGORIA_ERRO = "Fechamento - Erro"            # Cor Vermelha
CATEGORIA_DUPLICADO = "Fechamento - Duplicado"  # Cor Amarela
CATEGORIA_SUCESSO = "Fechamento - OK"           # Cor Verde

BASES = [
    "PTB", "UBR", "SNC", "MAU", "SVC", "CAR", "SJC", "JDS",
    "DQX", "CAN", "BRU", "ARA", "PLN", "RIB", "CORP", "BTM", "BRA"
]

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

MAPA_BASES = {
    "PTB": "PTB", "PORTO BELO": "PTB",
    "UBR": "UBR", "UBERLANDIA": "UBR", "UBERLÂNDIA": "UBR",
    "SNC": "SNC", "SENADOR CANEDO": "SNC",
    "MAU": "MAU", "MAUA": "MAU", "MAUÁ": "MAU",
    "SVC": "SVC", "SAO VICENTE": "SVC", "SÃO VICENTE": "SVC",
    "CAR": "CAR", "CARIACICA": "CAR",
    "SJC": "SJC",
    "JDS": "JDS",
    "DQX": "DQX", "DUQUE DE CAXIAS": "DQX",
    "CAN": "CAN", "CANOAS": "CAN",
    "BRU": "BRU",
    "ARA": "ARA", "ARAUCARIA": "ARA", "ARAUCÁRIA": "ARA",
    "PLN": "PLN", "PAULINIA": "PLN", "PAULÍNIA": "PLN",
    "RIB": "RIB", "RBP": "RIB", "RIBEIRAO PRETO": "RIB", "RIBEIRÃO PRETO": "RIB",
    "CORP": "CORP", "CORPORATIVO": "CORP",
    "BTM": "BTM", "BETIM": "BTM",
    "BRA": "BRA", "BARUERI": "BRU", "BRASILIA": "BRA", "BRASÍLIA": "BRA", "BSB": "BRA"
}


def limpar_assunto(texto):
    """Remove prefixos comuns como ENC:, RE:, RES:, FWD:, ANÁLISE: para manter o assunto limpo e pesquisável no Outlook."""
    if not texto:
        return ""
    padrao = r'^(?:(?:ENC|RE|RES|FWD|FW|AN[AÁ]LISE)\s*:\s*)+'
    return re.sub(padrao, '', str(texto).strip(), flags=re.IGNORECASE).strip()


# =====================================================
# IDENTIFICA BASE E MES
# =====================================================

def extrair_partes(texto):
    """Divide o texto em palavras isoladas, ignorando separadores e pontuações comuns."""
    if not texto:
        return set()
    return set(re.split(r'[-_/\\\s\[\]().,;:!ºª]+', str(texto).upper()))


def encontrar_base(texto):
    for parte in extrair_partes(texto):
        if parte in MAPA_BASES:
            return MAPA_BASES[parte]
    return None


def encontrar_mes(texto):
    for parte in extrair_partes(texto):
        if parte in MAPA_MESES:
            return MAPA_MESES[parte]
    return None


# =====================================================
# IDENTIFICA ANEXO CORRETO E DADOS
# =====================================================

def tem_aba_consolidado(anexo):
    """Verifica rapidamente se o anexo Excel possui a aba 'CONSOLIDADO'."""
    temp_path = os.path.join(tempfile.gettempdir(), anexo.FileName)
    try:
        anexo.SaveAsFile(temp_path)
        wb = openpyxl.load_workbook(temp_path, read_only=True)
        tem = "CONSOLIDADO" in wb.sheetnames
        wb.close()
        return tem
    except Exception:
        return False
    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


def localizar_anexo_e_dados(mail):
    """
    Identifica o anexo Excel correto e extrai a Base e o Mês.
    Procura tanto no assunto quanto no nome do arquivo anexo.
    """
    assunto = str(getattr(mail, "Subject", "") or "")
    base_assunto = encontrar_base(assunto)
    mes_assunto = encontrar_mes(assunto)

    anexos_validos = []
    for i in range(1, mail.Attachments.Count + 1):
        try:
            anexo = mail.Attachments.Item(i)
            nome = anexo.FileName
            nome_upper = nome.upper()
            if nome_upper.endswith(".XLSX") or nome_upper.endswith(".XLSM"):
                anexos_validos.append(anexo)
        except Exception:
            continue

    if not anexos_validos:
        return None, base_assunto, mes_assunto, "Nenhum anexo Excel (.xlsx / .xlsm) encontrado no e-mail."

    # Se houver mais de uma planilha Excel, prioriza apenas as que possuem a aba 'CONSOLIDADO'
    if len(anexos_validos) > 1:
        com_consolidado = [a for a in anexos_validos if tem_aba_consolidado(a)]
        if com_consolidado:
            anexos_validos = com_consolidado

    # 1. Tenta encontrar um anexo que case com a Base e Mês
    for anexo in anexos_validos:
        nome_anexo = anexo.FileName
        b = base_assunto or encontrar_base(nome_anexo)
        m = mes_assunto or encontrar_mes(nome_anexo)

        if b and m:
            partes_anexo = extrair_partes(nome_anexo)
            b_ok = any(MAPA_BASES.get(p) == b for p in partes_anexo)
            m_ok = any(MAPA_MESES.get(p) == m for p in partes_anexo)
            if b_ok and m_ok:
                return anexo, b, m, None

    # 2. Se houver apenas 1 anexo Excel e conseguirmos base e mês
    if len(anexos_validos) == 1:
        anexo = anexos_validos[0]
        b = base_assunto or encontrar_base(anexo.FileName)
        m = mes_assunto or encontrar_mes(anexo.FileName)
        if b and m:
            return anexo, b, m, None

    # 3. Se faltou Base ou Mês, monta mensagem de erro explicativa
    b_final = base_assunto or (encontrar_base(anexos_validos[0].FileName) if anexos_validos else None)
    m_final = mes_assunto or (encontrar_mes(anexos_validos[0].FileName) if anexos_validos else None)

    motivos = []
    if not b_final:
        motivos.append("Base não identificada (nem no assunto, nem no anexo)")
    if not m_final:
        motivos.append("Mês não identificado (nem no assunto, nem no anexo)")
    if b_final and m_final:
        motivos.append(f"Nenhum anexo corresponde à Base '{b_final}' e ao Mês '{m_final}'")

    return None, b_final, m_final, " | ".join(motivos)


# =====================================================
# LEITURA DO ANEXO EXCEL
# =====================================================

def ler_excel(caminho):
    wb = openpyxl.load_workbook(caminho, data_only=True)

    if "CONSOLIDADO" not in wb.sheetnames:
        wb.close()
        raise Exception("Aba 'CONSOLIDADO' não encontrada no anexo.")

    ws = wb["CONSOLIDADO"]
    fornecedor = ws["AQ5"].value
    valor = ws["BJ16"].value
    wb.close()

    return fornecedor, valor


# =====================================================
# GERENCIAMENTO E CONTROLE DE DUPLICIDADE
# =====================================================

def normalizar_valor(v):
    if v is None:
        return None
    try:
        return round(float(str(v).replace(",", ".")), 2)
    except (ValueError, TypeError):
        return str(v).strip()


def carregar_planilha_e_duplicidades(caminho):
    """
    Carrega o arquivo destino uma única vez e indexa:
    - IDs do Outlook (Coluna 9 / I)
    - Chaves de dados da linha: (Base, Fornecedor, Valor normalizado)
    - Arquivos já gravados: (Base, Nome do arquivo)
    """
    if not os.path.exists(caminho):
        raise FileNotFoundError(f"Planilha de destino não encontrada: {caminho}")

    wb = openpyxl.load_workbook(caminho)
    ws = wb.active

    ids_existentes = set()
    chaves_negocio = set()
    arquivos_existentes = set()

    for linha in range(2, ws.max_row + 1):
        b = ws.cell(linha, 1).value
        f = ws.cell(linha, 3).value
        v = ws.cell(linha, 5).value
        a = ws.cell(linha, 8).value
        eid = ws.cell(linha, 9).value

        if eid:
            ids_existentes.add(str(eid).strip())

        b_str = str(b).strip().upper() if b else None
        f_str = str(f).strip() if f else None
        v_norm = normalizar_valor(v)

        if b_str and f_str and v_norm is not None:
            chaves_negocio.add((b_str, f_str, v_norm))

        if b_str and a and v_norm is not None:
            arquivos_existentes.add((b_str, str(a).strip().upper(), v_norm))

    return wb, ws, ids_existentes, chaves_negocio, arquivos_existentes


def checar_duplicidade(entry_id, base, fornecedor, valor, nome_arquivo, ids, chaves, arquivos):
    """Verifica se a linha ou o e-mail já existem na planilha."""
    if entry_id and entry_id in ids:
        return True, "E-mail com mesmo EntryID já existe na planilha (Coluna I)"

    b_str = str(base).strip().upper() if base else None
    f_str = str(fornecedor).strip() if fornecedor else None
    v_norm = normalizar_valor(valor)

    if b_str and f_str and v_norm is not None:
        if (b_str, f_str, v_norm) in chaves:
            return True, f"Linha idêntica já cadastrada (Base: {b_str}, Fornecedor: {f_str}, Valor: {v_norm})"

    if b_str and nome_arquivo and v_norm is not None:
        arq_upper = str(nome_arquivo).strip().upper()
        if (b_str, arq_upper, v_norm) in arquivos:
            return True, f"Arquivo '{nome_arquivo}' com mesmo valor ({v_norm}) já registrado para a Base {b_str}"

    return False, None


def adicionar_linha(ws, base, assunto, fornecedor, valor, data_recebimento, nome_arquivo, entry_id):
    """Adiciona nova linha à planilha na memória."""
    linha = ws.max_row + 1

    ws.cell(linha, 1).value = base
    ws.cell(linha, 2).value = limpar_assunto(assunto)
    ws.cell(linha, 3).value = fornecedor

    # Coluna D vazia
    ws.cell(linha, 5).value = valor

    # Coluna F vazia
    ws.cell(linha, 7).value = data_recebimento.strftime("%d/%m/%Y") if hasattr(data_recebimento, "strftime") else str(data_recebimento)
    ws.cell(linha, 8).value = nome_arquivo
    ws.cell(linha, 9).value = entry_id


# =====================================================
# INTEGRAÇÃO COM OUTLOOK (PASTAS E CATEGORIAS)
# =====================================================

def buscar_pasta_recursiva(pasta_atual, nome_alvo):
    if pasta_atual.Name.upper() == nome_alvo.upper():
        return pasta_atual
    try:
        for subpasta in pasta_atual.Folders:
            res = buscar_pasta_recursiva(subpasta, nome_alvo)
            if res:
                return res
    except Exception:
        pass
    return None


def localizar_pasta_origem(outlook):
    for caixa in outlook.Folders:
        try:
            pasta = buscar_pasta_recursiva(caixa, NOME_PASTA_ORIGEM)
            if pasta:
                return pasta
        except Exception:
            continue
    return None


def obter_ou_criar_pasta_processados(pasta_origem):
    """Garante que a subpasta _Processados existe dentro da pasta de fechamento."""
    try:
        for sub in pasta_origem.Folders:
            if sub.Name.upper() == NOME_PASTA_PROCESSADOS.upper():
                return sub
        return pasta_origem.Folders.Add(NOME_PASTA_PROCESSADOS)
    except Exception as e:
        print(f"Aviso ao acessar pasta '{NOME_PASTA_PROCESSADOS}': {e}")
        return None


def configurar_categorias_outlook(session):
    """Cria as categorias personalizadas com as cores no Outlook do usuário."""
    try:
        cats = session.Categories
        nomes = {cats.Item(i).Name for i in range(1, cats.Count + 1)}
        cores = {
            CATEGORIA_ERRO: 1,      # Vermelho (olCategoryColorRed)
            CATEGORIA_DUPLICADO: 4, # Amarelo (olCategoryColorYellow)
            CATEGORIA_SUCESSO: 5    # Verde (olCategoryColorGreen)
        }
        for nome, cor in cores.items():
            if nome not in nomes:
                try:
                    cats.Add(nome, cor)
                except Exception:
                    pass
    except Exception:
        pass


def aplicar_categoria(mail, categoria):
    """Aplica a etiqueta colorida diretamente no e-mail no Outlook."""
    try:
        mail.Categories = categoria
        mail.Save()
    except Exception:
        pass


# =====================================================
# GRAVAÇÃO DO ARQUIVO DE LOG
# =====================================================

def salvar_log_execucao(total, sucessos, duplicados, erros):
    """Salva o resumo de cada execução em log_execucao.txt com data e hora."""
    try:
        agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        linhas = [
            "=" * 80,
            f"EXECUÇÃO: {agora}",
            f"Total de e-mails analisados: {total}",
            f"Novos registros gravados:    {len(sucessos)} (Movidos para '{NOME_PASTA_PROCESSADOS}')",
            f"Duplicados:                  {len(duplicados)} (Mantidos na pasta com etiqueta Amarela)",
            f"Erros/Pendentes:             {len(erros)} (Mantidos na pasta com etiqueta Vermelha)",
            "-" * 80
        ]

        if duplicados:
            linhas.append("[!] DUPLICADOS:")
            for i, dup in enumerate(duplicados, 1):
                linhas.append(f"  {i}. {dup['assunto']}")
                linhas.append(f"     Motivo: {dup['motivo']}")
            linhas.append("")

        if erros:
            linhas.append("[X] ERROS / PENDÊNCIAS:")
            for i, err in enumerate(erros, 1):
                linhas.append(f"  {i}. {err['assunto']}")
                linhas.append(f"     Motivo: {err['motivo']}")
        if sucessos:
            linhas.append("[OK] REGISTROS GRAVADOS COM SUCESSO:")
            for i, suc in enumerate(sucessos, 1):
                linhas.append(f"  {i}. {suc['assunto']}")
                linhas.append(f"     Base: {suc['base']} | Mês: {suc['mes']} | Fornecedor: {suc['fornecedor']} | Valor: {suc['valor']}")
            linhas.append("")

        if total == 0:
            linhas.append("Status: Nenhum e-mail encontrado na pasta para processar.")
        elif not duplicados and not erros:
            linhas.append("Status: Todos os e-mails foram processados com 100% de sucesso!")

        linhas.append("=" * 80 + "\n\n")

        with open(CAMINHO_LOG, "a", encoding="utf-8") as f:
            f.write("\n".join(linhas))
        print(f"Histórico detalhado salvo em: {CAMINHO_LOG}")
    except Exception as e:
        print(f"Aviso ao salvar arquivo de log: {e}")


def atualizar_outlook(session, tempo_espera=6):
    """Força sincronização/atualização (Enviar/Receber) no Outlook."""
    print("Atualizando Outlook (Enviar/Receber)...")
    try:
        session.SendAndReceive(False)
    except Exception as e:
        print(f"Aviso ao sincronizar (SendAndReceive): {e}")

    try:
        for i in range(1, session.SyncObjects.Count + 1):
            session.SyncObjects.Item(i).Start()
    except Exception as e:
        print(f"Aviso ao sincronizar (SyncObjects): {e}")

    # Aguarda brevemente para o Outlook sincronizar itens recebidos/movidos
    for i in range(tempo_espera, 0, -1):
        print(f"Sincronizando pastas com o servidor... {i}s", end="\r")
        time.sleep(1)
    print("Outlook atualizado com sucesso!          \n")


# =====================================================
# PROCESSAMENTO PRINCIPAL
# =====================================================

def main():
    print("=" * 80)
    print("INICIANDO AUTOMAÇÃO DE FECHAMENTO")
    print("=" * 80)

    # 1. Conecta no Outlook, atualiza (refresh) e configura categorias
    session = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    atualizar_outlook(session, tempo_espera=10)
    configurar_categorias_outlook(session)

    pasta = localizar_pasta_origem(session)
    if pasta is None:
        raise Exception(f"Pasta '{NOME_PASTA_ORIGEM}' não encontrada no Outlook.")

    pasta_processados = obter_ou_criar_pasta_processados(pasta)

    # 2. Carrega a planilha de destino uma única vez
    print(f"Carregando planilha destino: {PLANILHA_DESTINO}")
    wb_destino, ws_destino, ids_existentes, chaves_negocio, arquivos_existentes = carregar_planilha_e_duplicidades(PLANILHA_DESTINO)

    # 3. Pega uma foto instantânea dos e-mails
    itens_email = []
    try:
        for item in pasta.Items:
            if getattr(item, "Class", None) == 43:
                # Garante que o item realmente pertence à pasta atual (ignora itens já movidos)
                try:
                    if item.Parent.EntryID != pasta.EntryID:
                        continue
                except Exception:
                    pass
                itens_email.append(item)
    except Exception as e:
        print(f"Erro ao ler itens da pasta: {e}")

    print(f"Total de e-mails para analisar na pasta: {len(itens_email)}\n")

    sucessos = []
    duplicados = []
    erros = []

    for mail in itens_email:
        assunto = str(getattr(mail, "Subject", "Sem Assunto") or "")
        entry_id = str(getattr(mail, "EntryID", "") or "").strip()
        caminho_temp = None

        try:
            # Checagem preliminar rápida de EntryID
            if entry_id and entry_id in ids_existentes:
                motivo = "EntryID do e-mail já existe na coluna I da planilha"
                aplicar_categoria(mail, CATEGORIA_DUPLICADO)
                duplicados.append({"assunto": assunto, "motivo": motivo})
                print(f"[DUPLICADO] {assunto}")
                continue

            # Localiza anexo correto e obtém base/mês
            anexo, base, mes, motivo_anexo = localizar_anexo_e_dados(mail)

            if not anexo:
                aplicar_categoria(mail, CATEGORIA_ERRO)
                erros.append({"assunto": assunto, "motivo": motivo_anexo})
                print(f"[ERRO] {assunto} -> {motivo_anexo}")
                continue

            # Salva anexo temporário para leitura
            caminho_temp = os.path.join(tempfile.gettempdir(), anexo.FileName)
            anexo.SaveAsFile(caminho_temp)

            # Extrai valores da aba CONSOLIDADO
            fornecedor, valor = ler_excel(caminho_temp)

            # Valida duplicidade da linha (Base, Fornecedor, Valor ou Nome do Arquivo)
            eh_duplicado, motivo_dup = checar_duplicidade(
                entry_id=entry_id,
                base=base,
                fornecedor=fornecedor,
                valor=valor,
                nome_arquivo=anexo.FileName,
                ids=ids_existentes,
                chaves=chaves_negocio,
                arquivos=arquivos_existentes
            )

            if eh_duplicado:
                aplicar_categoria(mail, CATEGORIA_DUPLICADO)
                duplicados.append({"assunto": assunto, "motivo": motivo_dup})
                print(f"[DUPLICADO] {assunto} -> {motivo_dup}")
                continue

            # Adiciona na memória do Excel
            adicionar_linha(
                ws=ws_destino,
                base=base,
                assunto=assunto,
                fornecedor=fornecedor,
                valor=valor,
                data_recebimento=mail.ReceivedTime,
                nome_arquivo=anexo.FileName,
                entry_id=entry_id
            )

            # Atualiza índices em memória para evitar duplicatas dentro do próprio lote
            if entry_id:
                ids_existentes.add(entry_id)
            if base and fornecedor and valor is not None:
                v_norm = normalizar_valor(valor)
                chaves_negocio.add((str(base).strip().upper(), str(fornecedor).strip(), v_norm))
            if base and anexo.FileName and valor is not None:
                arquivos_existentes.add((str(base).strip().upper(), str(anexo.FileName).strip().upper(), normalizar_valor(valor)))

            assunto_limpo = limpar_assunto(assunto)
            sucessos.append({
                "mail": mail,
                "assunto": assunto_limpo,
                "base": base,
                "mes": mes,
                "fornecedor": fornecedor,
                "valor": valor,
                "arquivo": anexo.FileName
            })

            print(f"[OK] {base} | {mes} | Forn: {fornecedor} | Val: {valor} | {assunto_limpo}")

        except Exception as e:
            aplicar_categoria(mail, CATEGORIA_ERRO)
            erros.append({"assunto": assunto, "motivo": str(e)})
            print(f"[ERRO] {assunto} -> {e}")

        finally:
            if caminho_temp and os.path.exists(caminho_temp):
                try:
                    os.remove(caminho_temp)
                except Exception:
                    pass

    # =====================================================
    # SALVAMENTO SEGURO DA PLANILHA E MOVIMENTAÇÃO DE E-MAILS
    # =====================================================
    salvou_com_sucesso = False

    if sucessos:
        print("\nSalvando novos registros na planilha de destino...")
        try:
            wb_destino.save(PLANILHA_DESTINO)
            salvou_com_sucesso = True
            print("Planilha salva com sucesso!")
        except Exception as e:
            print("\n" + "!" * 80)
            print("ATENÇÃO: FALHA AO SALVAR A PLANILHA!")
            print(f"Erro: {e}")
            print("Verifique se o arquivo está aberto no Excel ou em sincronização no OneDrive.")
            print("NENHUM e-mail foi movido para a pasta de processados para evitar perda de dados.")
            print("!" * 80 + "\n")
    else:
        print("\nNenhum registro novo para salvar.")

    wb_destino.close()

    # Aplica categoria verde e move para _Processados SOMENTE após confirmação de salvamento
    if salvou_com_sucesso and pasta_processados:
        print(f"Movendo {len(sucessos)} e-mail(s) para a pasta '{NOME_PASTA_PROCESSADOS}'...")
        for item in sucessos:
            try:
                aplicar_categoria(item["mail"], CATEGORIA_SUCESSO)
                item["mail"].Move(pasta_processados)
            except Exception as e:
                print(f"Não foi possível mover o e-mail '{item['assunto']}': {e}")

    # =====================================================
    # PAINEL DE RELATÓRIO E AVISOS FINAL
    # =====================================================
    print("\n" + "=" * 80)
    print("                     RELATÓRIO FINAL DE EXECUÇÃO")
    print("=" * 80)
    print(f"Total analisados:  {len(itens_email)}")
    print(f"Novos gravados:    {len(sucessos)} (Movidos para '{NOME_PASTA_PROCESSADOS}')")
    print(f"Duplicados:        {len(duplicados)} (Mantidos na pasta com etiqueta Amarela)")
    print(f"Erros/Pendentes:   {len(erros)} (Mantidos na pasta com etiqueta Vermelha)")
    print("=" * 80)

    if duplicados:
        print("\n" + "[!] " * 20)
        print("AVISO: E-MAILS DUPLICADOS IDENTIFICADOS (Marcados com Amarelo no Outlook):")
        print("[!] " * 20)
        for i, dup in enumerate(duplicados, 1):
            print(f"  {i}. {dup['assunto']}")
            print(f"     Motivo: {dup['motivo']}\n")

    if erros:
        print("\n" + "[X] " * 20)
        print("AVISO: E-MAILS COM ERRO OU PENDÊNCIA (Marcados com Vermelho no Outlook):")
        print("[X] " * 20)
        for i, err in enumerate(erros, 1):
            print(f"  {i}. {err['assunto']}")
            print(f"     Motivo: {err['motivo']}\n")

    if len(itens_email) == 0:
        print("\nNenhum e-mail encontrado na pasta '_Teste_Fechamento' para processar.")
        print("Dica: Arraste ou copie os e-mails de fechamento para a pasta '_Teste_Fechamento' antes de executar.")
    elif not duplicados and not erros:
        print("\nTodos os e-mails foram processados com 100% de sucesso!")

    print("=" * 80)

    # 4. Salva histórico permanente no arquivo log_execucao.txt
    salvar_log_execucao(len(itens_email), sucessos, duplicados, erros)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[ERRO CRÍTICO]: {e}")
    finally:
        input("\nPressione ENTER para fechar...")