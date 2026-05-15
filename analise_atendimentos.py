import requests
import json
import time
import re
import os
import sys
import zipfile
import pyperclip
from collections import Counter
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright
from xml.sax.saxutils import escape
from xml.etree import ElementTree as ET

# ─────────────────────────────────────────────
#  CONFIGURAÇÕES
# ─────────────────────────────────────────────
APP_NAME = "AnaliseAtendimentos"
TEMPLATE_EXCEL_FILENAME = "Relatorio_Atendimentos_Mai2026.xlsx"

def obter_pasta_app() -> Path:
    """Pasta do .py em desenvolvimento ou pasta do .exe quando empacotado."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def obter_pasta_config() -> Path:
    """Pasta privada do usuario para arquivos sensiveis e estado de sessao."""
    base = os.getenv("APPDATA")
    if base:
        return Path(base) / APP_NAME
    return Path.home() / f".{APP_NAME.lower()}"


APP_DIR = obter_pasta_app()
CONFIG_DIR = obter_pasta_config()
CONFIG_DIR.mkdir(parents=True, exist_ok=True)


def carregar_env(path: Path) -> None:
    """Carrega um .env simples sem depender de pacote externo."""
    if not path.exists():
        return

    for linha in path.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        chave, valor = linha.split("=", 1)
        chave = chave.strip()
        valor = valor.strip().strip('"').strip("'")
        if chave and chave not in os.environ:
            os.environ[chave] = valor


carregar_env(APP_DIR / ".env")
carregar_env(CONFIG_DIR / ".env")

GEMINI_STATE_PATH = str(CONFIG_DIR / "gemini_session.json")
GEMINI_TIMEOUT = int(os.getenv("GEMINI_TIMEOUT", "120"))
PAGE_SIZE = int(os.getenv("AUVO_PAGE_SIZE", "100"))
AUVO_TIMEOUT = int(os.getenv("AUVO_TIMEOUT", "90"))
AUVO_MAX_TENTATIVAS = 4
OUTPUT_DIR = str(Path(os.getenv("ANALISE_OUTPUT_DIR", str(Path.home() / "Downloads"))).expanduser())
Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
SALVAR_DADOS_BRUTOS = os.getenv("ANALISE_SALVAR_DADOS_BRUTOS", "").strip().lower() in {"1", "true", "sim", "yes"}


def resolver_template_excel() -> str:
    caminho_env = os.getenv("ANALISE_TEMPLATE_EXCEL")
    candidatos = []
    if caminho_env:
        candidatos.append(Path(caminho_env).expanduser())
    candidatos.extend([
        APP_DIR / TEMPLATE_EXCEL_FILENAME,
        Path(OUTPUT_DIR) / TEMPLATE_EXCEL_FILENAME,
    ])

    for caminho in candidatos:
        if caminho.exists():
            return str(caminho)

    return str(candidatos[0] if candidatos else APP_DIR / TEMPLATE_EXCEL_FILENAME)


TEMPLATE_EXCEL_PATH = resolver_template_excel()


def obter_credenciais_auvo() -> tuple[str, str]:
    api_key = os.getenv("AUVO_API_KEY", "").strip()
    api_token = os.getenv("AUVO_API_TOKEN", "").strip()

    if not api_key or not api_token:
        raise RuntimeError(
            "Credenciais da Auvo nao configuradas. Defina AUVO_API_KEY e AUVO_API_TOKEN "
            "em variaveis de ambiente ou em um arquivo .env ao lado do .exe/.py."
        )

    return api_key, api_token

# ─────────────────────────────────────────────
#  SELEÇÃO DE PERÍODO PELO USUÁRIO
# ─────────────────────────────────────────────
def solicitar_periodo() -> tuple[str, str]:
    """Solicita data inicial e final ao usuário e retorna strings ISO."""
    print("=" * 55)
    print("   ANÁLISE DE ATENDIMENTOS — Configuração de Período")
    print("=" * 55)
    print("Informe o período para buscar os atendimentos.\n")

    def ler_data(label: str) -> datetime:
        while True:
            entrada = input(f"  {label} (DD/MM/AAAA): ").strip()
            try:
                return datetime.strptime(entrada, "%d/%m/%Y")
            except ValueError:
                print("  ⚠️  Formato inválido. Use DD/MM/AAAA.\n")

    inicio = ler_data("Data inicial")
    fim    = ler_data("Data final  ")

    if fim < inicio:
        print("\n⚠️  A data final é anterior à inicial. As datas serão invertidas automaticamente.")
        inicio, fim = fim, inicio

    inicio_iso = inicio.strftime("%Y-%m-%dT00:00:00")
    fim_iso    = fim.strftime("%Y-%m-%dT23:59:59")

    print(f"\n  ✅ Período selecionado: {inicio.strftime('%d/%m/%Y')} → {fim.strftime('%d/%m/%Y')}\n")
    return inicio_iso, fim_iso


# ─────────────────────────────────────────────
#  AUTENTICAÇÃO AUVO
# ─────────────────────────────────────────────
def request_auvo(method: str, url: str, **kwargs) -> requests.Response:
    """Executa uma chamada na API Auvo com retry para lentidão/instabilidade."""
    kwargs.setdefault("timeout", AUVO_TIMEOUT)

    for tentativa in range(1, AUVO_MAX_TENTATIVAS + 1):
        try:
            resp = requests.request(method, url, **kwargs)

            if resp.status_code not in (429, 500, 502, 503, 504):
                return resp

            erro = f"HTTP {resp.status_code}: {resp.text[:200]}"

        except (
            requests.exceptions.Timeout,
            requests.exceptions.ConnectionError,
        ) as exc:
            erro = f"{type(exc).__name__}: {exc}"

        if tentativa == AUVO_MAX_TENTATIVAS:
            raise RuntimeError(
                f"API Auvo falhou após {AUVO_MAX_TENTATIVAS} tentativas. Último erro: {erro}"
            )

        espera = min(2 ** tentativa, 20)
        print(f"  ⚠️  API Auvo demorou/falhou ({erro}). Tentando novamente em {espera}s...")
        time.sleep(espera)


def autenticar() -> str:
    print("🔐 Autenticando na API Auvo...")
    url = "https://api.auvo.com.br/v2/login/"
    api_key, api_token = obter_credenciais_auvo()
    resp = request_auvo("post", url, json={"apiKey": api_key, "apiToken": api_token})
    resp.raise_for_status()
    token = resp.json()["result"]["accessToken"]
    print("  ✅ Autenticado com sucesso.\n")
    return token

# ─────────────────────────────────────────────
#  BUSCA DE ATENDIMENTOS (paginada)
# ─────────────────────────────────────────────
def buscar_atendimentos(token: str, start_date: str, end_date: str) -> list[dict]:
    """Retorna lista de atendimentos no período, de todos os atendentes, percorrendo todas as páginas."""
    print(f"📋 Buscando atendimentos de {start_date[:10]} a {end_date[:10]}...")

    url     = "https://api.auvo.com.br/v2/tasks/"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    todos   = []
    pagina  = 1

    while True:
        params = {
            "paramFilter": json.dumps({
                "startDate": start_date,
                "endDate":   end_date,
                # Sem filtro de usuário/atendente: busca todos os atendimentos acessíveis no período.
            }),
            "page":         pagina,
            "pageSize":     PAGE_SIZE,
            "order":        "asc",
            "selectfields": "",
        }

        resp = request_auvo("get", url, headers=headers, params=params)

        if resp.status_code != 200:
            print(f"  ❌ Erro {resp.status_code} na página {pagina}: {resp.text[:200]}")
            break

        data    = resp.json()
        result  = data.get("result", {})
        lista   = result.get("entityList", [])

        if not lista:
            break

        todos.extend(lista)

        total = (
            result.get("totalRecords")
            or result.get("totalItems")
            or result.get("totalEntity")
            or result.get("totalRecord")
            or data.get("totalRecords")
            or data.get("totalItems")
        )

        total_paginas = (
            result.get("totalPages")
            or result.get("pageCount")
            or data.get("totalPages")
            or data.get("pageCount")
        )

        total_label = total if total is not None else "desconhecido"
        print(f"  📄 Página {pagina} — {len(lista)} registros | Total até agora: {len(todos)}/{total_label}")

        if total is not None and len(todos) >= int(total):
            break

        if total_paginas is not None and pagina >= int(total_paginas):
            break

        if len(lista) < PAGE_SIZE:
            break

        pagina += 1

    print(f"\n  ✅ Total de atendimentos recuperados: {len(todos)}\n")
    return todos


# ─────────────────────────────────────────────
#  EXPORTAÇÃO EXCEL DETALHADA
# ─────────────────────────────────────────────
def limpar_texto_excel(valor) -> str:
    if valor is None:
        return ""

    texto = str(valor)
    return re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", "", texto)


def primeiro_valor(dados: dict, chaves: list[str], padrao: str = "") -> str:
    for chave in chaves:
        valor = dados.get(chave)
        if valor not in (None, ""):
            if isinstance(valor, dict):
                for subchave in ("name", "description", "fullName", "login", "id"):
                    if valor.get(subchave) not in (None, ""):
                        return limpar_texto_excel(valor.get(subchave))
            return limpar_texto_excel(valor)
    return padrao


def buscar_valor_recursivo(dados, palavras_chave: tuple[str, ...]) -> str:
    if isinstance(dados, dict):
        for chave, valor in dados.items():
            chave_normalizada = str(chave).lower()
            if all(palavra in chave_normalizada for palavra in palavras_chave) and valor not in (None, ""):
                if isinstance(valor, dict):
                    nome = primeiro_valor(valor, ["name", "description", "fullName", "login", "id"])
                    if nome:
                        return nome
                elif not isinstance(valor, (list, dict)):
                    return limpar_texto_excel(valor)

        for valor in dados.values():
            encontrado = buscar_valor_recursivo(valor, palavras_chave)
            if encontrado:
                return encontrado

    elif isinstance(dados, list):
        for item in dados:
            encontrado = buscar_valor_recursivo(item, palavras_chave)
            if encontrado:
                return encontrado

    return ""


def obter_atendente(tarefa: dict) -> str:
    atendente = primeiro_valor(tarefa, [
        "userToName",
        "userToDescription",
        "userTo",
        "responsibleName",
        "executorName",
        "employeeName",
        "technicianName",
        "assignedUserName",
        "userName",
        "userFromName",
    ])

    if atendente:
        return atendente

    for palavras in (("user", "name"), ("user", "description"), ("employee", "name"), ("technician", "name")):
        atendente = buscar_valor_recursivo(tarefa, palavras)
        if atendente:
            return atendente

    return "N/D"


def obter_id_atendente(tarefa: dict) -> str:
    return primeiro_valor(tarefa, [
        "idUserTo",
        "userToId",
        "userTo",
        "responsibleId",
        "executorId",
        "employeeId",
        "technicianId",
        "assignedUserId",
        "userId",
    ])


def obter_relato(tarefa: dict) -> str:
    return limpar_texto_excel(
        tarefa.get("taskReport")
        or tarefa.get("report")
        or tarefa.get("description")
        or tarefa.get("orientation")
        or ""
    )


def achatar_json(dados, prefixo: str = "") -> dict:
    campos = {}

    if isinstance(dados, dict):
        for chave, valor in dados.items():
            nome = f"{prefixo}.{chave}" if prefixo else str(chave)
            if isinstance(valor, dict):
                campos.update(achatar_json(valor, nome))
            elif isinstance(valor, list):
                campos[nome] = json.dumps(valor, ensure_ascii=False)
            else:
                campos[nome] = limpar_texto_excel(valor)

    return campos


def montar_linhas_excel(atendimentos: list[dict]) -> list[list]:
    cabecalho_base = [
        "ID atendimento",
        "Data",
        "Cliente",
        "ID cliente",
        "Atendente",
        "ID atendente",
        "Tipo",
        "Status",
        "Prioridade",
        "Endereco",
        "Cidade",
        "UF",
        "Orientacao",
        "Relato/Descricao",
    ]

    campos_ja_mapeados = {
        "taskID", "taskId", "id", "idTask",
        "taskDate", "date", "scheduleDate", "checkInDate",
        "customerDescription", "customerName", "clientName", "customer",
        "customerId", "idCustomer", "clientId",
        "userToName", "userToDescription", "userTo", "responsibleName", "executorName",
        "employeeName", "technicianName", "assignedUserName", "userName", "userFromName",
        "idUserTo", "userToId", "responsibleId", "executorId", "employeeId", "technicianId",
        "assignedUserId", "userId",
        "taskTypeDescription", "taskType", "typeDescription", "type",
        "taskStatusDescription", "statusDescription", "status", "taskStatus",
        "priority", "priorityDescription", "address", "street", "location",
        "city", "cityDescription", "state", "uf", "orientation",
        "taskReport", "report", "description",
    }

    dados_achatados = [achatar_json(tarefa) for tarefa in atendimentos]
    campos_extras = sorted({
        chave
        for dados in dados_achatados
        for chave in dados
        if chave.split(".", 1)[0] not in campos_ja_mapeados
    })

    cabecalho = cabecalho_base + campos_extras + [
        "Dados brutos (JSON)",
    ]

    linhas = [cabecalho]

    for tarefa, extras in zip(atendimentos, dados_achatados):
        linha_base = [
            primeiro_valor(tarefa, ["taskID", "taskId", "id", "idTask"]),
            primeiro_valor(tarefa, ["taskDate", "date", "scheduleDate", "checkInDate"]),
            primeiro_valor(tarefa, ["customerDescription", "customerName", "clientName", "customer"]),
            primeiro_valor(tarefa, ["customerId", "idCustomer", "clientId"]),
            obter_atendente(tarefa),
            obter_id_atendente(tarefa),
            primeiro_valor(tarefa, ["taskTypeDescription", "taskType", "typeDescription", "type"]),
            primeiro_valor(tarefa, ["taskStatusDescription", "statusDescription", "status", "taskStatus"]),
            primeiro_valor(tarefa, ["priority", "priorityDescription"]),
            primeiro_valor(tarefa, ["address", "street", "location"]),
            primeiro_valor(tarefa, ["city", "cityDescription"]),
            primeiro_valor(tarefa, ["state", "uf"]),
            limpar_texto_excel(tarefa.get("orientation") or ""),
            obter_relato(tarefa),
        ]

        linhas.append([
            *linha_base,
            *(extras.get(campo, "") for campo in campos_extras),
            limpar_texto_excel(json.dumps(tarefa, ensure_ascii=False)),
        ])

    return linhas


def montar_resumo_atendentes(atendimentos: list[dict]) -> list[list]:
    contador = Counter(obter_atendente(tarefa) for tarefa in atendimentos)
    linhas = [["Atendente", "Quantidade de atendimentos", "% do total"]]
    total = len(atendimentos) or 1

    for atendente, qtd in contador.most_common():
        linhas.append([atendente, qtd, round(qtd / total, 4)])

    return linhas


def montar_linhas_analise_ia(analise_ia: str, start_date: str, end_date: str) -> list[list]:
    linhas = [
        ["Campo", "Conteudo"],
        ["Periodo", f"{start_date[:10]} a {end_date[:10]}"],
        ["Gerado em", datetime.now().strftime("%d/%m/%Y %H:%M:%S")],
        ["", ""],
        ["Linha", "Texto retornado pela IA"],
    ]

    for i, linha in enumerate((analise_ia or "").splitlines(), 1):
        if linha.strip():
            linhas.append([i, linha.strip()])

    if len(linhas) == 5:
        linhas.append([1, "Analise da IA nao disponivel."])

    return linhas


def coluna_excel(indice: int) -> str:
    letras = ""
    while indice:
        indice, resto = divmod(indice - 1, 26)
        letras = chr(65 + resto) + letras
    return letras


def xml_celula(valor, linha: int, coluna: int, estilo: int | None = None) -> str:
    ref = f"{coluna_excel(coluna)}{linha}"
    estilo_attr = f' s="{estilo}"' if estilo is not None else ""

    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        return f'<c r="{ref}"{estilo_attr}><v>{valor}</v></c>'

    texto = escape(limpar_texto_excel(valor))
    return f'<c r="{ref}" t="inlineStr"{estilo_attr}><is><t>{texto}</t></is></c>'


def xml_planilha(linhas: list[list], larguras: list[int]) -> str:
    cols = "".join(
        f'<col min="{i}" max="{i}" width="{largura}" customWidth="1"/>'
        for i, largura in enumerate(larguras, 1)
    )

    rows = []
    for i, linha in enumerate(linhas, 1):
        estilo = 1 if i == 1 else None
        celulas = "".join(xml_celula(valor, i, j, estilo) for j, valor in enumerate(linha, 1))
        rows.append(f'<row r="{i}">{celulas}</row>')

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
<cols>{cols}</cols>
<sheetData>{''.join(rows)}</sheetData>
<autoFilter ref="A1:{coluna_excel(len(linhas[0]))}{len(linhas)}"/>
</worksheet>'''


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


def limpar_markdown(texto: str) -> str:
    texto = re.sub(r"[*_`#]+", "", texto or "")
    texto = re.sub(r"^\s*[-•]\s*", "", texto)
    texto = re.sub(r"^\s*\d+[.)ºª-]*\s*", "", texto)
    return texto.strip()


def normalizar_texto(texto: str) -> str:
    mapa = str.maketrans("áàãâéêíóôõúçÁÀÃÂÉÊÍÓÔÕÚÇ", "aaaaeeioooucAAAAEEIOOOUC")
    return (texto or "").translate(mapa).lower()


def extrair_secao(texto: str, titulo: str, proximos_titulos: tuple[str, ...]) -> str:
    normalizado = normalizar_texto(texto)
    inicio = normalizado.find(normalizar_texto(titulo))
    if inicio == -1:
        return ""

    inicio_conteudo = texto.find("\n", inicio)
    if inicio_conteudo == -1:
        inicio_conteudo = inicio + len(titulo)
    else:
        inicio_conteudo += 1

    fim = len(texto)
    for proximo in proximos_titulos:
        pos = normalizado.find(normalizar_texto(proximo), inicio + len(titulo))
        if pos != -1:
            fim = min(fim, pos)

    secao = texto[inicio_conteudo:fim].strip()
    return re.sub(r"\s+\d+[.)]\s*$", "", secao).strip()


def dividir_itens(texto: str, limite: int) -> list[str]:
    itens = []
    atual = []

    for linha in (texto or "").splitlines():
        linha_limpa = linha.strip()
        if not linha_limpa:
            continue

        if re.match(r"^(\d+[.)ºª-]|\-|\*)\s+", linha_limpa) and atual:
            itens.append(" ".join(atual).strip())
            atual = [linha_limpa]
        else:
            atual.append(linha_limpa)

    if atual:
        itens.append(" ".join(atual).strip())

    return [limpar_markdown(item) for item in itens if limpar_markdown(item)][:limite]


def primeira_linha_util(texto: str, padrao: str = "") -> str:
    for linha in (texto or "").splitlines():
        linha = limpar_markdown(linha)
        if linha and not linha.endswith(":"):
            return linha
    return padrao


def resumir_texto(texto: str, limite: int = 180) -> str:
    texto = re.sub(r"\s+", " ", limpar_texto_excel(limpar_markdown(texto))).strip()
    if len(texto) <= limite:
        return texto
    return texto[: limite - 3].rstrip() + "..."


def extrair_problemas_frequentes(analise_ia: str, atendimentos: list[dict]) -> list[dict]:
    secao = extrair_secao(
        analise_ia,
        "PROBLEMAS MAIS FREQUENTES",
        ("CLIENTES QUE MAIS RECLAMAM", "CORRELAÇÕES IMPORTANTES", "CORRELACOES IMPORTANTES", "RESUMO EXECUTIVO"),
    )
    itens = dividir_itens(secao, 5)
    total = len(atendimentos) or 1
    problemas = []

    for item in itens:
        item_sem_titulo = re.sub(r"^problemas mais frequentes\s*:?", "", item, flags=re.I).strip()
        qtd_match = re.search(r"(\d+)\s*(?:ocorr|atend|caso|registro)", item_sem_titulo, flags=re.I)
        pct_match = re.search(r"(\d+(?:[,.]\d+)?)\s*%", item_sem_titulo)
        clientes_match = re.search(r"(?:clientes afetados|exemplos de clientes|clientes|exemplos)\s*:?\s*(.+)$", item_sem_titulo, flags=re.I)

        nome = item_sem_titulo
        nome = re.split(r"\s+-\s+|\s+–\s+|:", nome, maxsplit=1)[0]
        nome = re.sub(r"\([^)]*(?:ocorr|atend|caso|%)[^)]*\)", "", nome, flags=re.I).strip()

        qtd = int(qtd_match.group(1)) if qtd_match else ""
        pct = pct_match.group(1).replace(".", ",") + "%" if pct_match else (f"{round(qtd / total * 100, 1):.1f}%".replace(".", ",") if qtd else "")
        clientes = resumir_texto(clientes_match.group(1), 90) if clientes_match else resumir_texto(item_sem_titulo, 90)
        clientes = re.sub(r"\s+\d+[.)]\s*$", "", clientes).strip()

        if nome:
            problemas.append({
                "problema": resumir_texto(nome, 60),
                "qtd": qtd,
                "pct": pct,
                "clientes": clientes,
            })

    if problemas:
        return problemas[:5]

    contador = Counter(
        primeiro_valor(tarefa, ["taskTypeDescription", "taskType", "typeDescription", "type"], "Atendimento")
        for tarefa in atendimentos
    )
    for nome, qtd in contador.most_common(5):
        clientes = [
            primeiro_valor(tarefa, ["customerDescription", "customerName", "clientName", "customer"], "N/D")
            for tarefa in atendimentos
            if primeiro_valor(tarefa, ["taskTypeDescription", "taskType", "typeDescription", "type"], "Atendimento") == nome
        ][:3]
        problemas.append({
            "problema": resumir_texto(nome, 60),
            "qtd": qtd,
            "pct": f"{round(qtd / total * 100, 1):.1f}%".replace(".", ","),
            "clientes": ", ".join(clientes),
        })

    return problemas


def extrair_clientes_top(atendimentos: list[dict]) -> list[dict]:
    por_cliente = {}
    for tarefa in atendimentos:
        cliente = primeiro_valor(tarefa, ["customerDescription", "customerName", "clientName", "customer"], "N/D")
        relato = obter_relato(tarefa) or limpar_texto_excel(tarefa.get("orientation") or "")
        por_cliente.setdefault(cliente, {"qtd": 0, "ocorrencias": []})
        por_cliente[cliente]["qtd"] += 1
        if relato:
            por_cliente[cliente]["ocorrencias"].append(resumir_texto(relato, 120))

    ranking = sorted(por_cliente.items(), key=lambda item: item[1]["qtd"], reverse=True)[:10]
    return [
        {
            "cliente": cliente,
            "qtd": dados["qtd"],
            "ocorrencia": dados["ocorrencias"][0] if dados["ocorrencias"] else "Sem relato detalhado no atendimento.",
        }
        for cliente, dados in ranking
    ]


def extrair_resumo_executivo(analise_ia: str) -> list[tuple[str, str]]:
    secao = extrair_secao(analise_ia, "RESUMO EXECUTIVO", ())
    if not secao:
        secao = analise_ia

    frases = [
        resumir_texto(frase, 180)
        for frase in re.split(r"(?<=[.!?])\s+", limpar_markdown(secao))
        if len(limpar_markdown(frase)) > 25
    ]
    titulos = ["Resumo Executivo", "Ponto de Atenção", "Tendência do Período", "Próxima Ação"]
    return list(zip(titulos, frases[:4]))


def extrair_correlacoes(analise_ia: str) -> tuple[list[tuple[str, str]], str]:
    secao = extrair_secao(analise_ia, "CORRELAÇÕES IMPORTANTES", ("RESUMO EXECUTIVO",))
    if not secao:
        secao = extrair_secao(analise_ia, "CORRELACOES IMPORTANTES", ("RESUMO EXECUTIVO",))

    itens = dividir_itens(secao, 4)
    correlacoes = []
    for i, item in enumerate(itens[:4], 1):
        titulo = re.split(r":|\s+-\s+|\s+–\s+", item, maxsplit=1)[0]
        descricao = item[len(titulo):].lstrip(":-– ").strip() or item
        correlacoes.append((resumir_texto(titulo or f"Insight {i}", 45), resumir_texto(descricao, 210)))

    if not correlacoes:
        resumo = extrair_resumo_executivo(analise_ia)
        correlacoes = [(titulo, texto) for titulo, texto in resumo[:4]]

    recomendacao = ""
    for linha in analise_ia.splitlines():
        if "recomend" in normalizar_texto(linha):
            recomendacao = resumir_texto(linha, 160)
            break

    if not recomendacao and correlacoes:
        recomendacao = "Recomendação principal: priorizar os temas mais recorrentes e os clientes com maior volume no período."

    return correlacoes[:4], recomendacao


def coordenadas_celula(ref: str) -> tuple[str, int]:
    match = re.match(r"([A-Z]+)(\d+)", ref)
    if not match:
        raise ValueError(f"Referencia de celula invalida: {ref}")
    return match.group(1), int(match.group(2))


def indice_coluna_excel(letras: str) -> int:
    indice = 0
    for letra in letras:
        indice = indice * 26 + ord(letra) - 64
    return indice


def definir_celula_xml(root: ET.Element, ref: str, valor) -> None:
    ns = {"x": MAIN_NS}
    _, linha_num = coordenadas_celula(ref)
    sheet_data = root.find("x:sheetData", ns)
    if sheet_data is None:
        return

    linha = sheet_data.find(f"x:row[@r='{linha_num}']", ns)
    if linha is None:
        linha = ET.Element(f"{{{MAIN_NS}}}row", {"r": str(linha_num)})
        sheet_data.append(linha)

    celula = linha.find(f"x:c[@r='{ref}']", ns)
    if celula is None:
        celula = ET.Element(f"{{{MAIN_NS}}}c", {"r": ref})
        ref_col, _ = coordenadas_celula(ref)
        ref_idx = indice_coluna_excel(ref_col)
        inserido = False
        for pos, existente in enumerate(list(linha)):
            col, _ = coordenadas_celula(existente.attrib.get("r", "A1"))
            if indice_coluna_excel(col) > ref_idx:
                linha.insert(pos, celula)
                inserido = True
                break
        if not inserido:
            linha.append(celula)

    estilo = celula.attrib.get("s")
    celula.clear()
    celula.set("r", ref)
    if estilo is not None:
        celula.set("s", estilo)

    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        ET.SubElement(celula, f"{{{MAIN_NS}}}v").text = str(valor)
        return

    celula.set("t", "inlineStr")
    inline = ET.SubElement(celula, f"{{{MAIN_NS}}}is")
    texto = ET.SubElement(inline, f"{{{MAIN_NS}}}t")
    texto.text = limpar_texto_excel(valor)
    texto.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def atualizar_planilha_xml(xml_texto: str, valores: dict[str, object]) -> str:
    root = ET.fromstring(xml_texto)
    for ref, valor in valores.items():
        definir_celula_xml(root, ref, valor)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True).decode("utf-8")


def preencher_modelo_excel(
    modelo: str,
    destino: str,
    valores_por_planilha: dict[str, dict[str, object]],
) -> None:
    with zipfile.ZipFile(modelo, "r") as origem:
        conteudos = {info.filename: origem.read(info.filename) for info in origem.infolist()}

    for caminho, valores in valores_por_planilha.items():
        xml_atual = conteudos[caminho].decode("utf-8")
        conteudos[caminho] = atualizar_planilha_xml(xml_atual, valores).encode("utf-8")

    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as saida:
        for caminho, conteudo in conteudos.items():
            saida.writestr(caminho, conteudo)


def montar_valores_modelo(
    atendimentos: list[dict],
    start_date: str,
    end_date: str,
    analise_ia: str,
) -> dict[str, dict[str, object]]:
    problemas = extrair_problemas_frequentes(analise_ia, atendimentos)
    clientes = extrair_clientes_top(atendimentos)
    resumo = extrair_resumo_executivo(analise_ia)
    correlacoes, recomendacao = extrair_correlacoes(analise_ia)

    inicio = datetime.strptime(start_date[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
    fim = datetime.strptime(end_date[:10], "%Y-%m-%d").strftime("%d/%m/%Y")

    visao = {
        "D3": f"{inicio} a {fim}",
        "D4": datetime.now().strftime("%d/%m/%Y às %H:%M"),
    }
    for i in range(5):
        linha = 9 + i
        item = problemas[i] if i < len(problemas) else {"problema": "", "qtd": "", "pct": "", "clientes": ""}
        visao[f"B{linha}"] = f"{i + 1}º"
        visao[f"C{linha}"] = item["problema"]
        visao[f"D{linha}"] = item["qtd"] if item["qtd"] != "" else ""
        visao[f"E{linha}"] = item["pct"]
        visao[f"F{linha}"] = item["clientes"]

    for i, linha in enumerate((18, 21, 24, 27)):
        titulo, texto = resumo[i] if i < len(resumo) else (f"Insight {i + 1}", "")
        visao[f"B{linha}"] = titulo
        visao[f"B{linha + 1}"] = texto

    top_clientes = {}
    medalhas = ["🥇", "🥈", "🥉"]
    for i in range(10):
        linha = 5 + i
        cliente = clientes[i] if i < len(clientes) else {"cliente": "", "qtd": "", "ocorrencia": ""}
        top_clientes[f"B{linha}"] = medalhas[i] if i < 3 else i + 1
        top_clientes[f"C{linha}"] = cliente["cliente"]
        top_clientes[f"D{linha}"] = cliente["qtd"]
        top_clientes[f"E{linha}"] = cliente["ocorrencia"]

    corr = {}
    for i, linha in enumerate((4, 8, 12, 16)):
        titulo, texto = correlacoes[i] if i < len(correlacoes) else (f"Insight {i + 1}", "")
        corr[f"B{linha}"] = titulo
        corr[f"B{linha + 1}"] = texto
    corr["B22"] = recomendacao

    return {
        "xl/worksheets/sheet1.xml": visao,
        "xl/worksheets/sheet2.xml": top_clientes,
        "xl/worksheets/sheet3.xml": corr,
    }


def exportar_excel_detalhado(
    atendimentos: list[dict],
    start_date: str,
    end_date: str,
    analise_ia: str = "",
) -> str:
    inicio = start_date[:10].replace("-", "")
    fim    = end_date[:10].replace("-", "")
    nome   = os.path.join(OUTPUT_DIR, f"Relatorio_Atendimentos_{inicio}_{fim}.xlsx")

    if not os.path.exists(TEMPLATE_EXCEL_PATH):
        raise FileNotFoundError(f"Modelo do Excel nao encontrado: {TEMPLATE_EXCEL_PATH}")

    valores = montar_valores_modelo(atendimentos, start_date, end_date, analise_ia)

    try:
        preencher_modelo_excel(TEMPLATE_EXCEL_PATH, nome, valores)
    except PermissionError:
        nome = os.path.join(
            OUTPUT_DIR,
            f"Relatorio_Atendimentos_{inicio}_{fim}_{datetime.now().strftime('%H%M%S')}.xlsx",
        )
        preencher_modelo_excel(TEMPLATE_EXCEL_PATH, nome, valores)

    return nome


# ─────────────────────────────────────────────
#  MONTAGEM DO PROMPT
# ─────────────────────────────────────────────
def montar_prompt(atendimentos: list[dict], start_date: str, end_date: str) -> str:
    """Monta o texto que será colado no Gemini."""
    linhas = []
    for i, t in enumerate(atendimentos, 1):
        cliente   = t.get("customerDescription", "N/D")
        data_raw  = t.get("taskDate", "")
        data_fmt  = data_raw[:10] if data_raw else "N/D"
        descricao = (t.get("orientation") or "").strip()
        relato    = (t.get("taskReport") or t.get("report") or t.get("description") or "").strip()
        tipo      = t.get("taskType", "")

        linhas.append(
            f"[{i}] Cliente: {cliente} | Data: {data_fmt} | Tipo: {tipo}\n"
            f"    Descrição: {descricao}\n"
            f"    Relato: {relato}"
        )

    bloco = "\n\n".join(linhas)

    periodo_fmt = f"{start_date[:10]} a {end_date[:10]}"

    prompt = f"""Você é um analista de suporte técnico especialista em identificar padrões de reclamações.

Abaixo estão {len(atendimentos)} atendimentos realizados no período de {periodo_fmt}.

Analise TODOS os registros e retorne um relatório estruturado contendo:

1. **PROBLEMAS MAIS FREQUENTES**
   - Liste os top 5 problemas/reclamações mais recorrentes
   - Para cada problema: nome do problema, quantidade de ocorrências e percentual sobre o total
   - Exemplos de clientes afetados (até 3 por problema)

2. **CLIENTES QUE MAIS RECLAMAM**
   - Liste os top 10 clientes com maior número de atendimentos no período
   - Para cada cliente: nome, quantidade de atendimentos e principais problemas relatados

3. **CORRELAÇÕES IMPORTANTES**
   - Algum problema afeta desproporcionalmente algum grupo de clientes?
   - Existe algum padrão temporal (dia da semana, período do mês)?

4. **RESUMO EXECUTIVO**
   - Parágrafo curto (3–5 linhas) com os principais insights do período.

Seja direto, use os dados reais dos atendimentos. Não invente informações.

--- ATENDIMENTOS ---

{bloco}

--- FIM DOS ATENDIMENTOS ---"""

    return prompt


# ─────────────────────────────────────────────
#  INTEGRAÇÃO COM GEMINI VIA PLAYWRIGHT
# ─────────────────────────────────────────────
def chamar_gemini(prompt: str, tentativas: int = 3) -> str:
    """Abre o Gemini no Edge via Playwright, cola o prompt e retorna a resposta."""

    for i in range(tentativas):
        browser = None
        try:
            print(f"🤖 Tentativa {i+1}/{tentativas}: Acessando Gemini...")

            with sync_playwright() as p:
                browser = p.chromium.launch(
                    channel="msedge",
                    headless=False,
                    args=[
                        "--disable-gpu",
                        "--no-sandbox",
                        "--disable-blink-features=AutomationControlled",
                        "--disable-dev-shm-usage",
                        "--no-first-run",
                        "--no-default-browser-check",
                    ],
                )

                context = browser.new_context(
                    storage_state=GEMINI_STATE_PATH if os.path.exists(GEMINI_STATE_PATH) else None,
                    viewport={"width": 1280, "height": 720},
                    user_agent=(
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/120.0.0.0 Safari/537.36"
                    ),
                    java_script_enabled=True,
                )

                context.add_init_script("""
                    Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
                    Object.defineProperty(navigator, 'plugins',   { get: () => [1, 2, 3] });
                    Object.defineProperty(navigator, 'languages', { get: () => ['pt-BR', 'pt', 'en-US'] });
                    window.chrome = { runtime: {} };
                """)

                page = context.new_page()

                # Bloqueia recursos pesados para acelerar
                page.route(
                    "**/*",
                    lambda route: route.abort()
                    if route.request.resource_type in ["image", "stylesheet", "font", "media"]
                    else route.continue_(),
                )

                page.goto("https://gemini.google.com/app", wait_until="domcontentloaded", timeout=60_000)

                # Login manual se sessão não existir
                if "accounts.google" in page.url or "signin" in page.url:
                    print("\n⚠️  Sessão não encontrada. Faça login manualmente no navegador que abriu.")
                    print("    Após logar no Gemini, volte aqui e pressione ENTER...")
                    input()
                    context.storage_state(path=GEMINI_STATE_PATH)
                    print("  ✅ Sessão salva para próximas execuções!\n")

                # Localiza caixa de texto
                input_box = page.locator('div[contenteditable="true"]').first
                input_box.wait_for(state="visible", timeout=30_000)
                input_box.click()

                # Cola o prompt via clipboard (mais confiável para textos longos)
                pyperclip.copy(prompt)
                page.keyboard.press("Control+v")
                time.sleep(1.0)
                page.keyboard.press("Enter")

                print("⏳ Aguardando resposta do Gemini (pode demorar alguns minutos)...")

                # Aguarda o botão de copiar aparecer (sinal de que a resposta chegou)
                try:
                    page.wait_for_selector('[data-test-id*="copy"]', timeout=GEMINI_TIMEOUT * 1_000)
                except Exception:
                    print(f"⚠️  Gemini não respondeu em {GEMINI_TIMEOUT}s.")
                    return ""

                # Aguarda estabilização da resposta
                texto_anterior = ""
                estavel        = 0
                for _ in range(max(60, int(GEMINI_TIMEOUT / 0.5))):
                    time.sleep(0.5)
                    try:
                        texto_atual = page.locator("model-response").last.inner_text()
                        if texto_atual == texto_anterior and len(texto_atual.strip()) > 50:
                            estavel += 1
                            if estavel >= 4:
                                print(f"  ✅ Resposta estabilizou! ({len(texto_atual)} chars)")
                                break
                        else:
                            estavel = 0
                            print(f"  ⏳ Gerando... ({len(texto_atual)} chars)")
                        texto_anterior = texto_atual
                    except Exception:
                        pass

                # Clica no botão copiar do último bloco
                ultimo_botao = page.locator('[data-test-id*="copy"]').last
                ultimo_botao.scroll_into_view_if_needed()
                ultimo_botao.wait_for(state="visible", timeout=5_000)
                ultimo_botao.click(force=True)

                # Recupera texto do clipboard
                resposta = ""
                for _ in range(15):
                    time.sleep(0.2)
                    resposta = pyperclip.paste()
                    if resposta and len(resposta.strip()) >= 20:
                        break

                if not resposta or len(resposta.strip()) < 20:
                    print(f"  ⚠️  Tentativa {i+1}: Resposta vazia ou muito curta.")
                    continue

                # Salva sessão atualizada
                context.storage_state(path=GEMINI_STATE_PATH)
                print(f"  ✅ Gemini OK ({len(resposta.strip())} chars)\n")
                return resposta.strip()

        except Exception as e:
            print(f"  ⚠️  Tentativa {i+1} falhou: {e}")
            time.sleep(1.5)

        finally:
            if browser:
                try:
                    browser.close()
                except Exception:
                    pass

    return ""


# ─────────────────────────────────────────────
#  SALVAR RELATÓRIO EM ARQUIVO
# ─────────────────────────────────────────────
def salvar_relatorio(conteudo: str, start_date: str, end_date: str) -> str:
    inicio = start_date[:10].replace("-", "")
    fim    = end_date[:10].replace("-", "")
    nome   = f"relatorio_atendimentos_{inicio}_{fim}.txt"

    with open(nome, "w", encoding="utf-8") as f:
        f.write(f"RELATÓRIO DE ANÁLISE DE ATENDIMENTOS\n")
        f.write(f"Período: {start_date[:10]} → {end_date[:10]}\n")
        f.write(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write("=" * 60 + "\n\n")
        f.write(conteudo)

    return nome


def salvar_atendimentos_txt(
    atendimentos: list[dict],
    prompt: str,
    start_date: str,
    end_date: str,
) -> str:
    """Salva o prompt completo e o JSON bruto dos atendimentos, sem truncar texto."""
    inicio = start_date[:10].replace("-", "")
    fim    = end_date[:10].replace("-", "")
    nome   = os.path.join(OUTPUT_DIR, f"atendimentos_completos_{inicio}_{fim}.txt")

    with open(nome, "w", encoding="utf-8", newline="\n") as f:
        f.write("ATENDIMENTOS COMPLETOS DO PERIODO\n")
        f.write(f"Periodo: {start_date[:10]} -> {end_date[:10]}\n")
        f.write(f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
        f.write(f"Total de atendimentos: {len(atendimentos)}\n")
        f.write(f"Caracteres do prompt: {len(prompt)}\n")
        f.write("=" * 80 + "\n\n")
        f.write("PROMPT COMPLETO ENVIADO PARA ANALISE\n")
        f.write("=" * 80 + "\n\n")
        f.write(prompt)
        f.write("\n\n")
        f.write("=" * 80 + "\n")
        f.write("JSON BRUTO COMPLETO RETORNADO PELA API\n")
        f.write("=" * 80 + "\n\n")
        json.dump(atendimentos, f, ensure_ascii=False, indent=2)

    return nome


# ─────────────────────────────────────────────
#  EXECUÇÃO PRINCIPAL
# ─────────────────────────────────────────────
def main():
    # 1. Período
    start_date, end_date = solicitar_periodo()

    # 2. Autenticação
    token = autenticar()

    # 3. Buscar atendimentos
    atendimentos = buscar_atendimentos(token, start_date, end_date)

    if not atendimentos:
        print("❌ Nenhum atendimento encontrado no período informado.")
        return

    # 4. Montar prompt
    print("📝 Montando prompt para análise...\n")
    prompt = montar_prompt(atendimentos, start_date, end_date)

    print(f"  Prompt gerado: {len(prompt)} caracteres | {len(atendimentos)} atendimentos incluídos\n")
    if SALVAR_DADOS_BRUTOS:
        nome_txt = salvar_atendimentos_txt(atendimentos, prompt, start_date, end_date)
        print(f"  📄 TXT completo salvo em: {nome_txt}\n")
    else:
        print("  🔒 TXT com prompt/JSON bruto não foi salvo. Defina ANALISE_SALVAR_DADOS_BRUTOS=1 para ativar.\n")

    # 5. Enviar para Gemini
    print("=" * 55)
    print("   ENVIANDO PARA O GEMINI — Aguarde...")
    print("=" * 55 + "\n")
    resposta = chamar_gemini(prompt)

    if not resposta:
        print("❌ Não foi possível obter resposta do Gemini após todas as tentativas.")
        return

    # 6. Exibir e salvar no Excel
    print("\n" + "=" * 55)
    print("   ANÁLISE DO GEMINI")
    print("=" * 55 + "\n")
    print(resposta)

    print("\n📊 Exportando tabela Excel detalhada com a análise da IA...\n")
    nome_excel = exportar_excel_detalhado(atendimentos, start_date, end_date, resposta)
    print(f"\n\n✅ Tabela Excel salva em: {nome_excel}")


if __name__ == "__main__":
    main()

