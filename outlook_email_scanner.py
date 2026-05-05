"""
Script de varredura de emails do Outlook 365 via exchangelib (OAuth2).
Filtra emails do remetente noreply@uber.com com assunto e corpo
contendo "Your file export is ready" nas últimas 24 horas.
"""

from exchangelib import Account, DELEGATE, Configuration, Identity, OAuth2Credentials
from exchangelib.queryset import Q
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import msal
import requests
import csv
import json
import base64
import re
import os
from urllib.parse import urlparse, parse_qs, unquote
from dotenv import load_dotenv

load_dotenv()

# ============================================================
# Configurações
# ============================================================
EMAIL = os.getenv("OUTLOOK_EMAIL")
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
TENANT_ID = os.getenv("AZURE_TENANT_ID")

SHAREPOINT_URL = os.getenv("SHAREPOINT_URL")

FILTRO_REMETENTE = "noreply@uber.com"
FILTRO_ASSUNTO = "Relatório da Uber para Empresas da empresa AUTVIX ENGENHARIA E CONSULTORIA LTDA"
FILTRO_BODY = "O relatório da Uber para Empresas da empresa"

FUSO_BRASILIA = ZoneInfo("America/Sao_Paulo")
DIRETORIO_DOWNLOADS = os.path.join(os.path.dirname(__file__), "downloads")
ARQUIVO_DEBUG = os.path.join(os.path.dirname(__file__), "debug_emails.txt")
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

MESES_PT = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO",
}


def conectar_outlook(email: str, client_id: str, client_secret: str, tenant_id: str) -> Account:
    """Conecta à conta Outlook 365 via EWS com OAuth2."""
    credenciais = OAuth2Credentials(
        client_id=client_id,
        client_secret=client_secret,
        tenant_id=tenant_id,
        identity=Identity(primary_smtp_address=email),
    )
    config = Configuration(
        server="outlook.office365.com",
        credentials=credenciais,
    )
    conta = Account(
        primary_smtp_address=email,
        config=config,
        autodiscover=False,
        access_type=DELEGATE,
    )
    return conta


def buscar_email_mais_recente(conta: Account) -> dict | None:
    """
    Busca o email mais recente na caixa de entrada das últimas 24 horas
    que atenda aos critérios de filtro. Também grava em arquivo de debug
    todos os emails encontrados do remetente.

    Retorna um dicionário com os dados do email ou None se não encontrar.
    """
    agora = datetime.now(tz=FUSO_BRASILIA)
    inicio = agora - timedelta(days=10)
    inicio_debug = agora - timedelta(days=10)

    # Busca apenas pelo remetente nos últimos 10 dias (debug, sem filtrar assunto)
    debug_filtro = (
        Q(sender__icontains=FILTRO_REMETENTE)
        & Q(datetime_received__gte=inicio_debug)
    )
    emails_remetente = list(
        conta.inbox.filter(debug_filtro).order_by("-datetime_received")
    )

    timestamp = datetime.now(tz=FUSO_BRASILIA).strftime("%d/%m/%Y %H:%M:%S")
    with open(ARQUIVO_DEBUG, "a", encoding="utf-8") as f:
        f.write(f"\n{'=' * 70}\n")
        f.write(f"Execução em: {timestamp}\n")
        f.write(f"Janela: últimos 10 dias | Remetente filtrado: {FILTRO_REMETENTE}\n")
        f.write(f"Emails encontrados do remetente: {len(emails_remetente)}\n")
        f.write(f"{'=' * 70}\n")
        for i, email in enumerate(emails_remetente, 1):
            data_recv = email.datetime_received.astimezone(FUSO_BRASILIA).strftime("%d/%m/%Y %H:%M:%S")
            f.write(f"\n[{i}] {data_recv}\n")
            f.write(f"    Remetente : {email.sender}\n")
            f.write(f"    Assunto   : {email.subject}\n")
            f.write(f"    Tem anexos: {email.has_attachments}\n")
            f.write(f"    Msg-ID    : {email.message_id}\n")

    print(f"Debug: {len(emails_remetente)} email(s) do remetente registrados em '{ARQUIVO_DEBUG}'")

    # Aplica filtros completos (assunto + corpo) para encontrar o email-alvo
    filtro = (
        Q(sender__icontains=FILTRO_REMETENTE)
        & Q(subject__icontains=FILTRO_ASSUNTO)
        & Q(datetime_received__gte=inicio)
    )

    emails_encontrados = conta.inbox.filter(filtro).order_by("-datetime_received")

    for email in emails_encontrados:
        corpo = email.body or ""

        if FILTRO_BODY.lower() not in corpo.lower():
            continue

        return {
            "remetente": str(email.sender),
            "assunto": email.subject,
            "data_recebimento": email.datetime_received.astimezone(FUSO_BRASILIA).strftime("%d/%m/%Y %H:%M:%S"),
            "corpo": corpo,
            "tem_anexos": email.has_attachments,
            "id": email.message_id,
        }

    return None


def extrair_link_download(corpo_html: str) -> str | None:
    """
    Extrai o link de download do CSV no corpo HTML do email.
    Busca por padrões em português ("Baixar CSV") ou inglês ("Download File"/"Download CSV").
    """
    soup = BeautifulSoup(corpo_html, "html.parser")
    for link in soup.find_all("a"):
        texto = link.get_text(strip=True).lower()
        # Padrão PT: "Baixar CSV da transação"
        if "baixar" in texto and "csv" in texto:
            return link.get("href")
        # Padrão EN: "Download File" ou "Download CSV"
        if "download" in texto and ("file" in texto or "csv" in texto):
            return link.get("href")
    return None


def baixar_arquivo(url: str) -> str | None:
    """Faz o download do arquivo a partir da URL e salva localmente.
    Segue redirects (necessário para links tipo click.uber.com)."""
    os.makedirs(DIRETORIO_DOWNLOADS, exist_ok=True)

    response = requests.get(url, timeout=120, allow_redirects=True)

    if response.status_code == 403:
        print("ERRO 403: O link de download expirou (validade de ~5 minutos).")
        print("Gere um novo relatório no Uber for Business e execute o script logo após receber o email.")
        return None

    response.raise_for_status()

    # Tenta extrair o nome do arquivo do header Content-Disposition
    nome_arquivo = None
    cd = response.headers.get("Content-Disposition", "")
    if "filename=" in cd:
        match = re.search(r'filename="?([^"]+?)"?(;|$)', cd)
        if match:
            nome_arquivo = match.group(1)

    # Fallback: extrai da URL final (após redirect)
    if not nome_arquivo:
        url_final = response.url.split("?")[0]
        candidato = url_final.split("/")[-1]
        if candidato.lower().endswith(".csv"):
            nome_arquivo = candidato

    # Último fallback: nome com timestamp
    if not nome_arquivo:
        nome_arquivo = f"uber_export_{datetime.now(tz=FUSO_BRASILIA).strftime('%Y%m%d_%H%M%S')}.csv"

    caminho_destino = os.path.join(DIRETORIO_DOWNLOADS, nome_arquivo)

    with open(caminho_destino, "wb") as f:
        f.write(response.content)

    return caminho_destino


def obter_token_graph(client_id: str, client_secret: str, tenant_id: str) -> str:
    """Obtém um token de acesso para a Microsoft Graph API."""
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    resultado = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" not in resultado:
        raise Exception(f"Falha ao obter token Graph: {resultado.get('error_description', resultado)}")

    token = resultado["access_token"]

    # Diagnóstico: mostra as permissões presentes no token
    partes = token.split(".")
    payload = json.loads(base64.urlsafe_b64decode(partes[1] + "=="))
    roles = payload.get("roles", [])
    print(f"Permissões no token: {roles if roles else 'NENHUMA — adicione permissões no Azure AD'}")

    return token


def email_do_path_sharepoint(user_path: str, dominio: str) -> str:
    """
    Converte o user_path do SharePoint em email.
    Ex: 'marlon_almeida_autvix_com_br' + 'autvix.com.br' → 'marlon.almeida@autvix.com.br'
    """
    sufixo = "_" + dominio.replace(".", "_")
    if not user_path.endswith(sufixo):
        raise Exception(f"Não foi possível inferir email a partir de '{user_path}' (domínio esperado: {dominio})")
    nome = user_path[:-len(sufixo)].replace("_", ".")
    return f"{nome}@{dominio}"


def resolver_url_sharepoint(sharepoint_url: str, token: str, dominio_email: str) -> dict:
    """Resolve a URL do SharePoint para obter drive_id e item_id.
    Suporta dois formatos:
      1) URL de compartilhamento (com '/:x:/g/personal/.../IQ...')
      2) URL direta no formato 'doc2.aspx?sourcedoc=...&file=...'
    """
    headers = {"Authorization": f"Bearer {token}"}

    # Caso 1: URL no formato doc2.aspx (acesso direto ao OneDrive de outro usuário)
    if "doc2.aspx" in sharepoint_url and "/personal/" in sharepoint_url:
        parsed = urlparse(sharepoint_url)
        match = re.search(r"/personal/([^/]+)", parsed.path)
        if not match:
            raise Exception("Não foi possível extrair o owner da URL.")
        owner_path = match.group(1)
        owner_email = email_do_path_sharepoint(owner_path, dominio_email)

        params = parse_qs(parsed.query)
        if "file" not in params:
            raise Exception("Parâmetro 'file' não encontrado na URL.")
        file_name = unquote(params["file"][0])

        print(f"  - Owner: {owner_email}")
        print(f"  - Arquivo: {file_name}")

        url = f"{GRAPH_API_BASE}/users/{owner_email}/drive/root:/{file_name}"
        resp = requests.get(url, headers=headers, timeout=30)
        if resp.status_code == 401:
            print("ERRO 401: Verifique permissões Files.ReadWrite.All e Sites.ReadWrite.All no Azure AD.")
        resp.raise_for_status()
        item = resp.json()
        print(f"  - Item ID: {item['id']}")
        print(f"  - Web URL: {item.get('webUrl', '?')}")
        print(f"  - Última modificação: {item.get('lastModifiedDateTime', '?')}")
        return {
            "drive_id": item["parentReference"]["driveId"],
            "item_id": item["id"],
        }

    # Caso 2: URL de compartilhamento padrão
    url_limpa = sharepoint_url.split("?")[0]
    encoded = base64.urlsafe_b64encode(url_limpa.encode()).decode().rstrip("=")
    share_id = f"u!{encoded}"
    resp = requests.get(f"{GRAPH_API_BASE}/shares/{share_id}/driveItem", headers=headers, timeout=30)

    if resp.status_code == 401:
        print("ERRO 401: Verifique permissões Files.ReadWrite.All e Sites.ReadWrite.All no Azure AD.")
    resp.raise_for_status()

    item = resp.json()
    print(f"  - Nome do arquivo: {item.get('name', '?')}")
    print(f"  - Item ID: {item['id']}")
    print(f"  - Drive ID: {item['parentReference'].get('driveId', '?')}")
    print(f"  - Web URL: {item.get('webUrl', '?')}")
    print(f"  - Última modificação: {item.get('lastModifiedDateTime', '?')}")
    print(f"  - Modificado por: {item.get('lastModifiedBy', {}).get('user', {}).get('displayName', '?')}")
    return {
        "drive_id": item["parentReference"]["driveId"],
        "item_id": item["id"],
    }


def parse_csv_uber(caminho_csv: str) -> tuple[list[str], list[list[str]]]:
    """
    Faz o parse do CSV exportado pela Uber, ignorando as linhas iniciais
    de cabeçalho do relatório (Empresa, Administrador, Data, etc.).
    Retorna (cabeçalho, linhas_de_dados).
    """
    with open(caminho_csv, "r", encoding="utf-8") as f:
        leitor = csv.reader(f, delimiter=";")
        linhas = list(leitor)

    # Procura a linha de cabeçalho: contém "ID da viagem" ou começa com "Data da solicitação"
    indice_header = None
    for i, linha in enumerate(linhas):
        if not linha:
            continue
        linha_str = ";".join(linha).lower()
        if "id da viagem" in linha_str and "data da solicitação" in linha_str:
            indice_header = i
            break

    if indice_header is None:
        raise Exception("Cabeçalho de transações não encontrado no CSV.")

    cabecalho = linhas[indice_header]
    dados = [l for l in linhas[indice_header + 1:] if any(c.strip() for c in l)]
    return cabecalho, dados


def transformar_para_excel(cabecalho: list[str], dados: list[list[str]]) -> list[list[str]]:
    """Transforma as linhas do CSV no formato esperado pela planilha online (13 colunas)."""

    def idx(*nomes: str) -> int:
        """Retorna o índice da primeira coluna encontrada (na ordem fornecida)."""
        for nome in nomes:
            for i, h in enumerate(cabecalho):
                if h.strip() == nome:
                    return i
        return -1

    def get(linha: list[str], i: int) -> str:
        if i < 0 or i >= len(linha):
            return ""
        return linha[i].strip()

    i_nome = idx("Nome")
    i_sobrenome = idx("Sobrenome")
    i_funcionario = idx("ID do funcionário")
    i_data = idx("Data da solicitação (local)")
    i_hora = idx("Hora da solicitação (local)")
    i_codigo = idx("Código da despesa")
    i_detalhe = idx("Detalhamento da despesa")
    i_valor = idx("Valor da transação: BRL", "Valor total: BRL", "Valor da transação (moeda local)", "Valor total (moeda local)")
    i_partida = idx("Endereço de partida")
    i_destino = idx("Endereço de destino")

    resultado = []
    for linha in dados:
        data_raw = get(linha, i_data)

        # Pula linhas de subtotal/separador do CSV (ex: "--;;;;--;--;...")
        if data_raw == "--" or not data_raw:
            continue

        # Formato 1 (Nome + Sobrenome) ou Formato 2 (ID do funcionário com nome completo)
        if i_nome >= 0 and i_sobrenome >= 0:
            colaborador = f"{get(linha, i_nome)} {get(linha, i_sobrenome)}".strip()
        else:
            colaborador = get(linha, i_funcionario)

        mes, ano, data_formatada = "", "", ""
        try:
            dt = datetime.strptime(data_raw, "%m/%d/%Y")
            mes = MESES_PT.get(dt.month, str(dt.month))
            ano = str(dt.year)
            data_formatada = dt.strftime("%d/%m/%Y")
        except ValueError:
            # Data em formato inesperado: pula a linha (não é registro válido)
            continue

        resultado.append([
            colaborador,                  # COLABORADOR
            "",                           # SETOR
            mes,                          # MES
            ano,                          # ANO
            data_formatada,               # DATA SOLICITAÇÃO
            get(linha, i_hora),           # HORA SOLICITAÇÃO
            get(linha, i_codigo).upper(),  # CENTRO DE CUSTO (sempre em maiúsculas)
            get(linha, i_detalhe),        # DETALHAMENTO DA VIAGEM
            get(linha, i_valor),          # VALOR DA VIAGEM
            "",                           # REVISOR
            "",                           # VIAGEM APROVADA?
            get(linha, i_partida),        # Endereço de partida
            get(linha, i_destino),        # Endereço de destino
        ])

    return resultado


def obter_nome_aba(drive_id: str, item_id: str, token: str) -> str:
    """Detecta automaticamente o nome da primeira aba da planilha."""
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}/workbook/worksheets"
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()

    abas = resp.json().get("value", [])
    if not abas:
        raise Exception("Nenhuma aba encontrada na planilha.")

    nomes = [a["name"] for a in abas]
    print(f"Abas existentes na planilha: {nomes}")
    nome = abas[0]["name"]
    print(f"Aba selecionada para inserção: '{nome}'")
    return nome


def obter_ultima_linha(drive_id: str, item_id: str, nome_aba: str, token: str) -> int:
    """
    Retorna o número da última linha com dados REAIS na coluna A (COLABORADOR).
    Ignora linhas com formatação mas sem conteúdo, percorrendo de baixo para cima.
    """
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{nome_aba}')/usedRange(valuesOnly=true)"
    )
    resp = requests.get(url, headers=headers, timeout=60)
    resp.raise_for_status()

    data = resp.json()
    address = data.get("address", "")
    values = data.get("values", [])
    print(f"  - Endereço retornado pela API: '{address}'")
    print(f"  - Valores recebidos: {len(values)} linhas")

    # DEBUG: imprime as 3 PRIMEIRAS linhas (para confirmar que é o arquivo certo)
    print("  - DEBUG: PRIMEIRAS 3 linhas da planilha (compare com o seu navegador):")
    for i in range(min(3, len(values))):
        valores_linha = values[i] if i < len(values) else []
        print(f"      Linha {i + 1}: {[repr(v) for v in valores_linha]}")

    # Determina a linha inicial do usedRange (ex: "Sheet1!A1:M4566" → 1)
    range_part = address.split("!")[-1]
    start_cell = range_part.split(":")[0]
    match_inicio = re.search(r"\d+", start_cell)
    linha_inicial = int(match_inicio.group()) if match_inicio else 1

    # DEBUG: imprime as 5 últimas células de cada coluna (A-M) para diagnóstico
    print("  - DEBUG: últimas 5 linhas (todas as colunas):")
    for i in range(max(0, len(values) - 5), len(values)):
        linha_excel = linha_inicial + i
        valores_linha = values[i] if i < len(values) else []
        print(f"      Linha {linha_excel}: {[repr(v) for v in valores_linha]}")

    # Percorre de baixo para cima procurando a primeira linha com conteúdo REAL
    # em qualquer coluna (não só A, pois A pode ter strings vazias ocultas)
    for i in range(len(values) - 1, -1, -1):
        for celula in values[i]:
            if celula is None:
                continue
            texto = str(celula).strip()
            if texto:
                return linha_inicial + i

    return linha_inicial


def enviar_para_planilha(caminho_csv: str, token: str):
    """
    Lê o CSV da Uber, transforma para o formato da planilha online
    e ANEXA os novos registros após a última linha existente.
    """
    print("Resolvendo URL da planilha no SharePoint...")
    dominio = EMAIL.split("@")[1] if EMAIL else "autvix.com.br"
    info = resolver_url_sharepoint(SHAREPOINT_URL, token, dominio)
    drive_id = info["drive_id"]
    item_id = info["item_id"]

    nome_aba = obter_nome_aba(drive_id, item_id, token)

    print("Lendo e transformando dados do CSV...")
    cabecalho, dados_csv = parse_csv_uber(caminho_csv)
    print(f"  - {len(cabecalho)} colunas no CSV")
    print(f"  - {len(dados_csv)} linhas de dados encontradas")

    dados_transformados = transformar_para_excel(cabecalho, dados_csv)
    if not dados_transformados:
        print("Nenhuma linha de dado para enviar.")
        return

    # DEBUG: salva os dados transformados em JSON para inspeção
    debug_payload = os.path.join(os.path.dirname(__file__), "debug_payload.json")
    with open(debug_payload, "w", encoding="utf-8") as f:
        json.dump(
            {"total_linhas": len(dados_transformados), "values": dados_transformados},
            f, ensure_ascii=False, indent=2,
        )
    print(f"  - Payload de debug salvo em: {debug_payload}")
    print(f"  - Total de registros a enviar: {len(dados_transformados)}")

    print("Identificando última linha da planilha online...")
    ultima_linha = obter_ultima_linha(drive_id, item_id, nome_aba, token)
    print(f"  - Última linha com dados: {ultima_linha}")

    linha_inicial = ultima_linha + 1
    linha_final = linha_inicial + len(dados_transformados) - 1
    intervalo = f"A{linha_inicial}:M{linha_final}"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    # Limpa qualquer formatação direta pré-existente nas células-alvo
    # (evita texto invisível por fundo preto, fonte minúscula, etc.)
    print(f"Limpando formatação pré-existente em {intervalo}...")
    url_clear = (
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{nome_aba}')/range(address='{intervalo}')/clear"
    )
    requests.post(url_clear, headers=headers, json={"applyTo": "Formats"}, timeout=60)

    print(f"Anexando {len(dados_transformados)} linhas no intervalo {intervalo}...")

    url_range = (
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{nome_aba}')/range(address='{intervalo}')"
    )
    payload = {"values": dados_transformados}
    resp = requests.patch(url_range, headers=headers, json=payload, timeout=120)

    if not resp.ok:
        print(f"ERRO HTTP {resp.status_code} ao enviar dados:")
        print(resp.text[:2000])
        resp.raise_for_status()

    # Força altura de linha padrão e desoculta linhas (corrige "linha preta"
    # causada por rowHeight=0 ou rowHidden=true em formato pré-existente)
    print("Normalizando altura de linha e visibilidade das células inseridas...")
    url_format = (
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{nome_aba}')/range(address='{intervalo}')/format"
    )
    formato_payload = {"rowHidden": False, "rowHeight": 15}
    resp_fmt = requests.patch(url_format, headers=headers, json=formato_payload, timeout=60)
    if not resp_fmt.ok:
        print(f"  AVISO: ajuste de altura/visibilidade falhou ({resp_fmt.status_code}): {resp_fmt.text[:300]}")

    # Remove qualquer fill/cor de fundo explícito
    url_fill = (
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{nome_aba}')/range(address='{intervalo}')/format/fill/clear"
    )
    requests.post(url_fill, headers=headers, timeout=30)

    print("Dados anexados com sucesso à planilha online!")


def main():
    variaveis_faltando = []
    if not EMAIL:
        variaveis_faltando.append("OUTLOOK_EMAIL")
    if not CLIENT_ID:
        variaveis_faltando.append("AZURE_CLIENT_ID")
    if not CLIENT_SECRET:
        variaveis_faltando.append("AZURE_CLIENT_SECRET")
    if not TENANT_ID:
        variaveis_faltando.append("AZURE_TENANT_ID")

    if variaveis_faltando:
        print(f"ERRO: Defina as variáveis no arquivo .env: {', '.join(variaveis_faltando)}")
        return None

    print(f"[{datetime.now(tz=FUSO_BRASILIA).strftime('%d/%m/%Y %H:%M:%S')}] Conectando ao Outlook via OAuth2...")
    conta = conectar_outlook(EMAIL, CLIENT_ID, CLIENT_SECRET, TENANT_ID)

    print("Buscando o email mais recente com os filtros configurados (últimas 24h)...")
    email_encontrado = buscar_email_mais_recente(conta)

    if not email_encontrado:
        print("Nenhum email encontrado com os critérios especificados.")
        return None

    print(f"Email encontrado: {email_encontrado['data_recebimento']} - {email_encontrado['assunto']}")

    # Salva o HTML do email para inspeção/debug
    caminho_html_debug = os.path.join(os.path.dirname(__file__), "ultimo_email.html")
    with open(caminho_html_debug, "w", encoding="utf-8") as f:
        f.write(email_encontrado["corpo"])
    print(f"HTML do email salvo em: {caminho_html_debug}")

    link = extrair_link_download(email_encontrado["corpo"])
    if not link:
        print("ERRO: Link de download não encontrado no corpo do email.")
        print(f"Inspecione o arquivo '{caminho_html_debug}' para ver a estrutura do email.")
        return email_encontrado

    print("Link de download extraído. Baixando arquivo...")
    caminho = baixar_arquivo(link)

    if not caminho:
        return email_encontrado

    print(f"Arquivo salvo em: {caminho}")

    email_encontrado["link_download"] = link
    email_encontrado["arquivo_local"] = caminho

    # Enviar para planilha online
    if not SHAREPOINT_URL:
        print("AVISO: SHAREPOINT_URL não definida no .env. Pulando envio para planilha.")
        return email_encontrado

    print("\nObtendo token para Microsoft Graph API...")
    token_graph = obter_token_graph(CLIENT_ID, CLIENT_SECRET, TENANT_ID)
    enviar_para_planilha(caminho, token_graph)

    return email_encontrado


if __name__ == "__main__":
    email_filtrado = main()
