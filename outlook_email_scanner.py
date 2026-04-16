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
import os
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

FILTRO_REMETENTE = "rodrigo.malavasi@autvix.com.br"
FILTRO_ASSUNTO = "Your file export is ready"
FILTRO_BODY = "Your file export is ready"

FUSO_BRASILIA = ZoneInfo("America/Sao_Paulo")
DIRETORIO_DOWNLOADS = os.path.join(os.path.dirname(__file__), "downloads")
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"


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
    que atenda aos critérios de filtro.

    Retorna um dicionário com os dados do email ou None se não encontrar.
    """
    agora = datetime.now(tz=FUSO_BRASILIA)
    inicio = agora - timedelta(hours=24)

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
    """Extrai o link de 'Download File' do corpo HTML do email."""
    soup = BeautifulSoup(corpo_html, "html.parser")
    for link in soup.find_all("a"):
        texto = link.get_text(strip=True)
        if "download" in texto.lower() and "file" in texto.lower():
            return link.get("href")
    return None


def baixar_arquivo(url: str) -> str | None:
    """Faz o download do arquivo a partir da URL e salva localmente."""
    os.makedirs(DIRETORIO_DOWNLOADS, exist_ok=True)

    response = requests.get(url, timeout=120)

    if response.status_code == 403:
        print("ERRO 403: O link de download expirou (validade de ~5 minutos).")
        print("Gere um novo relatório no Uber for Business e execute o script logo após receber o email.")
        return None

    response.raise_for_status()

    # Extrai o nome do arquivo da URL (antes dos query params)
    caminho_url = url.split("?")[0]
    nome_arquivo = caminho_url.split("/")[-1]

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


def resolver_url_compartilhamento(sharepoint_url: str, token: str) -> dict:
    """Resolve uma URL de compartilhamento do SharePoint para obter drive_id e item_id."""
    # Remove query params da URL (ex: ?e=NA1fmo) antes de codificar
    url_limpa = sharepoint_url.split("?")[0]
    encoded = base64.urlsafe_b64encode(url_limpa.encode()).decode().rstrip("=")
    share_id = f"u!{encoded}"

    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(f"{GRAPH_API_BASE}/shares/{share_id}/driveItem", headers=headers, timeout=30)

    if resp.status_code == 401:
        print("ERRO 401: Verifique no Azure AD se as permissões abaixo estão com 'Admin consent granted':")
        print("  - Microsoft Graph → Files.ReadWrite.All (Application)")
        print("  - Microsoft Graph → Sites.ReadWrite.All (Application)")
    resp.raise_for_status()

    item = resp.json()
    return {
        "drive_id": item["parentReference"]["driveId"],
        "item_id": item["id"],
    }


def ler_csv(caminho_csv: str) -> list[list[str]]:
    """Lê o arquivo CSV e retorna como lista de listas (linhas x colunas)."""
    linhas = []
    with open(caminho_csv, "r", encoding="utf-8") as f:
        leitor = csv.reader(f)
        for linha in leitor:
            linhas.append(linha)
    return linhas


def obter_nome_aba(drive_id: str, item_id: str, token: str) -> str:
    """Detecta automaticamente o nome da primeira aba da planilha."""
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}/workbook/worksheets"
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()

    abas = resp.json().get("value", [])
    if not abas:
        raise Exception("Nenhuma aba encontrada na planilha.")

    nome = abas[0]["name"]
    print(f"Aba detectada: '{nome}'")
    return nome


def enviar_para_planilha(caminho_csv: str, token: str):
    """Envia os dados do CSV para a planilha online no SharePoint."""
    print("Resolvendo URL da planilha no SharePoint...")
    info = resolver_url_compartilhamento(SHAREPOINT_URL, token)
    drive_id = info["drive_id"]
    item_id = info["item_id"]

    nome_aba = obter_nome_aba(drive_id, item_id, token)

    print("Lendo arquivo CSV...")
    dados = ler_csv(caminho_csv)

    if not dados:
        print("ERRO: CSV vazio, nada para enviar.")
        return

    num_linhas = len(dados)
    num_colunas = max(len(linha) for linha in dados)

    # Padroniza todas as linhas para ter o mesmo número de colunas
    for linha in dados:
        while len(linha) < num_colunas:
            linha.append("")

    # Converte índice de coluna para letra (0=A, 1=B, ..., 25=Z, 26=AA, ...)
    def col_letra(n):
        resultado = ""
        while n >= 0:
            resultado = chr(n % 26 + ord("A")) + resultado
            n = n // 26 - 1
        return resultado

    ultima_col = col_letra(num_colunas - 1)
    intervalo = f"A1:{ultima_col}{num_linhas}"

    print(f"Enviando {num_linhas} linhas x {num_colunas} colunas (intervalo: {intervalo})...")

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    # Limpa a planilha antes de inserir novos dados
    url_clear = (
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{nome_aba}')/range(address='A:ZZ')/clear"
    )
    requests.post(url_clear, headers=headers, json={"applyTo": "All"}, timeout=60)

    # Envia os dados para o intervalo
    url_range = (
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{nome_aba}')/range(address='{intervalo}')"
    )
    payload = {"values": dados}
    resp = requests.patch(url_range, headers=headers, json=payload, timeout=120)
    resp.raise_for_status()

    print(f"Dados enviados com sucesso para a planilha online!")


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

    link = extrair_link_download(email_encontrado["corpo"])
    if not link:
        print("ERRO: Link de download não encontrado no corpo do email.")
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
