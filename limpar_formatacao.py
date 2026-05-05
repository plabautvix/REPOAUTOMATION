"""
Script utilitário (uso único): remove TODA a formatação condicional da aba
'Controle de Uber' da planilha online no SharePoint via Microsoft Graph API.

Use quando suspeitar que regras de formatação condicional estão deixando
células ilegíveis (ex: fundo preto + texto preto).
"""

from outlook_email_scanner import (
    obter_token_graph,
    resolver_url_sharepoint,
    GRAPH_API_BASE,
    SHAREPOINT_URL,
    CLIENT_ID,
    CLIENT_SECRET,
    TENANT_ID,
    EMAIL,
)
import requests


NOME_ABA = "Controle de Uber"


def listar_formatacoes_condicionais(drive_id, item_id, nome_aba, token):
    """Lista todas as regras de formatação condicional da aba."""
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{nome_aba}')/conditionalFormats"
    )
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json().get("value", [])


def remover_formatacao(drive_id, item_id, nome_aba, format_id, token):
    """Remove uma regra de formatação condicional específica."""
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{nome_aba}')/conditionalFormats/{format_id}"
    )
    resp = requests.delete(url, headers=headers, timeout=30)
    resp.raise_for_status()


def main():
    print("Obtendo token Graph...")
    token = obter_token_graph(CLIENT_ID, CLIENT_SECRET, TENANT_ID)

    print("Resolvendo URL da planilha...")
    dominio = EMAIL.split("@")[1] if EMAIL else "autvix.com.br"
    info = resolver_url_sharepoint(SHAREPOINT_URL, token, dominio)
    drive_id = info["drive_id"]
    item_id = info["item_id"]

    print(f"Listando formatações condicionais na aba '{NOME_ABA}'...")
    formatacoes = listar_formatacoes_condicionais(drive_id, item_id, NOME_ABA, token)
    print(f"  - {len(formatacoes)} regras encontradas")

    if not formatacoes:
        print("Nenhuma regra para remover. Nada a fazer.")
        return

    for i, fmt in enumerate(formatacoes, 1):
        fmt_id = fmt.get("id", "?")
        fmt_type = fmt.get("type", "?")
        print(f"  [{i}/{len(formatacoes)}] Removendo regra (id={fmt_id}, tipo={fmt_type})...")
        remover_formatacao(drive_id, item_id, NOME_ABA, fmt_id, token)

    print(f"\nTodas as {len(formatacoes)} regras foram removidas com sucesso.")


if __name__ == "__main__":
    main()
