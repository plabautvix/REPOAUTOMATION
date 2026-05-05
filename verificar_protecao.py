"""
Verifica o status de proteção da aba 'Controle de Uber' na planilha online.
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
import json


NOME_ABA = "Controle de Uber"


def main():
    print("Obtendo token Graph...")
    token = obter_token_graph(CLIENT_ID, CLIENT_SECRET, TENANT_ID)

    print("Resolvendo URL da planilha...")
    dominio = EMAIL.split("@")[1] if EMAIL else "autvix.com.br"
    info = resolver_url_sharepoint(SHAREPOINT_URL, token, dominio)
    drive_id = info["drive_id"]
    item_id = info["item_id"]

    headers = {"Authorization": f"Bearer {token}"}

    # 1. Status de proteção da aba
    print(f"\n=== Verificando proteção da aba '{NOME_ABA}' ===")
    url = (
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{NOME_ABA}')/protection"
    )
    resp = requests.get(url, headers=headers, timeout=30)
    if resp.ok:
        protection = resp.json()
        print(json.dumps(protection, indent=2, ensure_ascii=False))
    else:
        print(f"Erro ao consultar protection: {resp.status_code} - {resp.text[:300]}")

    # 2. Verifica tabelas existentes na aba (uma tabela limitada pode ser a causa)
    print(f"\n=== Verificando tabelas (Excel Tables) na aba '{NOME_ABA}' ===")
    url_tables = (
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{NOME_ABA}')/tables"
    )
    resp = requests.get(url_tables, headers=headers, timeout=30)
    if resp.ok:
        tables = resp.json().get("value", [])
        if not tables:
            print("Nenhuma tabela formatada encontrada.")
        for t in tables:
            print(f"  - Nome: {t.get('name')}")
            print(f"    Range: {t.get('range', {}).get('address', '?')}")
            # Pega o range usando endpoint específico
            url_range = (
                f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
                f"/workbook/tables/{t['id']}/range"
            )
            r2 = requests.get(url_range, headers=headers, timeout=30)
            if r2.ok:
                addr = r2.json().get("address", "?")
                print(f"    Endereço da tabela: {addr}")
    else:
        print(f"Erro ao consultar tabelas: {resp.status_code} - {resp.text[:300]}")

    # 3. Verifica intervalos nomeados
    print(f"\n=== Verificando intervalos nomeados (Named Ranges) ===")
    url_names = (
        f"{GRAPH_API_BASE}/drives/{drive_id}/items/{item_id}"
        f"/workbook/names"
    )
    resp = requests.get(url_names, headers=headers, timeout=30)
    if resp.ok:
        names = resp.json().get("value", [])
        if not names:
            print("Nenhum intervalo nomeado.")
        for n in names:
            print(f"  - {n.get('name')}: {n.get('value', '?')[:120]}")


if __name__ == "__main__":
    main()
