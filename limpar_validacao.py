"""
Remove validação de dados de toda a aba 'Controle de Uber'.
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
# Range a limpar: deixa o cabeçalho intocado (linha 1) e limpa de A2 até o fim
RANGE_LIMPEZA = "A2:M1048576"


def main():
    print("Obtendo token Graph...")
    token = obter_token_graph(CLIENT_ID, CLIENT_SECRET, TENANT_ID)

    print("Resolvendo URL da planilha...")
    dominio = EMAIL.split("@")[1] if EMAIL else "autvix.com.br"
    info = resolver_url_sharepoint(SHAREPOINT_URL, token, dominio)
    drive_id = info["drive_id"]
    item_id = info["item_id"]

    headers = {"Authorization": f"Bearer {token}"}
    GRAPH_BETA = "https://graph.microsoft.com/beta"

    # Limpa a validação de dados no range (endpoint disponível apenas em /beta)
    print(f"\nRemovendo validação de dados em {RANGE_LIMPEZA}...")
    url_clear = (
        f"{GRAPH_BETA}/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{NOME_ABA}')/range(address='{RANGE_LIMPEZA}')"
        f"/dataValidation/clear"
    )
    resp = requests.post(url_clear, headers=headers, timeout=60)
    if resp.ok:
        print("Validação de dados removida com sucesso!")
    else:
        print(f"ERRO ao limpar via /beta: {resp.status_code}: {resp.text[:300]}")
        print("\nTentando abordagem alternativa: aplicar validação 'permitir tudo'...")

        # Alternativa: sobrescrever a validação com uma regra que aceita qualquer valor
        url_override = (
            f"{GRAPH_BETA}/drives/{drive_id}/items/{item_id}"
            f"/workbook/worksheets('{NOME_ABA}')/range(address='{RANGE_LIMPEZA}')"
            f"/dataValidation"
        )
        payload = {
            "rule": None,
            "errorAlert": {"showAlert": False},
            "ignoreBlanks": True,
        }
        resp2 = requests.patch(url_override, headers=headers, json=payload, timeout=60)
        if resp2.ok:
            print("Validação sobrescrita com regra permissiva.")
        else:
            print(f"ERRO {resp2.status_code}: {resp2.text[:500]}")


if __name__ == "__main__":
    main()
