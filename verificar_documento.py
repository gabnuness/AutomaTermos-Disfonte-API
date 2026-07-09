import requests
import json
import os

url = "https://api.autentique.com.br/v2/graphql"
token = os.environ.get("AUTENTIQUE_TOKEN")

# Cole aqui o ID que apareceu no seu último teste
DOCUMENT_ID = "ff92d6874d4032de058163de699cf17cfb994f84f7ac61c14"

query = """
query VerificarDocumento($id: UUID!) {
    document(id: $id) {
        id
        name
        signatures {
            public_id
            positions {
                element
                x
                y
                z
            }
            link {
                short_link
            }
        }
    }
}
"""

payload = {
    "query": query,
    "variables": {"id": DOCUMENT_ID}
}
headers = {"Authorization": f"Bearer {token}"}

response = requests.post(url, json=payload, headers=headers)
print(json.dumps(response.json(), indent=2, ensure_ascii=False))