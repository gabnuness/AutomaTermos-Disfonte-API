import requests
import json
import os

url = "https://api.autentique.com.br/v2/graphql"
token = os.environ.get("AUTENTIQUE_TOKEN")

# ==========================
# SÓ MEXA NESSAS 5 LINHAS
# ==========================
CAMINHO_PDF = r"P:\Publico\TI Suporte Técnico\TERMOS AUTOMATICOS\Termo_TESTE.pdf"  # um PDF que você já gerou antes
TELEFONE = "54984481615"  # seu telefone, sem formatação
CPF = "03718084058"       # seu CPF, sem pontuação
X = 39.1
Y = 18.5
# ==========================

query = """
mutation CreateDocumentMutation($document: DocumentInput!, $signers: [SignerInput!]!, $file: Upload!) {
    createDocument(sandbox: true, document: $document, signers: $signers, file: $file) {
        id
        name
    }
}
"""

variables = {
    "document": {"name": f"TESTE_x{X}_y{Y}"},
    "signers": [{
        "phone": "+55" + TELEFONE,
        "delivery_method": "DELIVERY_METHOD_WHATSAPP",
        "action": "SIGN",
        "configs": {"cpf": CPF},
        #"security_verifications": [{"type": "UPLOAD"}],
        "positions": [{"x": str(X), "y": str(Y), "z": 3, "element": "SIGNATURE"}]
    }]
}

payload = {
    "operations": json.dumps({"query": query, "variables": variables}),
    "map": json.dumps({"file": ["variables.file"]})
}
headers = {"Authorization": f"Bearer {token}"}

with open(CAMINHO_PDF, "rb") as f:
    response = requests.post(url, data=payload, files={"file": f}, headers=headers)

print(response.json())