from docxtpl import DocxTemplate
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import os
from docx2pdf import convert  # Nova biblioteca para conversão
import requests
import json


# ==========================
# ENVIO PARA AUTENTIQUE
# ==========================
def enviar_para_autentique(caminho_pdf, nome_funcionario, cpf_funcionario, telefone_funcionario):
    """
    Envia o PDF do termo para o Autentique, já com telefone, CPF,
    validação por documento e posição de assinatura configurados.
    """
    url = "https://api.autentique.com.br/v2/graphql"
    
    token = os.environ.get("AUTENTIQUE_TOKEN")  # <<< cole seu token gerado aqui

    if not token:
        print("\n[ERRO] Variável de ambiente AUTENTIQUE_TOKEN não encontrada.")
        print("Configure com: setx AUTENTIQUE_TOKEN \"seu_token\" e rabra o terminal.")
        exit()

    cpf_limpo = cpf_funcionario.replace(".", "").replace("-", "")
    telefone_limpo = "+55" + telefone_funcionario.replace("(", "").replace(")", "").replace(" ", "").replace("-", "")

    query = """
    mutation CreateDocumentMutation(
        $document: DocumentInput!,
        $signers: [SignerInput!]!,
        $file: Upload!
    ) {
        createDocument(sandbox: true, document: $document, signers: $signers, file: $file) {
            id
            name
            signatures {
                public_id
                name
                email
                link { short_link }
            }
        }
    }
    """

    variables = {
        "document": {"name": f"Termo_{nome_funcionario}"},
        "signers": [
            {
                "phone": telefone_limpo,
                "delivery_method": "DELIVERY_METHOD_WHATSAPP",
                "action": "SIGN",
                "configs": {"cpf": cpf_limpo},
                "security_verifications": [{"type": "UPLOAD"}],
                "positions": [
                    {"x": "31.1", "y": "18.5", "z": 3, "element": "SIGNATURE"}
                ]
            }
        ]
    }

    # Monta o JSON corretamente, sem gambiarra de string
    operations = json.dumps({"query": query, "variables": variables})

    payload = {
        "operations": operations,
        "map": json.dumps({"file": ["variables.file"]})
    }

    headers = {"Authorization": f"Bearer {token}"}

    with open(caminho_pdf, "rb") as f:
        files = {"file": f}
        response = requests.post(url, data=payload, files=files, headers=headers)

    resultado = response.json()

    if "errors" in resultado:
        print(f"\n[ERRO Autentique] {resultado['errors']}")
    else:
        print(f"\n[OK] Documento enviado para assinatura no Autentique!")
        doc = resultado["data"]["createDocument"]
        print(f"ID do documento: {doc['id']}")

    return resultado

# ==========================
# CONFIGURAÇÃO DO TKINTER (Janela oculta)
# ==========================
root = tk.Tk()
root.withdraw()  # Oculta a janela principal do tkinter

# ==========================
# DATA ATUAL
# ==========================
hoje = datetime.now()

meses = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
}

dia_atual = f"{hoje.day:02d}"
mes_atual = meses[hoje.month]
ano_atual = str(hoje.year)
data_completa = f"{dia_atual} de {mes_atual} de {ano_atual}"

# ==========================
# DADOS (INPUT)
# ==========================
nome = input("Nome: ").upper()
cargo = input("Cargo: ").upper()
setor = input("Área: ").upper()
cidade = input("Cidade: ").upper()
cpf = input("CPF: ")
modelo = input("Modelo: ").upper()
imei = input("IMEI: ")
numero = input("Número: ")
valor = float(input("Valor: "))

# ==========================
# FORMATAÇÃO DO VALOR
# ==========================
valor_formatado = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ==========================
# PARCELAMENTO
# ==========================
if valor <= 200:
    parcelas = 3
elif valor <= 400:
    parcelas = 5
else:
    parcelas = 8

valor_parcela = valor / parcelas
valor_parcela = f"R$ {valor_parcela:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ==========================
# ABRIR DOCUMENTO
# ==========================
try:
    doc = DocxTemplate("modelo.docx")
except Exception as e:
    print(f"Erro ao abrir 'modelo.docx'. Verifique se o arquivo está na mesma pasta do script. Erro: {e}")
    exit()

contexto = {
    "nome": nome,
    "cargo": cargo,
    "setor": setor,
    "cidade": cidade,
    "cpf": cpf,
    "modelo": modelo,
    "imei": imei,
    "numero": numero,
    "valor": valor_formatado,
    "data": data_completa,

    # Variáveis específicas para o final do modelo.docx
    "dia": dia_atual,
    "mês": mes_atual,
    "ano": ano_atual,

    # Segunda página (Tabela)
    "parcelas": parcelas,
    "valor_parcela": valor_parcela
}

doc.render(contexto)

# ==========================
# SELECIONAR CAMINHO PARA SALVAR
# ==========================
print("\nEscolha onde deseja salvar o Termo de Responsabilidade...")

# Abre a janela de diálogo
caminho_docx = filedialog.asksaveasfilename(
    initialfile=f"Termo_{nome.replace(' ', '_')}.docx",
    defaultextension=".docx",
    filetypes=[("Documentos Word", "*.docx")],
    title="Salvar Termo de Responsabilidade"
)

if caminho_docx:
    try:
        # 1. Salva o arquivo Word temporário/final solicitado pelo usuário
        doc.save(caminho_docx)
        print(f"\n[OK] Documento Word gerado em:\n{caminho_docx}")
        
        # 2. Define o caminho do PDF mudando a extensão de .docx para .pdf
        caminho_pdf = os.path.splitext(caminho_docx)[0] + ".pdf"
        
        print("Convertendo para PDF, por favor aguarde...")
        # 3. Faz a mágica da conversão utilizando o Word em background
        convert(caminho_docx, caminho_pdf)
        print(f"[OK] PDF gerado com sucesso em:\n{caminho_pdf}")

        
        enviar_para_autentique(caminho_pdf, nome, cpf, numero)
        
        # Opcional: Se você NÃO quiser guardar o arquivo .docx e quiser APENAS o PDF,
        # descomente a linha abaixo para apagar o Word depois que o PDF for criado:
        # os.remove(caminho_docx)

    except PermissionError:
        print("\n[ERRO] Não foi possível salvar. O arquivo já está aberto por outro programa (Word ou leitor de PDF).")
    except Exception as e:
        print(f"\n[ERRO] Ocorreu um problema: {e}")
else:
    print("\nOperação cancelada pelo usuário.")