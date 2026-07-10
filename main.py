from docxtpl import DocxTemplate
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import os
from docx2pdf import convert  # Nova biblioteca para conversão
import requests
import json

# ==========================
# FUNÇÃO PARA NORMALIZAR TELEFONE
# ==========================
def normalizar_telefone(numero_digitado, ddd_padrao="54"):
    """
    Aceita o número em qualquer formato e devolve ele padronizado.
    - DDD é opcional (assume 54 se não vier)
    - O 9 inicial é opcional (assume que existe se não vier)
    Retorna: (numero_e164, numero_formatado)
    Exemplo: "984481615" -> ("+5554984481615", "(54) 98448-1615")
    """
    apenas_digitos = "".join(filter(str.isdigit, numero_digitado))

    # Remove o código do país se a pessoa digitou (ex: 55 54 98448-1615)
    if apenas_digitos.startswith("55") and len(apenas_digitos) > 11:
        apenas_digitos = apenas_digitos[2:]

    if len(apenas_digitos) == 8:
        # Só o número, sem DDD e sem o 9 (ex: 84481615)
        apenas_digitos = ddd_padrao + "9" + apenas_digitos
    elif len(apenas_digitos) == 9:
        # Número com o 9, mas sem DDD (ex: 984481615)
        apenas_digitos = ddd_padrao + apenas_digitos
    elif len(apenas_digitos) == 10:
        # DDD + número sem o 9 (ex: 5484481615)
        apenas_digitos = apenas_digitos[:2] + "9" + apenas_digitos[2:]
    elif len(apenas_digitos) == 11:
        # Já está completo (DDD + 9 + número)
        pass
    else:
        raise ValueError(f"Número de telefone inválido: '{numero_digitado}' (ficou com {len(apenas_digitos)} dígitos após limpeza)")

    ddd = apenas_digitos[:2]
    resto = apenas_digitos[2:]  # 9XXXXXXXX
    numero_formatado = f"({ddd}) {resto[:5]}-{resto[5:]}"
    numero_e164 = "+55" + apenas_digitos

    return numero_e164, numero_formatado

# ==========================
# ENVIO PARA AUTENTIQUE
# ==========================
def enviar_para_autentique(caminho_pdf, nome_funcionario, cpf_funcionario, telefone_funcionario, folder_id):
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
    telefone_limpo, _ = normalizar_telefone(telefone_funcionario)

    query = """
    mutation CreateDocumentMutation(
        $document: DocumentInput!,
        $signers: [SignerInput!]!,
        $file: Upload!,
        $folder_id: UUID!
    ) {
        createDocument(
        sandbox: true, 
        document: $document, 
        signers: $signers, 
        file: $file, 
        folder_id: $folder_id
        ) {
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
                #"security_verifications": [{"type": "UPLOAD"}],
                "positions": [
                    {"x": "31.1", "y": "18.5", "z": 3, "element": "SIGNATURE"}
                ]
            }
        ],
        "folder_id": folder_id
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

try:
    telefone_e164, telefone_formatado = normalizar_telefone(numero)
    print(f"Telefone confirmado: {telefone_formatado}")
except ValueError as e:
    print(f"\n[ERRO] {e}")
    exit()

valor = float(input("Valor: "))

# ==========================
# ESCOLHA DO DEPARTAMENTO (pasta no Autentique)
# =========================

pastas_autentique = {
    "1": {
        "nome": "Administrativo",
        "folder_id": "4da8da649f9bf29f06d51421e9e6f92487525e72",
        "subpastas": None
    },
    "2": {
        "nome": "Armazém",
        "folder_id": "585c17bb4b6c7da4d0ce0783e69befc39c6bf163",
        "subpastas": {
            "1": {"nome": "Turno Dia", "folder_id": "c327d88bac0e4bc77cf3eb3308bbad37c8be343f"},
            "2": {"nome": "Turno Noite", "folder_id": "a47d75a1222438f58d2888e6476e801ae14df769"},
        }
    },
    "3": {
        "nome": "Comercial",
        "folder_id": "2f331f102b2a2d04d50937b744db9282c9061c52",
        "subpastas": {
            "1": {"nome": "Representante de Negócios", "folder_id": "d6c1bc409b754aa78cb34f70ab5897f850507a31"},
            "2": {"nome": "Promotor de Vendas", "folder_id": "b81c765e693a968c8c1c8dd6b126b221abcc8fde"},
        }
    },
    "4": {
        "nome": "Entrega",
        "folder_id": "b52cc860d3a3cc2db7545c9a631de160652ba580",
        "subpastas": None
    },
    "5": {
        "nome": "Puxada",
        "folder_id": "86a14c460a74386b495c9b17d4fdc622701e6f7f",
        "subpastas": None
    },
}


print("\nEscolha a pasta de destino no Autentique:")
for chave, pasta in pastas_autentique.items():
    print(f"{chave} - {pasta['nome']}")

opcao_pasta = input("Opção: ").strip()
while opcao_pasta not in pastas_autentique:
    opcao_pasta = input("Opção inválida. Tente novamente: ").strip()

pasta_escolhida = pastas_autentique[opcao_pasta]

if pasta_escolhida["subpastas"]:
    print(f"\nEscolha a subpasta de {pasta_escolhida['nome']}:")
    for chave, sub in pasta_escolhida["subpastas"].items():
        print(f"{chave} - {sub['nome']}")

    opcao_sub = input("Opção: ").strip()
    while opcao_sub not in pasta_escolhida["subpastas"]:
        opcao_sub = input("Opção inválida. Tente novamente: ").strip()

    folder_id_final = pasta_escolhida["subpastas"][opcao_sub]["folder_id"]
    print(f"Pasta selecionada: {pasta_escolhida['nome']} > {pasta_escolhida['subpastas'][opcao_sub]['nome']}")
else:
    folder_id_final = pasta_escolhida["folder_id"]
    print(f"Pasta selecionada: {pasta_escolhida['nome']}")


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

        
        enviar_para_autentique(caminho_pdf, nome, cpf, numero, folder_id_final)
        
        # Opcional: Se você NÃO quiser guardar o arquivo .docx e quiser APENAS o PDF,
        # descomente a linha abaixo para apagar o Word depois que o PDF for criado:
        # os.remove(caminho_docx)

    except PermissionError:
        print("\n[ERRO] Não foi possível salvar. O arquivo já está aberto por outro programa (Word ou leitor de PDF).")
    except Exception as e:
        print(f"\n[ERRO] Ocorreu um problema: {e}")
else:
    print("\nOperação cancelada pelo usuário.")