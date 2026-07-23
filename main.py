from docxtpl import DocxTemplate
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import os
from docx2pdf import convert  # Nova biblioteca para conversão
import requests
import json


# ==========================
# FUNÇÃO PARA LIMPAR CPF
# ==========================
def limpar_cpf(cpf_bruto):
    """Remove qualquer caractere que não seja número e valida o tamanho."""
    apenas_digitos = "".join(filter(str.isdigit, cpf_bruto))
    
    if len(apenas_digitos) != 11:
        raise ValueError(f"CPF inválido: '{cpf_bruto}' tem {len(apenas_digitos)} dígitos (deveria ter 11)")
    
    return apenas_digitos


# ==========================
# FUNÇÃO PARA NORMALIZAR TELEFONE
# ==========================
def normalizar_telefone(numero_digitado, ddd_padrao="54", permitir_vazio=False):
    """
    Aceita o número em qualquer formato e devolve ele padronizado.
    - DDD é opcional (assume 54 se não vier)
    - O 9 inicial é opcional (assume que existe se não vier)
    - Se o valor vier vazio e permitir_vazio=True, retorna None e "S/N"
    Retorna: (numero_e164, numero_formatado)
    Exemplo: "984481615" -> ("+5554984481615", "(54) 98448-1615")
    """
    if numero_digitado is None:
        if permitir_vazio:
            return None, "S/N"
        raise ValueError("Número de telefone inválido: valor vazio")

    texto = str(numero_digitado).strip()
    if texto == "":
        if permitir_vazio:
            return None, "S/N"
        raise ValueError("Número de telefone inválido: valor vazio")

    if texto.upper() in {"S/N", "SN"}:
        if permitir_vazio:
            return None, "S/N"
        return None, "S/N"

    apenas_digitos = "".join(filter(str.isdigit, texto))

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
def enviar_para_autentique(caminho_pdf, nome_funcionario, cpf_funcionario, telefone_funcionario, folder_id, delivery_method):
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

    cpf_limpo = limpar_cpf(cpf_funcionario)
    telefone_limpo, _ = normalizar_telefone(telefone_funcionario, permitir_vazio=True)

    query = """
    mutation CreateDocumentMutation(
        $document: DocumentInput!,
        $signers: [SignerInput!]!,
        $file: Upload!,
        $folder_id: UUID!
    ) {
        createDocument(
        sandbox: false, 
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

    signer = {
        "name": nome_funcionario,
        "action": "SIGN",
        "configs": {"cpf": cpf_limpo},
        #"security_verifications": [{"type": "UPLOAD"}],
        "positions": [
            {"x": "31.1", "y": "18.5", "z": 3, "element": "SIGNATURE"}
        ]
    }

    # O campo "phone" só é enviado no modo WhatsApp.
    # No modo Link, omitimos o telefone de propósito: se o signatário tiver um
    # telefone associado, o Autentique passa a exigir um código de confirmação
    # (por padrão via WhatsApp) antes de liberar a assinatura — o que trava o
    # fluxo para quem não tem WhatsApp, como os motoristas da Entrega.
    if delivery_method == "DELIVERY_METHOD_WHATSAPP" and telefone_limpo:
        signer["phone"] = telefone_limpo
        signer["delivery_method"] = delivery_method
    else:
        signer["delivery_method"] = delivery_method

    variables = {
        "document": {"name": f"Termo_{nome_funcionario}"},
        "signers": [signer],
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

        # Se for modo LINK (setor Entrega), mostra o link pra você copiar e enviar manualmente
        if delivery_method == "DELIVERY_METHOD_LINK":
            link_encontrado = None
            for assinatura in doc["signatures"]:
                if assinatura.get("link") and assinatura["link"].get("short_link"):
                    link_encontrado = assinatura["link"]["short_link"]
                    break

            if link_encontrado:
                print("\n" + "="*60)
                print("LINK DE ASSINATURA (envie manualmente ao motorista):")
                print(link_encontrado)
                print("="*60)
            else:
                print("\n[AVISO] Não encontrei o link de assinatura na resposta. Confira manualmente no painel.")

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
# SISTEMA DE ETAPAS (permite digitar 'voltar' para corrigir um campo anterior)
# ==========================
def pedir(pergunta, transform=None):
    """
    Pede um valor ao usuário. Se ele digitar 'voltar', retorna o texto especial
    'VOLTAR' para o loop principal saber que deve retroceder uma etapa.
    Se 'transform' levantar ValueError, a pergunta é repetida (sem voltar etapa).
    """
    while True:
        valor = input(pergunta)
        if valor.strip().lower() == "voltar":
            return "VOLTAR"
        if transform:
            try:
                return transform(valor)
            except ValueError as e:
                print(f"[ERRO] {e}")
                continue
        return valor


def validar_cpf_input(v):
    limpar_cpf(v)  # só valida o formato (11 dígitos); o valor salvo continua como digitado
    return v


def validar_telefone_input(v):
    if v is None:
        return "S/N"

    valor = str(v).strip()
    if valor == "":
        return "S/N"

    if valor.upper() in {"S/N", "SN"}:
        return "S/N"

    _, formatado = normalizar_telefone(valor)
    return formatado


def validar_valor_input(v):
    return float(v.replace(",", "."))


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


def validar_opcao_pasta(v):
    if v not in pastas_autentique:
        raise ValueError("Opção inválida. Escolha um número da lista.")
    return v


def deve_executar(etapa, dados):
    """Algumas etapas só existem dependendo de respostas anteriores."""
    if etapa == "subpasta":
        pasta = pastas_autentique[dados["pasta"]]
        return bool(pasta["subpastas"])
    return True


def executar_etapa(etapa, dados):
    if etapa == "nome":
        return pedir("Nome: ", lambda v: v.upper())
    if etapa == "cargo":
        return pedir("Cargo: ", lambda v: v.upper())
    if etapa == "setor":
        return pedir("Área: ", lambda v: v.upper())
    if etapa == "cidade":
        return pedir("Cidade: ", lambda v: v.upper())
    if etapa == "cpf":
        return pedir("CPF: ", validar_cpf_input)
    if etapa == "modelo":
        return pedir("Modelo: ", lambda v: v.upper())
    if etapa == "imei":
        return pedir("IMEI: ", lambda v: v)
    if etapa == "numero":
        return pedir("Número (deixe em branco para S/N): ", validar_telefone_input)
    if etapa == "valor":
        return pedir("Valor: ", validar_valor_input)

    if etapa == "pasta":
        print("\nEscolha a pasta de destino no Autentique:")
        for chave, pasta in pastas_autentique.items():
            print(f"{chave} - {pasta['nome']}")
        return pedir("Opção: ", validar_opcao_pasta)

    if etapa == "subpasta":
        pasta = pastas_autentique[dados["pasta"]]
        print(f"\nEscolha a subpasta de {pasta['nome']}:")
        for chave, sub in pasta["subpastas"].items():
            print(f"{chave} - {sub['nome']}")
        print("0 - Nenhuma (manter direto na pasta principal, sem subpasta)")

        opcoes_validas = set(pasta["subpastas"].keys()) | {"0"}

        def validar_sub(v):
            if v not in opcoes_validas:
                raise ValueError("Opção inválida.")
            return v

        return pedir("Opção: ", validar_sub)

    if etapa == "envio":
        pasta = pastas_autentique[dados["pasta"]]
        sugestao = "2" if pasta["nome"] == "Entrega" else "1"
        print("\nComo deseja enviar o termo para assinatura?")
        print("1 - WhatsApp (automático)")
        print("2 - Link manual (você envia por fora: SMS, ligação, etc.)")

        def validar_envio(v):
            if v == "":
                return sugestao
            if v not in ("1", "2"):
                raise ValueError("Opção inválida. Digite 1 ou 2.")
            return v

        return pedir(f"Opção [Enter para usar sugestão: {sugestao}]: ", validar_envio)


etapas = ["nome", "cargo", "setor", "cidade", "cpf", "modelo",
          "imei", "numero", "valor", "pasta", "subpasta", "envio"]

dados = {}
print("(Em qualquer campo, digite 'voltar' para corrigir o anterior)\n")

i = 0
while i < len(etapas):
    etapa = etapas[i]

    if not deve_executar(etapa, dados):
        i += 1
        continue

    resultado = executar_etapa(etapa, dados)

    if resultado == "VOLTAR":
        i -= 1
        while i >= 0 and not deve_executar(etapas[i], dados):
            i -= 1
        if i < 0:
            i = 0
        continue

    dados[etapa] = resultado
    i += 1

# ==========================
# TRADUZINDO AS RESPOSTAS PARA AS VARIÁVEIS USADAS NO RESTO DO SCRIPT
# ==========================
nome = dados["nome"]
cargo = dados["cargo"]
setor = dados["setor"]
cidade = dados["cidade"]
cpf = dados["cpf"]
modelo = dados["modelo"]
imei = dados["imei"]
numero = dados["numero"]
valor = dados["valor"]

pasta_escolhida = pastas_autentique[dados["pasta"]]

if pasta_escolhida["subpastas"]:
    opcao_sub = dados["subpasta"]
    if opcao_sub == "0":
        folder_id_final = pasta_escolhida["folder_id"]
        print(f"\nPasta selecionada: {pasta_escolhida['nome']} (sem subpasta)")
    else:
        folder_id_final = pasta_escolhida["subpastas"][opcao_sub]["folder_id"]
        print(f"\nPasta selecionada: {pasta_escolhida['nome']} > {pasta_escolhida['subpastas'][opcao_sub]['nome']}")
else:
    folder_id_final = pasta_escolhida["folder_id"]
    print(f"\nPasta selecionada: {pasta_escolhida['nome']}")

if dados["envio"] == "1":
    delivery_method_final = "DELIVERY_METHOD_WHATSAPP"
    print("Envio selecionado: WhatsApp")
else:
    delivery_method_final = "DELIVERY_METHOD_LINK"
    print("Envio selecionado: Link manual")

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

        
        enviar_para_autentique(caminho_pdf, nome, cpf, numero, folder_id_final, delivery_method_final)
        
        # Opcional: Se você NÃO quiser guardar o arquivo .docx e quiser APENAS o PDF,
        # descomente a linha abaixo para apagar o Word depois que o PDF for criado:
        # os.remove(caminho_docx)

    except PermissionError:
        print("\n[ERRO] Não foi possível salvar. O arquivo já está aberto por outro programa (Word ou leitor de PDF).")
    except Exception as e:
        print(f"\n[ERRO] Ocorreu um problema: {e}")
else:
    print("\nOperação cancelada pelo usuário.")