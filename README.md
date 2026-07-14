# Termo Automático - Disfonte

Automação para geração e envio de Termos de Responsabilidade para Smartphone, com assinatura eletrônica via Autentique.

## O que o script faz

1. Coleta os dados do funcionário via terminal (nome, cargo, área, cidade, CPF, modelo do aparelho, IMEI, número, valor)
2. Gera o Termo de Responsabilidade preenchido a partir do modelo (`modelo.docx`)
3. Converte o documento gerado para PDF
4. Envia o PDF para assinatura eletrônica no Autentique, já configurado com:
   - Envio via WhatsApp
   - Validação de identidade por documento com foto
   - CPF do signatário
   - Posição fixa do campo de assinatura (calibrada para o layout do `modelo.docx`)
   - Pasta de destino no Autentique, escolhida no momento da execução

## Pré-requisitos

- Python 3.10+
- Microsoft Word instalado (necessário para a conversão DOCX → PDF via `docx2pdf`)
- Uma chave de API do Autentique (gerada em **Chaves de API** no painel do Autentique)

## Instalação

```bash
python -m venv .venv
.venv\Scripts\Activate.ps1        # Windows PowerShell

pip install docxtpl docx2pdf requests
```

## Configuração do token da API

O token **não** fica salvo no código. Ele é lido de uma variável de ambiente do Windows.

1. Gere seu token em: painel do Autentique → Chaves de API
2. No PowerShell, rode uma única vez:

```powershell
setx AUTENTIQUE_TOKEN "seu_token_aqui"
```

3. **Feche e abra novamente** o terminal (e o VS Code, se for o caso) para a variável ser reconhecida.
4. Para conferir se foi salva corretamente:

```powershell
python -c "import os; print(os.environ.get('AUTENTIQUE_TOKEN'))"
```

Se aparecer o token na tela, está tudo certo.

## Arquivos do projeto

| Arquivo | Função |
|---|---|
| `main.py` | Script principal — roda o fluxo completo |
| `modelo.docx` | Modelo do Termo de Responsabilidade, com variáveis `{{ }}` do docxtpl |
| `calibrar_posicao.py` | Script auxiliar usado para calibrar a posição (x, y, z) do campo de assinatura no PDF |
| `verificar_documento.py` | Script auxiliar para consultar um documento já criado e conferir se a posição de assinatura foi salva |

## Como usar

```powershell
python main.py
```

O script vai pedir, em ordem:

1. Nome, cargo, área, cidade, CPF, modelo, IMEI, número de telefone e valor do aparelho
2. A pasta de destino no Autentique (Administrativo, Armazém, Comercial, Entrega ou Puxada — algumas com subpastas)
3. O local para salvar o `.docx` (o `.pdf` é salvo automaticamente ao lado, com o mesmo nome)

Ao final, o termo é enviado automaticamente para assinatura via WhatsApp, na pasta escolhida.

## Sandbox (modo de teste)

Enquanto a mutation `createDocument` estiver com `sandbox: true`, os documentos criados:
- Não consomem créditos do plano
- Não têm validade jurídica
- São apagados automaticamente após alguns dias

**Antes de usar com funcionários reais**, remova ou altere essa linha para `sandbox: false` (ou apague o parâmetro) dentro da função `enviar_para_autentique()` em `main.py`.

## Pastas do Autentique (IDs configurados)

Os IDs das pastas ficam no dicionário `pastas_autentique`, dentro de `main.py`. Caso novas pastas ou subpastas sejam criadas no Autentique, é necessário:

1. Obter o novo `id` via consulta GraphQL (`query { folders { data { id name } } }`, testável em [Altair](https://altair.autentique.com.br))
2. Adicionar a entrada correspondente no dicionário `pastas_autentique`

## Posição do campo de assinatura

A posição (`x`, `y`, `z`) do campo de assinatura foi calibrada manualmente para o layout atual do `modelo.docx` (página 3, linha de assinatura). Valores atuais:

```python
{"x": "31.1", "y": "18.5", "z": 3, "element": "SIGNATURE"}
```

Se o `modelo.docx` for alterado de forma que mude a quantidade de páginas ou a posição da linha de assinatura, é necessário recalibrar usando o `calibrar_posicao.py`.

## Segurança

- O token de API nunca deve ser digitado diretamente no código-fonte
- O CPF e telefone dos funcionários são normalizados e validados antes do envio (ver funções `limpar_cpf()` e `normalizar_telefone()` em `main.py`)
- Recomenda-se não versionar (Git) arquivos com dados reais de funcionários (termos gerados, PDFs assinados, etc.)
