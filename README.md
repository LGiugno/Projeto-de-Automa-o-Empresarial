# Robô Comparativo MultiA

Automação de fechamento de laudos imobiliários para os sistemas **MultiA Mais** e **MultiA Avaliações**. Lê dados de uma planilha Google Sheets, envia comparativos via API REST, atualiza campos da avaliação, gera o laudo PDF e aplica assinaturas digitais tudo em segundo plano, sem abrir navegador.

---

## Funcionalidades

- Cadastro automático de comparativos com upload de imagens
- Leitura de dados diretamente do Google Sheets (localidade, área, valor, fonte, etc.)
- Atualização de campos da avaliação: `PERCENTFORCADA`, `METODO`, `JUSTOFORCADA`, `VALIDADELAUDO`
- Atualização de `VALORUNIDADE` dos grupos de vistoria por correspondência de nomes
- Geração e download automático do laudo PDF
- Assinatura digital do laudo com certificados `.pfx` via JSignPdf (Java)
- Interface gráfica dark mode com log em tempo real
- Upload paralelo de comparativos (3 threads simultâneas)

---

## Pré-requisitos

- Python 3.11+
- Java instalado e no PATH (para assinatura digital)
- [JSignPdf](https://sourceforge.net/projects/jsignpdf/files/stable/) — baixar e colocar o `JSignPdf.jar` na pasta `JSignPdf/`
- Google Service Account com acesso à planilha (ver abaixo)

---

## Instalação

```bash
pip install -r requirements.txt
```

---

## Configuração

### 1. Variáveis de ambiente

Copie `.env.example` para `.env` e preencha com suas credenciais:

```bash
cp .env.example .env
```

O arquivo `.env` **nunca deve ser commitado**. Ele está no `.gitignore`.

### 2. Credenciais Google

Crie um Service Account no [Google Cloud Console](https://console.cloud.google.com/):

1. Crie um projeto
2. Ative a **Google Sheets API** e a **Google Drive API**
3. Crie um Service Account e baixe a chave JSON
4. Renomeie para `credentials.json` e coloque na pasta do projeto
5. Compartilhe a planilha com o e-mail do Service Account

Use `credentials.example.json` como referência da estrutura esperada.

### 3. Certificados digitais

Coloque os arquivos `.pfx` na pasta `Assinaturas/`. Esses arquivos **nunca devem ser commitados**.

---

## Execução

```bash
python robo.py
```

---

## Compilar para EXE (Windows)

```bash
pyinstaller --onefile --windowed --icon="fechamento.ico" robo.py
```

O EXE gerado fica em `dist\robo.exe`.  
Para distribuição: `robo.exe` + `credentials.json` + pasta `Assinaturas/` na mesma pasta.

---

## Estrutura de pastas esperada

```
Comparativos/
├── 15.250/
│   ├── 1.png
│   ├── 2.png
│   └── 3.png
└── 6.699/
    ├── 1.jpg
    └── 2.jpg
```

Cada subpasta deve ter o nome da matrícula. As imagens devem ser nomeadas com o número do comparativo (ex: `1.jpg`, `2.png`).

---

## Estrutura da planilha Google Sheets

Cada aba da planilha corresponde a uma matrícula (nome da aba = nome da subpasta).

| Célula(s) | Conteúdo |
|---|---|
| B13 | Código/matrícula para busca no sistema |
| D29 | Tipo de unidade: `"Área (m²)"` ou `"Área (ha)"` |
| A55:A74 | Número do comparativo |
| B55:B74 | Localidade |
| C55:C74 | Fonte |
| D55:D74 | Área |
| E55:E74 | Valor |
| C80:C86 | Nome do grupo de vistoria (col C), valor em col E |
| A92:A111 | Nome do grupo de vistoria (col A), valor em col E |
| A113:A132 | Nome do grupo de vistoria (col A), valor em col E |
| C140 | Liquidação Forçada (PERCENTFORCADA) |
| B141 | Método de avaliação |
| C141 | JUSTOFORCADA: `"APLICA"` ou `"S"` → `"S"`, senão `"N"` |

---

## Endpoints da API utilizados

| Método | Endpoint | Descrição |
|---|---|---|
| GET | `/multia/avaliacoes` | Busca avaliação por código |
| GET | `/multia/dadosavaliacao/{uuid}` | Dados completos da avaliação |
| GET | `/multia/buscardadosvistoriaimovel/{uuid}` | Grupos de vistoria |
| POST | `/multia/adicionarcomparativo/{uuid}` | Adiciona comparativo com imagem |
| POST | `/multia/editaravaliacao/{uuid}` | Edita campos da avaliação |
| POST | `/multia/salvargrupoimovel/{uuid}/{REG}` | Atualiza VALORUNIDADE dos grupos |

---

## Métodos de avaliação disponíveis

| Na planilha | Enviado para o sistema |
|---|---|
| Comparativo | Comparativo direto de dados de mercado |
| Evolutivo | Evolutivo |
| Capitalização da renda | Capitalização da renda |
| Involutivo | Involutivo |
