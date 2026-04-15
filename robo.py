"""
Robô Comparativo MultiA - 100% em Segundo Plano
================================================
Usa API REST + Google Sheets API. Não abre navegador.
Roda completamente em background sem interferir na tela.

Uso:
    python robo.py

Estrutura esperada:
    Comparativos/
    ├── 15.250/
    │   ├── 1.png
    │   ├── 2.png
    │   └── 3.png
    └── 6.699/
        ├── 1.jpg
        └── 2.jpg
"""

import os
import sys
import re
import json
import time
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path
from threading import Thread
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from typing import Optional

import customtkinter as ctk
import requests
import gspread
from google.oauth2.service_account import Credentials

# subprocess — usado para chamar JSignPdf (assinatura via Java)
import subprocess

def _base_dir() -> Path:
    """Retorna a pasta do EXE (compilado) ou do script (dev)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    else:
        return Path(__file__).parent

# ============================================================
# CONFIGURAÇÃO
# Preencha com as credenciais do seu ambiente.
# NÃO commite credenciais reais no repositório.
# Use variáveis de ambiente ou um arquivo .env (ver .env.example).
# ============================================================

def _carregar_sistema(nome: str) -> dict:
    """Carrega configuração do sistema a partir de variáveis de ambiente."""
    prefixo = nome.upper().replace(" ", "_").replace("Ã", "A").replace("Ç", "C")
    return {
        "base_url":      os.environ.get(f"{prefixo}_BASE_URL", ""),
        "origin":        os.environ.get(f"{prefixo}_ORIGIN", ""),
        "referer":       os.environ.get(f"{prefixo}_REFERER", ""),
        "authorization": os.environ.get(f"{prefixo}_AUTHORIZATION", ""),
        "login_fixo":    os.environ.get(f"{prefixo}_LOGIN", ""),
        "senha_fixa":    os.environ.get(f"{prefixo}_SENHA", ""),
        "jwt_fixo":      os.environ.get(f"{prefixo}_JWT", ""),
    }

SISTEMAS = {
    "MultiA Mais":       _carregar_sistema("MULTIA_MAIS"),
    "MultiA Avaliações": _carregar_sistema("MULTIA_AVALIACOES"),
}

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Senha dos certificados .pfx — lida de variável de ambiente
SENHA_CERTIFICADOS = os.environ.get("CERT_PASSWORD", "")


# ============================================================
# CLASSES DE DADOS
# ============================================================

@dataclass
class ComparativoData:
    numero: int
    localidade: str
    fonte: str
    area: str
    valor: str
    unidade: str
    imagem_path: str


@dataclass
class ConfigData:
    sistema: str = "MultiA Mais"
    planilha_id: str = ""
    pasta_comparativos: str = ""
    credentials_path: str = ""
    usuario: str = ""
    senha: str = ""
    excluir_imagens: bool = True
    gerar_laudo: bool = True


# ============================================================
# CLIENTE API MULTIA
# ============================================================

class MultiAAPI:
    """Cliente para a API do sistema MultiA - tudo via HTTP, sem navegador."""

    def __init__(self, sistema_config: dict, jwt: str, logger: logging.Logger):
        self.config = sistema_config
        self.jwt = jwt
        self.logger = logger
        self.session = requests.Session()
        self.session.headers.update({
            "authorization": sistema_config["authorization"],
            "jwt": jwt,
            "Accept": "*/*",
            "Origin": sistema_config["origin"],
            "Referer": sistema_config["referer"],
        })
        self.base_url = sistema_config["base_url"]

    def buscar_avaliacoes(self, busca: str, page: int = 0, page_size: int = 50) -> dict:
        """Busca avaliações por matrícula/documento."""
        url = f"{self.base_url}/multia/avaliacoes"
        params = {
            "sortField": "",
            "sortOrder": "",
            "pageSize": page_size,
            "page": page,
            "busca": busca,
            "REGSTATUS": "",
            "DATACRIACAO": "",
        }
        self.logger.info(f"  Buscando avaliações para: {busca}")
        resp = self.session.get(url, params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def buscar_avaliacao_por_codigo(self, codigo: str) -> Optional[dict]:
        """Busca avaliação diretamente pelo Código (REG único do sistema)."""
        url = f"{self.base_url}/multia/avaliacoes"
        params = {
            "sortField": "", "sortOrder": "",
            "pageSize": 50, "page": 0,
            "busca": codigo, "REGSTATUS": "", "DATACRIACAO": "",
        }
        self.logger.info(f"  Buscando avaliação pelo código: {codigo}")
        resp = self.session.get(url, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        if data.get("status") != "sucesso":
            return None
        for av in data.get("dados", {}).get("avaliacoes", []):
            if str(av.get("REG", "")).strip() == str(codigo).strip():
                return av
        return None

    def buscar_dados_avaliacao(self, uuid: str) -> dict:
        """Busca dados completos de uma avaliação."""
        url = f"{self.base_url}/multia/dadosavaliacao/{uuid}"
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def buscar_dados_vistoria(self, uuid: str) -> dict:
        """Busca dados da vistoria de uma avaliação."""
        url = f"{self.base_url}/multia/buscardadosvistoriaimovel/{uuid}"
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def adicionar_comparativo(self, uuid: str, localidade: str, unidade: str,
                               area: str, valor: str, fonte: str,
                               imagem_path: str) -> dict:
        """Adiciona um comparativo com imagem via multipart/form-data."""
        url = f"{self.base_url}/multia/adicionarcomparativo/{uuid}"

        data = {
            "LOCALIDADE": localidade,
            "UNIDADE": unidade,
            "AREA": str(area),
            "VALOR": str(valor),
            "FONTE": fonte,
        }

        filename = os.path.basename(imagem_path)
        mime = "image/png" if imagem_path.lower().endswith(".png") else "image/jpeg"

        with open(imagem_path, "rb") as f:
            files = {"arquivo": (filename, f, mime)}
            resp = self.session.post(url, data=data, files=files, timeout=60)

        resp.raise_for_status()
        return resp.json()

    def editar_avaliacao(self, uuid: str, **campos) -> dict:
        """Edita campos da avaliação (ex: PERCENTFORCADA, PERCENTJUSTA, PARECERLAUDO, etc)."""
        url = f"{self.base_url}/multia/editaravaliacao/{uuid}"
        self.logger.info(f"  POST {url} → {campos}")
        resp = self.session.post(url, data=campos, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def salvar_grupo_imovel(self, uuid: str, reg: str, **campos) -> dict:
        """Edita campos de um grupo do imóvel."""
        url = f"{self.base_url}/multia/salvargrupoimovel/{uuid}/{reg}"
        resp = self.session.post(url, data=campos, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def buscar_grupos_vistoria(self, uuid: str) -> list[dict]:
        """Busca os grupos de vistoria do imóvel via GET."""
        url = f"{self.base_url}/multia/buscardadosvistoriaimovel/{uuid}"
        self.logger.info(f"  GET {url}")
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        return data.get("dados", {}).get("grupos", [])

    def buscar_nome_laudo(self, uuid: str) -> str:
        """Retorna o nome sugerido para o arquivo do laudo."""
        url = f"{self.base_url}/multia/dadosnomearquivolaudo/{uuid}"
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        return data.get("dados", "") or "laudo"

    def gerar_laudo(self, uuid: str, logo_bytes: bytes) -> bytes:
        """Gera o laudo PDF enviando a logo da empresa. Retorna os bytes do PDF."""
        url = f"{self.base_url}/multia/gerarlaudo/{uuid}"
        files = {"LOGO": ("blob", logo_bytes, "image/png")}
        resp = self.session.post(url, files=files, timeout=120)
        resp.raise_for_status()
        return resp.content


# ============================================================
# ASSINATURA DIGITAL PDF
# ============================================================

# Coordenadas fixas dos campos de assinatura (em pontos PDF)
ASSINATURA_COORDS = {
    "assinatura1": {"x": 124, "y": 338, "w": 172, "h": 80},
    "assinatura2": {"x": 300, "y": 338, "w": 173, "h": 80},
}

# Mapeamento de sistema → certificados
CERT_POR_SISTEMA = {
    "MultiA Mais": {
        "pessoa":  os.environ.get("CERT_PESSOA_MAIS", "leandro.pfx"),
        "empresa": os.environ.get("CERT_EMPRESA_MAIS", "multia_mais.pfx"),
    },
    "MultiA Avaliações": {
        "pessoa":  os.environ.get("CERT_PESSOA_AVAL", "matheus.pfx"),
        "empresa": os.environ.get("CERT_EMPRESA_AVAL", "multia_avaliacoes.pfx"),
    },
}

# Mapa nome-exibição → arquivo .pfx
OPCOES_PESSOA = {
    "Leandro": os.environ.get("CERT_PESSOA_MAIS", "leandro.pfx"),
    "Matheus": os.environ.get("CERT_PESSOA_AVAL", "matheus.pfx"),
}
OPCOES_EMPRESA = {
    "MultiA Mais":       os.environ.get("CERT_EMPRESA_MAIS", "multia_mais.pfx"),
    "MultiA Avaliações": os.environ.get("CERT_EMPRESA_AVAL", "multia_avaliacoes.pfx"),
}


def dialogo_assinaturas(nome_laudo: str, sistema: str) -> Optional[tuple[str, str]]:
    """
    Exibe uma janela modal para o usuário escolher os certificados de assinatura.

    Retorna (pfx_pessoa, pfx_empresa) com os nomes dos arquivos .pfx,
    ou None se o usuário cancelar.
    """
    resultado = [None]

    def _build():
        win = tk.Toplevel()
        win.title("Selecionar Assinaturas")
        win.resizable(False, False)
        win.grab_set()

        BG      = "#1E1E2E"
        FG      = "#E2E8F0"
        FG_MUT  = "#94A3B8"
        ACCENT  = "#6366F1"
        CARD    = "#2A2A3E"
        BORDER  = "#3B3B52"
        OK_CLR  = "#22C55E"

        win.configure(bg=BG)

        frm_top = tk.Frame(win, bg=BG, padx=24, pady=18)
        frm_top.pack(fill="x")
        tk.Label(frm_top, text="✍  Assinaturas do Laudo",
                 font=("Segoe UI", 13, "bold"), bg=BG, fg=FG).pack(anchor="w")
        tk.Label(frm_top, text=f"Laudo:  {nome_laudo}",
                 font=("Segoe UI", 10), bg=BG, fg=FG_MUT).pack(anchor="w", pady=(4, 0))

        sep = tk.Frame(win, bg=BORDER, height=1)
        sep.pack(fill="x", padx=16)

        frm_body = tk.Frame(win, bg=BG, padx=24, pady=16)
        frm_body.pack(fill="x")

        default_pessoa  = "Leandro" if sistema == "MultiA Mais" else "Matheus"
        default_empresa = "MultiA Mais" if sistema == "MultiA Mais" else "MultiA Avaliações"

        var_pessoa  = tk.StringVar(value=default_pessoa)
        var_empresa = tk.StringVar(value=default_empresa)

        def _grupo(parent, titulo, opcoes: dict, var):
            frm = tk.Frame(parent, bg=CARD, bd=0, relief="flat",
                           highlightbackground=BORDER, highlightthickness=1)
            frm.pack(fill="x", pady=(0, 12))

            tk.Label(frm, text=titulo, font=("Segoe UI", 9, "bold"),
                     bg=CARD, fg=ACCENT, padx=14, pady=8).pack(anchor="w")

            sep2 = tk.Frame(frm, bg=BORDER, height=1)
            sep2.pack(fill="x", padx=0)

            for nome in opcoes:
                rb = tk.Radiobutton(
                    frm, text=f"  {nome}",
                    variable=var, value=nome,
                    font=("Segoe UI", 10),
                    bg=CARD, fg=FG,
                    activebackground=CARD, activeforeground=FG,
                    selectcolor=ACCENT,
                    bd=0, padx=14, pady=6,
                    cursor="hand2",
                )
                rb.pack(anchor="w")

        _grupo(frm_body, "Assinatura Pessoa",  OPCOES_PESSOA,  var_pessoa)
        _grupo(frm_body, "Assinatura Empresa", OPCOES_EMPRESA, var_empresa)

        sep2 = tk.Frame(win, bg=BORDER, height=1)
        sep2.pack(fill="x", padx=16)

        frm_btn = tk.Frame(win, bg=BG, padx=24, pady=14)
        frm_btn.pack(fill="x")

        def _confirmar():
            resultado[0] = (
                OPCOES_PESSOA[var_pessoa.get()],
                OPCOES_EMPRESA[var_empresa.get()],
            )
            win.destroy()

        def _cancelar():
            resultado[0] = None
            win.destroy()

        btn_ok = tk.Button(
            frm_btn, text="✓  Confirmar e Assinar",
            font=("Segoe UI", 10, "bold"),
            bg=OK_CLR, fg="#000000",
            activebackground="#16A34A", activeforeground="#000000",
            bd=0, padx=18, pady=8, cursor="hand2",
            command=_confirmar,
        )
        btn_ok.pack(side="right", padx=(8, 0))

        btn_cancel = tk.Button(
            frm_btn, text="Pular assinatura",
            font=("Segoe UI", 10),
            bg=CARD, fg=FG_MUT,
            activebackground=BORDER, activeforeground=FG,
            bd=0, padx=14, pady=8, cursor="hand2",
            command=_cancelar,
        )
        btn_cancel.pack(side="right")

        win.update_idletasks()
        w = win.winfo_reqwidth()
        h = win.winfo_reqheight()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        win.wait_window()

    _build()
    return resultado[0]


def _gerar_imagem_assinatura(nome: str, largura_pts: int, altura_pts: int,
                             bg_img_path: Optional[Path] = None) -> Path:
    """
    Gera uma imagem PNG com o visual da assinatura centralizado.

    - nome: nome extraído do certificado
    - largura_pts / altura_pts: dimensões da caixa em pontos PDF
    - bg_img_path: caminho do bg_assinatura.png (logo — marca d'água)
    """
    from PIL import Image, ImageDraw, ImageFont
    from datetime import datetime
    import tempfile, textwrap

    SCALE = 5
    w = largura_pts * SCALE
    h = altura_pts * SCALE

    img = Image.new("RGB", (w, h), (255, 255, 255))

    if bg_img_path and bg_img_path.exists():
        try:
            logo = Image.open(bg_img_path).convert("RGBA")
            pixels = logo.load()
            for py in range(logo.height):
                for px in range(logo.width):
                    r_val, g_val, b_val, a_val = pixels[px, py]
                    if r_val > 240 and g_val > 240 and b_val > 240:
                        pixels[px, py] = (r_val, g_val, b_val, 0)
            ratio = (h * 0.715) / logo.height
            new_w = int(logo.width * ratio)
            new_h = int(logo.height * ratio)
            logo = logo.resize((new_w, new_h), Image.LANCZOS)
            r, g, b, a = logo.split()
            a = a.point(lambda p: int(p * 0.12))
            logo = Image.merge("RGBA", (r, g, b, a))
            img_rgba = img.convert("RGBA")
            lx = (w - new_w) // 2
            ly = (h - new_h) // 2
            img_rgba.paste(logo, (lx, ly), logo)
            img = img_rgba.convert("RGB")
        except Exception:
            pass

    draw = ImageDraw.Draw(img)

    def _carregar_fonte(tamanho):
        candidatas = [
            "arial.ttf", "Arial.ttf",
            "calibri.ttf", "Calibri.ttf",
            "segoeui.ttf", "SegoeUI.ttf",
            "C:/Windows/Fonts/arial.ttf",
            "C:/Windows/Fonts/calibri.ttf",
            "C:/Windows/Fonts/segoeui.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ]
        for nome_fonte in candidatas:
            try:
                return ImageFont.truetype(nome_fonte, tamanho)
            except (OSError, IOError):
                continue
        return ImageFont.load_default()

    TAMANHO_FONTE = int(h * 0.098)
    if TAMANHO_FONTE < 14:
        TAMANHO_FONTE = 14
    font = _carregar_fonte(TAMANHO_FONTE)

    timestamp = datetime.now().strftime("%Y.%m.%d %H:%M:%S BRT")
    margem = int(w * 0.05)
    largura_util = w - 2 * margem

    def _quebrar_texto(texto, font, max_width):
        """Quebra texto em múltiplas linhas respeitando a largura máxima."""
        palavras = texto.split()
        linhas = []
        linha_atual = ""
        for palavra in palavras:
            teste = f"{linha_atual} {palavra}".strip()
            bbox = draw.textbbox((0, 0), teste, font=font)
            if (bbox[2] - bbox[0]) <= max_width:
                linha_atual = teste
            else:
                if linha_atual:
                    linhas.append(linha_atual)
                linha_atual = palavra
        if linha_atual:
            linhas.append(linha_atual)
        return linhas

    linhas_final = ["Assinado digitalmente por"]
    linhas_final += _quebrar_texto(nome, font, largura_util)
    linhas_final.append(f"Data: {timestamp}")

    espacamento = int(TAMANHO_FONTE * 0.55)
    alturas = []
    for linha in linhas_final:
        bbox = draw.textbbox((0, 0), linha, font=font)
        alturas.append(bbox[3] - bbox[1])

    total_h = sum(alturas) + espacamento * (len(linhas_final) - 1)
    y = max(4, (h - total_h) // 2 - int(h * 0.06))

    for i, linha in enumerate(linhas_final):
        bbox = draw.textbbox((0, 0), linha, font=font)
        tw = bbox[2] - bbox[0]
        x = (w - tw) // 2
        draw.text((x, y), linha, fill=(0, 0, 0), font=font)
        y += alturas[i] + espacamento

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False, prefix="sig_visual_")
    img.save(tmp.name, "PNG", quality=100)
    tmp.close()
    return Path(tmp.name)


def assinar_pdf(caminho_pdf: Path, pfx_pessoa: str, pfx_empresa: str,
                logger: logging.Logger) -> bool:
    """
    Assina digitalmente o PDF usando JSignPdf (Java).

    Requer:
      - Java instalado (java no PATH)
      - JSignPdf.jar na pasta JSignPdf/ ao lado do robo.exe/robo.py
        Download: https://sourceforge.net/projects/jsignpdf/files/stable/

    Assina 2 vezes em sequência:
      1. Assinatura da Pessoa  (campo esquerdo)
      2. Assinatura da Empresa (campo direito)
    """
    pasta_assinaturas = _base_dir() / "Assinaturas"
    pasta_jsignpdf    = _base_dir() / "JSignPdf"
    jar_path          = pasta_jsignpdf / "JSignPdf.jar"

    pfx_path_pessoa  = pasta_assinaturas / pfx_pessoa
    pfx_path_empresa = pasta_assinaturas / pfx_empresa

    if not jar_path.exists():
        logger.error(f"  JSignPdf.jar não encontrado em: {jar_path}")
        logger.error(f"  Baixe em: https://sourceforge.net/projects/jsignpdf/files/stable/")
        return False

    for pfx in (pfx_path_pessoa, pfx_path_empresa):
        if not pfx.exists():
            logger.error(f"  Certificado não encontrado: {pfx}")
            return False

    c1 = ASSINATURA_COORDS["assinatura1"]
    c2 = ASSINATURA_COORDS["assinatura2"]

    def _total_paginas(pdf: Path) -> int:
        try:
            data = pdf.read_bytes()
            import re as _re
            counts = _re_pages.findall(data)
            if counts:
                return max(int(c) for c in counts)
        except Exception:
            pass
        return 1

    import re as _re_mod
    _re_pages = _re_mod.compile(rb'/Count\s+(\d+)')

    total_pags = _total_paginas(caminho_pdf)

    def _assinar(pfx_path: Path, campo: int, coords: dict, pdf_entrada: Path, pdf_saida: Path) -> bool:
        """Chama JSignPdf via linha de comando para uma assinatura."""
        import tempfile, shutil

        llx = coords["x"]
        lly = coords["y"]
        urx = coords["x"] + coords["w"]
        ury = coords["y"] + coords["h"]

        with tempfile.TemporaryDirectory(prefix="jsignpdf_") as tmp_dir:
            tmp_in     = Path(tmp_dir) / "input.pdf"
            tmp_signed = Path(tmp_dir) / "input_signed.pdf"

            shutil.copy2(pdf_entrada, tmp_in)

            def _nome_cert(pfx_path):
                try:
                    from cryptography.hazmat.primitives.serialization import pkcs12
                    from cryptography.x509.oid import NameOID
                    with open(pfx_path, "rb") as f:
                        _, cert, _ = pkcs12.load_key_and_certificates(
                            f.read(), SENHA_CERTIFICADOS.encode())
                    for attr in cert.subject:
                        if attr.oid == NameOID.COMMON_NAME:
                            return attr.value
                except Exception:
                    pass
                return pfx_path.stem

            nome = _nome_cert(pfx_path)

            bg_logo = _base_dir() / "bg_assinatura.png"
            img_assinatura = _gerar_imagem_assinatura(
                nome=nome,
                largura_pts=coords["w"],
                altura_pts=coords["h"],
                bg_img_path=bg_logo if bg_logo.exists() else None,
            )

            cmd = [
                "java", "-jar", str(jar_path),
                "-kst", "PKCS12",
                "-ksf", str(pfx_path),
                "-ksp", SENHA_CERTIFICADOS,
                "-V",
                "-pg", str(total_pags),
                "-llx", str(llx),
                "-lly", str(lly),
                "-urx", str(urx),
                "-ury", str(ury),
                "--render-mode", "DESCRIPTION_ONLY",
                "--l2-text", "",
                "--bg-path", str(img_assinatura),
                "--bg-scale", "0",
                "-d",  tmp_dir,
                "-q",
            ]

            if campo == 2:
                cmd.append("-a")

            cmd.append(str(tmp_in))

            logger.info(f"  Executando JSignPdf: {pfx_path.name} → campo {campo}")
            try:
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=120,
                    cwd=str(pasta_jsignpdf),
                    creationflags=subprocess.CREATE_NO_WINDOW,
                )
                if result.returncode != 0:
                    logger.error(f"  JSignPdf erro (código {result.returncode}): {result.stderr[:400]}")
                    return False

                if not tmp_signed.exists():
                    logger.error(f"  JSignPdf: arquivo assinado não encontrado: {tmp_signed}")
                    return False

                shutil.copy2(tmp_signed, pdf_saida)
                return True

            except FileNotFoundError:
                logger.error("  Java não encontrado. Instale o Java e adicione ao PATH.")
                return False
            except subprocess.TimeoutExpired:
                logger.error("  JSignPdf excedeu o tempo limite (120s).")
                return False
            finally:
                try:
                    img_assinatura.unlink(missing_ok=True)
                except Exception:
                    pass

    try:
        pdf_temp1 = caminho_pdf.parent / (caminho_pdf.stem + "_sig1.pdf")
        pdf_temp2 = caminho_pdf.parent / (caminho_pdf.stem + "_sig2.pdf")

        logger.info(f"  Assinando com {pfx_pessoa} (Pessoa)...")
        ok1 = _assinar(pfx_path_pessoa, 1, c1, caminho_pdf, pdf_temp1)
        if not ok1 or not pdf_temp1.exists():
            logger.error("  ✗ Assinatura 1 falhou")
            return False
        logger.info("  ✓ Assinatura 1 (Pessoa) aplicada")

        logger.info(f"  Assinando com {pfx_empresa} (Empresa)...")
        ok2 = _assinar(pfx_path_empresa, 2, c2, pdf_temp1, pdf_temp2)
        if not ok2 or not pdf_temp2.exists():
            logger.error("  ✗ Assinatura 2 falhou")
            pdf_temp1.unlink(missing_ok=True)
            return False
        logger.info("  ✓ Assinatura 2 (Empresa) aplicada")

        caminho_pdf.unlink()
        pdf_temp2.rename(caminho_pdf)
        pdf_temp1.unlink(missing_ok=True)

        logger.info(f"  ✓ PDF assinado salvo: {caminho_pdf}")
        return True

    except Exception as e:
        logger.error(f"  ✗ Erro ao assinar PDF: {e}")
        for tmp in [pdf_temp1, pdf_temp2]:
            try:
                tmp.unlink(missing_ok=True)
            except Exception:
                pass
        return False


# ============================================================
# CLIENTE GOOGLE SHEETS
# ============================================================

class PlanilhaClient:
    """Lê dados da planilha de comparativos."""

    def __init__(self, credentials_path: str, planilha_id: str, logger: logging.Logger):
        self.logger = logger
        self.planilha_id = planilha_id
        self._cache_abas = {}

        creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
        self.gc = gspread.authorize(creds)
        self.planilha = self.gc.open_by_key(planilha_id)
        self.logger.info(f"  Planilha conectada: {self.planilha.title}")

    def _get_aba_valores(self, matricula: str):
        """Retorna (aba, todos_valores) usando cache para evitar downloads duplicados."""
        if matricula in self._cache_abas:
            return self._cache_abas[matricula]

        try:
            aba = self.planilha.worksheet(matricula)
        except gspread.exceptions.WorksheetNotFound:
            matricula_alt = matricula.replace(".", "")
            try:
                aba = self.planilha.worksheet(matricula_alt)
            except gspread.exceptions.WorksheetNotFound:
                raise ValueError(f"Aba '{matricula}' não encontrada na planilha")

        todos_valores = aba.get_all_values()
        self._cache_abas[matricula] = (aba, todos_valores)
        return aba, todos_valores

    def ler_dados_matricula(self, matricula: str, numeros_imagens: list[int]) -> list[ComparativoData]:
        """Lê dados dos comparativos de uma aba da planilha."""
        self.logger.info(f"  Lendo aba '{matricula}' da planilha...")

        aba, todos_valores = self._get_aba_valores(matricula)

        matricula_sistema = ""
        if len(todos_valores) >= 13 and len(todos_valores[12]) >= 2:
            matricula_sistema = todos_valores[12][1].strip()
        if not matricula_sistema:
            matricula_sistema = matricula
            self.logger.warning(f"  B13 vazia — usando nome da pasta como matrícula: {matricula}")
        else:
            self.logger.info(f"  Matrícula do sistema (B13): {matricula_sistema}")

        unidade = ""
        if len(todos_valores) >= 29 and len(todos_valores[28]) >= 4:
            raw_unidade = todos_valores[28][3].strip()
            match_unidade = re.search(r'\(([^)]+)\)', raw_unidade)
            if match_unidade:
                unidade = match_unidade.group(1).strip()
            else:
                unidade = raw_unidade
        if not unidade:
            unidade = "m²"

        liq_forcada = ""
        if len(todos_valores) >= 140 and len(todos_valores[139]) >= 3:
            liq_forcada = todos_valores[139][2].strip()

        metodo_raw = ""
        if len(todos_valores) >= 141 and len(todos_valores[140]) >= 2:
            metodo_raw = todos_valores[140][1].strip()

        MAPA_METODO = {
            "comparativo": "Comparativo direto de dados de mercado",
            "evolutivo": "Evolutivo",
            "capitalização da renda": "Capitalização da renda",
            "capitalizacao da renda": "Capitalização da renda",
            "involutivo": "Involutivo",
        }
        metodo = MAPA_METODO.get(metodo_raw.lower(), metodo_raw)

        justo_raw = ""
        if len(todos_valores) >= 141 and len(todos_valores[140]) >= 3:
            justo_raw = todos_valores[140][2].strip()

        justo_forcada = "S" if justo_raw.upper() in ("S", "APLICA") else "N" if justo_raw else ""

        comparativos = []
        for num_img in sorted(numeros_imagens):
            encontrado = False
            for row_idx in range(54, min(74, len(todos_valores))):
                row = todos_valores[row_idx]
                if len(row) < 1:
                    continue

                celula_a = row[0].strip()
                try:
                    num_celula = int(re.sub(r'\D', '', celula_a)) if celula_a else 0
                except ValueError:
                    continue

                if num_celula == num_img:
                    localidade = row[1].strip() if len(row) > 1 else ""
                    fonte = row[2].strip() if len(row) > 2 else ""
                    area = row[3].strip() if len(row) > 3 else ""
                    valor = row[4].strip() if len(row) > 4 else ""

                    comparativos.append(ComparativoData(
                        numero=num_img,
                        localidade=localidade,
                        fonte=fonte,
                        area=area,
                        valor=valor,
                        unidade=unidade,
                        imagem_path="",
                    ))
                    encontrado = True
                    break

            if not encontrado:
                self.logger.warning(f"    Imagem {num_img}: dados não encontrados na planilha")

        return comparativos, liq_forcada, metodo, justo_forcada, matricula_sistema

    def ler_grupos_vistoria(self, matricula: str) -> list[dict]:
        """Lê nomes e valores de área dos grupos de vistoria na planilha.

        Busca nas células:
          - C80:C86  → nome em coluna C, valor em coluna E (mesma linha)
          - A92:A111 → nome em coluna A, valor em coluna E (mesma linha)
          - A113:A132 → nome em coluna A, valor em coluna E (mesma linha)

        Retorna lista de dicts: [{"nome": str, "valor": str}, ...]
        Ignora linhas com célula de nome vazia.
        """
        self.logger.info(f"  Lendo grupos de vistoria da aba '{matricula}'...")

        _, todos_valores = self._get_aba_valores(matricula)
        grupos = []

        faixas = [
            (79, 86, 2, 4),
            (91, 111, 0, 4),
            (112, 132, 0, 4),
        ]

        for row_ini, row_fim, col_nome, col_valor in faixas:
            for row_idx in range(row_ini, min(row_fim, len(todos_valores))):
                row = todos_valores[row_idx]
                if len(row) <= col_nome:
                    continue
                nome = row[col_nome].strip()
                if not nome:
                    continue
                valor = row[col_valor].strip() if len(row) > col_valor else ""

                if valor == "" and col_nome == 0:
                    val_c = row[2].strip() if len(row) > 2 else ""
                    try:
                        if float(val_c.replace(",", ".")) == 0:
                            valor = "0"
                            self.logger.info(f"    Grupo '{nome}': col E vazia, col C=0 → enviando 0")
                    except (ValueError, AttributeError):
                        pass

                if valor != "":
                    grupos.append({"nome": nome, "valor": valor})
                    self.logger.info(f"    Grupo encontrado: '{nome}' → {valor}")

        self.logger.info(f"  Total de grupos lidos: {len(grupos)}")
        return grupos


# ============================================================
# MOTOR DO ROBÔ
# ============================================================

class RoboComparativo:
    """Orquestra todo o processo em segundo plano."""

    def __init__(self, config: ConfigData, logger: logging.Logger,
                 callback_progresso=None, callback_confirmar=None,
                 callback_validade=None, mapa_assinaturas=None):
        self.config = config
        self.logger = logger
        self.callback_progresso = callback_progresso or (lambda msg: None)
        self.callback_confirmar = callback_confirmar
        self.callback_validade = callback_validade
        self.mapa_assinaturas = mapa_assinaturas or {}
        self.api: Optional[MultiAAPI] = None
        self.planilha: Optional[PlanilhaClient] = None
        self._cancelado = False

    def cancelar(self):
        self._cancelado = True

    def _log(self, msg: str):
        self.logger.info(msg)
        self.callback_progresso(msg)

    def _obter_jwt(self, sistema_config: dict) -> str:
        """Retorna o JWT configurado para o sistema."""
        jwt = sistema_config.get("jwt_fixo", "")
        if not jwt:
            raise ValueError(f"JWT não configurado para o sistema '{self.config.sistema}'")
        self._log(f"  JWT carregado para {self.config.sistema}")
        return jwt

    def _listar_subpastas(self) -> list[tuple[str, list[str]]]:
        """Lista subpastas e suas imagens."""
        pasta = Path(self.config.pasta_comparativos)
        resultado = []

        for subpasta in sorted(pasta.iterdir()):
            if not subpasta.is_dir():
                continue

            imagens = sorted([
                f.name for f in subpasta.iterdir()
                if f.suffix.lower() in IMAGE_EXTENSIONS
            ], key=lambda x: int(re.sub(r'\D', '', Path(x).stem) or 0))

            if imagens:
                resultado.append((subpasta.name, imagens))

        return resultado

    def _extrair_numeros(self, texto: str) -> str:
        """Extrai apenas dígitos de uma string."""
        return re.sub(r'\D', '', texto)

    def _encontrar_avaliacao_por_codigo(self, codigo: str) -> Optional[dict]:
        """Busca avaliação diretamente pelo Código único do sistema (campo REG)."""
        av = self.api.buscar_avaliacao_por_codigo(codigo)
        if av:
            self._log(f"  ✓ Avaliação encontrada: Código={av.get('REG')}, Doc={av.get('documento')}, Status={av.get('STATUS')}")
            return av
        self._log(f"  ✗ Nenhuma avaliação encontrada com Código '{codigo}'")
        return None

    def _processar_matricula(self, matricula: str, imagens: list[str]):
        """Processa uma matrícula completa."""
        self._log(f"\n{'='*60}")
        self._log(f"PROCESSANDO MATRÍCULA: {matricula}")
        self._log(f"{'='*60}")

        pasta_imgs = Path(self.config.pasta_comparativos) / matricula

        numeros_imagens = []
        mapa_imagens = {}
        for img in imagens:
            stem = Path(img).stem
            try:
                num = int(re.sub(r'\D', '', stem))
                numeros_imagens.append(num)
                mapa_imagens[num] = str(pasta_imgs / img)
            except ValueError:
                self._log(f"  AVISO: Imagem '{img}' ignorada (nome não numérico)")

        if not numeros_imagens:
            self._log("  Nenhuma imagem válida encontrada")
            return

        try:
            comparativos, liq_forcada, metodo, justo_forcada, matricula_sistema = self.planilha.ler_dados_matricula(
                matricula, numeros_imagens
            )
        except Exception as e:
            self._log(f"  ERRO ao ler planilha: {e}")
            return

        for comp in comparativos:
            comp.imagem_path = mapa_imagens.get(comp.numero, "")

        self._log(f"  Planilha: {len(comparativos)} comparativos lidos")
        if liq_forcada:
            self._log(f"  Liq. Forçada: {liq_forcada}")

        self._log(f"  Buscando no sistema pelo código: {matricula_sistema}")
        avaliacao = self._encontrar_avaliacao_por_codigo(matricula_sistema)
        if not avaliacao:
            self._log("  PULANDO matrícula - avaliação não encontrada")
            return

        uuid = avaliacao.get("UUID") or avaliacao.get("uuid")
        if not uuid:
            self._log("  ERRO: UUID não encontrado na avaliação")
            return

        self._log(f"  UUID: {uuid}")

        if self.callback_progresso:
            self.callback_progresso(f"__UUID__:{uuid}")

        def _enviar_comparativo(args):
            """Envia um comparativo — roda em thread separada."""
            idx, comp, total = args
            if self._cancelado:
                return idx, comp.numero, False, "CANCELADO"

            if not comp.imagem_path or not os.path.exists(comp.imagem_path):
                return idx, comp.numero, False, "arquivo não encontrado"

            try:
                resultado = self.api.adicionar_comparativo(
                    uuid=uuid,
                    localidade=comp.localidade,
                    unidade=comp.unidade,
                    area=comp.area,
                    valor=comp.valor,
                    fonte=comp.fonte,
                    imagem_path=comp.imagem_path,
                )

                if resultado.get("status") == "sucesso":
                    reg = resultado.get("dados")
                    if self.config.excluir_imagens:
                        try:
                            os.remove(comp.imagem_path)
                        except OSError:
                            pass
                    return idx, comp.numero, True, f"REG: {reg}"
                else:
                    return idx, comp.numero, False, f"Erro: {resultado}"

            except requests.exceptions.HTTPError as e:
                return idx, comp.numero, False, f"HTTP {e.response.status_code}: {e.response.text[:200]}"
            except Exception as e:
                return idx, comp.numero, False, str(e)

        tarefas = []
        for i, comp in enumerate(comparativos, 1):
            if not comp.imagem_path or not os.path.exists(comp.imagem_path):
                self._log(f"  [{i}/{len(comparativos)}] Imagem {comp.numero}: arquivo não encontrado")
                continue
            self._log(f"  [{i}/{len(comparativos)}] Comparativo {comp.numero}: "
                      f"{comp.localidade} | {comp.area} {comp.unidade} | R$ {comp.valor}")
            tarefas.append((i, comp, len(comparativos)))

        if tarefas:
            max_workers = min(3, len(tarefas))
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(_enviar_comparativo, t): t for t in tarefas}
                for future in as_completed(futures):
                    idx, num, ok, msg = future.result()
                    if ok:
                        self._log(f"    ✓ Comparativo {num} adicionado ({msg})")
                    else:
                        self._log(f"    ✗ Comparativo {num}: {msg}")

            if self._cancelado:
                self._log("  CANCELADO pelo usuário")
                return

        if liq_forcada or metodo or justo_forcada:
            self._log(f"\n  Atualizando campos da avaliação...")
            if liq_forcada:
                self._log(f"    PERCENTFORCADA: {liq_forcada.replace('%','').strip()}%")
            if metodo:
                self._log(f"    METODO: {metodo}")
            if justo_forcada:
                self._log(f"    JUSTOFORCADA: {justo_forcada}")
            try:
                campos_editar = {}
                if liq_forcada:
                    campos_editar["PERCENTFORCADA"] = liq_forcada.replace("%", "").replace(",", ".").strip()
                if metodo:
                    campos_editar["METODO"] = metodo
                if justo_forcada:
                    campos_editar["JUSTOFORCADA"] = justo_forcada
                validade = self.callback_validade() if self.callback_validade else ""
                if validade:
                    campos_editar["VALIDADELAUDO"] = validade
                    self._log(f"    VALIDADELAUDO: {validade}")

                resultado = self.api.editar_avaliacao(uuid, **campos_editar)

                if resultado.get("status") == "sucesso":
                    self._log(f"    ✓ Campos atualizados com sucesso")
                else:
                    self._log(f"    ✗ Erro ao atualizar: {resultado}")

            except requests.exceptions.HTTPError as e:
                self._log(f"    ✗ Erro HTTP: {e.response.status_code} - {e.response.text[:200]}")
            except Exception as e:
                self._log(f"    ✗ Erro: {e}")

        self._log(f"\n  Buscando grupos de vistoria na planilha...")
        try:
            grupos_planilha = self.planilha.ler_grupos_vistoria(matricula)
        except Exception as e:
            self._log(f"  ERRO ao ler grupos da planilha: {e}")
            grupos_planilha = []

        if grupos_planilha:
            self._log(f"  Buscando grupos no sistema...")
            try:
                grupos_sistema = self.api.buscar_grupos_vistoria(uuid)
                self._log(f"  {len(grupos_sistema)} grupo(s) encontrado(s) no sistema")

                for gp in grupos_planilha:
                    nome_planilha = gp["nome"].strip().lower()
                    valor_planilha = gp["valor"].strip().replace(".", "").replace(",", ".")

                    def _buscar_grupo(candidatos, nome_alvo, exato=True):
                        com_s, com_n = None, None
                        for gs in candidatos:
                            nome_gs = (gs.get("NOME") or "").strip().lower()
                            ok = (nome_gs == nome_alvo) if exato else (nome_alvo in nome_gs or nome_gs in nome_alvo)
                            if ok:
                                if gs.get("APLICAAREA") == "S" and com_s is None:
                                    com_s = gs
                                elif gs.get("APLICAAREA") != "S" and com_n is None:
                                    com_n = gs
                        return com_s or com_n

                    match = _buscar_grupo(grupos_sistema, nome_planilha, exato=True)

                    if not match:
                        match = _buscar_grupo(grupos_sistema, nome_planilha, exato=False)

                    if match:
                        reg_grupo = str(match["REG"])
                        nome_real = (match.get("NOME") or "").strip()
                        self._log(f"  ✓ Match: '{gp['nome']}' → REG {reg_grupo}")
                        try:
                            resultado = self.api.salvar_grupo_imovel(
                                uuid, reg_grupo,
                                NOME=nome_real,
                                VALORUNIDADE=valor_planilha,
                                CONSTRUCAO=match.get("CONSTRUCAO", "N"),
                                AVERBADO=match.get("AVERBADO", "S"),
                                AREA=match.get("AREA") or "",
                                APLICAAREA=match.get("APLICAAREA", "N"),
                                TIPOMEDIDA=match.get("TIPOMEDIDA", ""),
                                OBS=match.get("OBS") or "",
                                ORDEM=match.get("ORDEM", 1),
                            )
                            if resultado.get("status") == "sucesso":
                                self._log(f"    ✓ VALORUNIDADE atualizado → {valor_planilha}")
                            else:
                                self._log(f"    ✗ Erro: {resultado}")
                        except requests.exceptions.HTTPError as e:
                            self._log(f"    ✗ Erro HTTP: {e.response.status_code} - {e.response.text[:200]}")
                        except Exception as e:
                            self._log(f"    ✗ Erro: {e}")
                    else:
                        self._log(f"  ✗ Grupo '{gp['nome']}' não encontrado no sistema")

            except requests.exceptions.HTTPError as e:
                self._log(f"  ERRO ao buscar grupos do sistema: {e.response.status_code} - {e.response.text[:200]}")
            except Exception as e:
                self._log(f"  ERRO ao buscar grupos do sistema: {e}")
        else:
            self._log(f"  Nenhum grupo de vistoria encontrado na planilha, pulando.")

        if not self.config.gerar_laudo:
            self._log(f"\n  Geração de laudo desativada, pulando...")
            self._log(f"\n  Matrícula {matricula} finalizada!")
            return

        self._log(f"\n  Gerando laudo PDF...")
        try:
            logo_url = f"{SISTEMAS[self.config.sistema]['origin']}/empresas/{SISTEMAS[self.config.sistema]['origin'].split('//')[1].split('.')[0]}/logoFundoBranco.png"
            logo_resp = self.api.session.get(logo_url, timeout=30)
            logo_resp.raise_for_status()
            logo_bytes = logo_resp.content
            self._log(f"  Logo baixada: {len(logo_bytes)} bytes")

            pdf_bytes = self.api.gerar_laudo(uuid, logo_bytes)

            if pdf_bytes:
                nome_laudo = self.api.buscar_nome_laudo(uuid)
                nome_arquivo = re.sub(r'[<>:"/\\|?*]', '_', nome_laudo) + ".pdf"

                pasta_laudos = _base_dir() / "Matrícula"
                pasta_laudos.mkdir(exist_ok=True)
                caminho_pdf = pasta_laudos / nome_arquivo

                with open(caminho_pdf, "wb") as f:
                    f.write(pdf_bytes)

                self._log(f"  ✓ Laudo salvo: {caminho_pdf}")

                escolha = self.mapa_assinaturas.get(matricula)
                if escolha is None:
                    self._log(f"  Assinatura pulada para esta matrícula")
                else:
                    pfx_pessoa, pfx_empresa = escolha
                    self._log(f"\n  Assinando com: {pfx_pessoa} + {pfx_empresa}")
                    ok = assinar_pdf(caminho_pdf, pfx_pessoa, pfx_empresa, self.logger)
                    if ok:
                        self._log(f"  ✓ Laudo assinado com sucesso")
                    else:
                        self._log(f"  ✗ Assinatura falhou — laudo salvo sem assinatura")
            else:
                self._log(f"  ✗ Laudo gerado veio vazio")

        except requests.exceptions.HTTPError as e:
            self._log(f"  ✗ Erro HTTP ao gerar laudo: {e.response.status_code} - {e.response.text[:200]}")
        except Exception as e:
            self._log(f"  ✗ Erro ao gerar laudo: {e}")

        self._log(f"\n  Matrícula {matricula} finalizada!")

    def executar(self):
        """Executa o robô completo."""
        try:
            self._log("=" * 60)
            self._log("ROBÔ COMPARATIVO - INICIANDO")
            self._log("=" * 60)

            sistema_config = SISTEMAS[self.config.sistema]

            subpastas = self._listar_subpastas()
            if not subpastas:
                self._log("Nenhuma subpasta com imagens encontrada!")
                return

            self._log(f"\nSubpastas encontradas:")
            for nome, imgs in subpastas:
                self._log(f"  📁 {nome} ({len(imgs)} imagens)")

            self._log(f"\nConectando ao Google Sheets...")
            try:
                self.planilha = PlanilhaClient(
                    self.config.credentials_path,
                    self.config.planilha_id,
                    self.logger,
                )
            except Exception as e:
                self._log(f"ERRO ao conectar planilha: {e}")
                return

            self._log(f"\nAutenticando na API ({self.config.sistema})...")
            jwt = self._obter_jwt(sistema_config)
            self.api = MultiAAPI(sistema_config, jwt, self.logger)

            self._log("Testando conexão com a API...")
            try:
                teste = self.api.buscar_avaliacoes("9999999", page_size=1)
                if teste.get("status") == "sucesso":
                    self._log("  ✓ API conectada com sucesso!")
                else:
                    self._log(f"  ✗ API retornou: {teste}")
                    return
            except Exception as e:
                self._log(f"  ✗ Falha na conexão: {e}")
                return

            total = len(subpastas)
            for idx, (matricula, imagens) in enumerate(subpastas, 1):
                if self._cancelado:
                    self._log("\nEXECUÇÃO CANCELADA")
                    break

                self.callback_progresso(f"PROGRESSO: {idx}/{total}")
                if self.planilha:
                    self.planilha._cache_abas.clear()
                self._processar_matricula(matricula, imagens)

            self._log("\n" + "=" * 60)
            self._log("EXECUÇÃO FINALIZADA")
            self._log("=" * 60)

        except Exception as e:
            self._log(f"\nERRO FATAL: {e}")
            import traceback
            self._log(traceback.format_exc())


# ============================================================
# INTERFACE GRÁFICA (CustomTkinter)
# ============================================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

C = {
    "bg":        "#0B0F1A",
    "surface":   "#111827",
    "surface2":  "#1A2236",
    "border":    "#1E2D45",
    "text":      "#F0F4FF",
    "muted":     "#6B7FA3",
    "accent":    "#F5C400",
    "accent_hov":"#D4A900",
    "accent_dim":"#2A2400",
    "blue":      "#1D4ED8",
    "blue_hov":  "#1E40AF",
    "blue_lite":  "#1E3A5F",
    "neutral":   "#6B7FA3",
    "neut_bg":   "#1A2236",
    "neut_brd":  "#1E2D45",
    "ok":        "#22C55E",
    "err":       "#EF4444",
    "warn":      "#F59E0B",
    "log_bg":    "#070B14",
    "log_fg":    "#CBD5E1",
    "log_ok":    "#4ADE80",
    "log_err":   "#F87171",
    "log_warn":  "#FBBF24",
    "log_info":  "#60A5FA",
    "log_muted": "#334155",
}

FONT_TITLE  = ("Segoe UI", 18, "bold")
FONT_LABEL  = ("Segoe UI", 13)
FONT_SMALL  = ("Segoe UI", 11)
FONT_BTN    = ("Segoe UI", 13, "bold")
FONT_INPUT  = ("Segoe UI", 12)
FONT_LOG    = ("Consolas", 10)
FONT_MONO   = ("Cascadia Code", 10)


class App:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("MultiA — Robô Comparativo")
        self.root.geometry("820x640")
        self.root.resizable(True, True)
        self.root.minsize(740, 560)
        self.root.configure(fg_color=C["bg"])
        self.root.configure(bg=C["bg"])

        self.config = ConfigData()
        self.robo: Optional[RoboComparativo] = None
        self.executando = False

        self._setup_logger()
        self._build_ui()
        self._carregar_config()

    def _setup_logger(self):
        self.logger = logging.getLogger("RoboComparativo")
        self.logger.setLevel(logging.DEBUG)
        handler = logging.StreamHandler(sys.stdout)
        handler.setLevel(logging.DEBUG)
        formatter = logging.Formatter("%(asctime)s - %(message)s", datefmt="%H:%M:%S")
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)

    def _build_ui(self):
        root = self.root

        header = ctk.CTkFrame(root, fg_color=C["surface"],
                              corner_radius=0, height=48, border_width=0)
        header.pack(fill="x")
        header.pack_propagate(False)

        title_frame = ctk.CTkFrame(header, fg_color="transparent")
        title_frame.pack(side="left", padx=18, pady=8)

        ctk.CTkLabel(title_frame, text="MultiA",
                     font=("Segoe UI", 15, "bold"),
                     text_color=C["accent"]).pack(side="left")
        ctk.CTkLabel(title_frame, text="  Robô Comparativo",
                     font=("Segoe UI", 11),
                     text_color=C["muted"]).pack(side="left")

        self._dot_canvas = tk.Canvas(header, width=10, height=10,
                                      bg=C["surface"], highlightthickness=0)
        self._dot_canvas.pack(side="right", padx=18, pady=19)
        self._dot = self._dot_canvas.create_oval(1, 1, 9, 9, fill=C["muted"], outline="")

        ctk.CTkFrame(root, fg_color=C["border"], height=1, corner_radius=0).pack(fill="x")

        body = ctk.CTkFrame(root, fg_color=C["bg"])
        body.pack(fill="both", expand=True, padx=20, pady=12)

        body.columnconfigure(0, weight=2, minsize=320)
        body.columnconfigure(1, weight=3, minsize=340)
        body.rowconfigure(0, weight=1)

        left = ctk.CTkFrame(body, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self._section_title(left, "Sistema")
        card_sis = self._card(left)

        self.var_sistema = ctk.StringVar(value="MultiA Mais")
        radio_row = ctk.CTkFrame(card_sis, fg_color="transparent")
        radio_row.pack(fill="x")
        for nome in SISTEMAS:
            ctk.CTkRadioButton(
                radio_row, text=nome,
                variable=self.var_sistema, value=nome,
                font=("Segoe UI", 13), text_color=C["text"],
                fg_color=C["accent"], hover_color=C["accent_hov"],
                border_color=C["border"],
            ).pack(side="left", padx=(0, 20))

        self._section_title(left, "ID da Planilha")
        card_pl = self._card(left)
        self.entry_planilha = ctk.CTkEntry(
            card_pl, font=FONT_INPUT,
            fg_color=C["surface2"], border_color=C["border"],
            text_color=C["text"], border_width=1, corner_radius=8, height=44,
            placeholder_text="Cole o ID aqui...",
            placeholder_text_color=C["muted"],
        )
        self.entry_planilha.pack(fill="x")
        self.entry_planilha.insert(0, self.config.planilha_id)

        self._section_title(left, "Credenciais Google")
        card_cr = self._card(left)
        row_cr = ctk.CTkFrame(card_cr, fg_color="transparent")
        row_cr.pack(fill="x")
        self.entry_creds = ctk.CTkEntry(
            row_cr, font=FONT_INPUT,
            fg_color=C["surface2"], border_color=C["border"],
            text_color=C["text"], border_width=1, corner_radius=8, height=44,
            placeholder_text="credentials.json",
            placeholder_text_color=C["muted"],
        )
        self.entry_creds.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(
            row_cr, text="📂", font=("Segoe UI", 10),
            fg_color=C["neut_bg"], text_color=C["text"],
            hover_color=C["blue_lite"],
            border_width=1, border_color=C["neut_brd"],
            corner_radius=8, height=44, width=54,
            command=self._selecionar_credentials,
        ).pack(side="left")

        self._section_title(left, "Pasta Comparativos")
        card_pa = self._card(left)
        row_pa = ctk.CTkFrame(card_pa, fg_color="transparent")
        row_pa.pack(fill="x")
        self.entry_pasta = ctk.CTkEntry(
            row_pa, font=FONT_INPUT,
            fg_color=C["surface2"], border_color=C["border"],
            text_color=C["text"], border_width=1, corner_radius=8, height=44,
            placeholder_text="Selecione a pasta...",
            placeholder_text_color=C["muted"],
        )
        self.entry_pasta.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(
            row_pa, text="📂", font=("Segoe UI", 10),
            fg_color=C["neut_bg"], text_color=C["text"],
            hover_color=C["blue_lite"],
            border_width=1, border_color=C["neut_brd"],
            corner_radius=8, height=44, width=54,
            command=self._selecionar_pasta,
        ).pack(side="left")

        self.lbl_subpastas = ctk.CTkLabel(
            card_pa, text="Nenhuma pasta selecionada",
            font=("Segoe UI", 11), text_color=C["muted"], anchor="w",
        )
        self.lbl_subpastas.pack(anchor="w", pady=(6, 0))

        opt_btn_row = ctk.CTkFrame(left, fg_color="transparent")
        opt_btn_row.pack(fill="x", pady=(8, 0))

        self.var_excluir = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            opt_btn_row,
            text="Excluir imagens após envio",
            variable=self.var_excluir,
            font=("Segoe UI", 12), text_color=C["muted"],
            fg_color=C["accent"], hover_color=C["accent_hov"],
            border_color=C["border"], checkmark_color="#000000",
        ).pack(side="left")

        self.var_gerar_laudo = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            opt_btn_row,
            text="Gerar e assinar laudo",
            variable=self.var_gerar_laudo,
            font=("Segoe UI", 12), text_color=C["muted"],
            fg_color=C["accent"], hover_color=C["accent_hov"],
            border_color=C["border"], checkmark_color="#000000",
        ).pack(side="left", padx=(16, 0))

        btn_row = ctk.CTkFrame(left, fg_color="transparent")
        btn_row.pack(fill="x", pady=(10, 0))

        self.btn_executar = ctk.CTkButton(
            btn_row, text="▶  EXECUTAR",
            font=FONT_BTN,
            fg_color=C["accent"], hover_color=C["accent_hov"],
            text_color="#0B0F1A", corner_radius=10, height=52,
            command=self._executar,
        )
        self.btn_executar.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.btn_cancelar = ctk.CTkButton(
            btn_row, text="■  PARAR",
            font=FONT_BTN,
            fg_color=C["neut_bg"], hover_color=C["blue_lite"],
            text_color=C["muted"],
            border_width=1, border_color=C["neut_brd"],
            corner_radius=10, height=52, state="disabled",
            command=self._cancelar,
        )
        self.btn_cancelar.pack(side="left", fill="x", expand=True)

        status_card = ctk.CTkFrame(
            left, fg_color=C["surface"],
            border_color=C["border"], border_width=1, corner_radius=8,
        )
        status_card.pack(fill="x", pady=(10, 0))

        status_inner = ctk.CTkFrame(status_card, fg_color="transparent")
        status_inner.pack(fill="x", padx=12, pady=8)

        ctk.CTkLabel(status_inner, text="STATUS",
                     font=("Segoe UI", 7, "bold"),
                     text_color=C["muted"]).pack(side="left")
        ctk.CTkFrame(status_inner, fg_color=C["border"],
                     width=1, height=14, corner_radius=0).pack(side="left", padx=10)

        self.progress_var = ctk.StringVar(value="Aguardando...")
        self.lbl_status = ctk.CTkLabel(
            status_inner, textvariable=self.progress_var,
            font=("Segoe UI", 9, "bold"),
            text_color=C["accent"], anchor="w",
        )
        self.lbl_status.pack(side="left", fill="x", expand=True)

        right = ctk.CTkFrame(body, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")

        self._section_title(right, "Log de Execução")

        log_card = ctk.CTkFrame(
            right, fg_color=C["log_bg"],
            corner_radius=10, border_color=C["border"], border_width=1,
        )
        log_card.pack(fill="both", expand=True)

        log_hdr = ctk.CTkFrame(log_card, fg_color=C["surface"],
                                corner_radius=0, height=28)
        log_hdr.pack(fill="x")
        log_hdr.pack_propagate(False)

        dots = ctk.CTkFrame(log_hdr, fg_color="transparent")
        dots.pack(side="left", padx=10, pady=7)
        for col in ("#FF5F57", "#FFBD2E", "#28C840"):
            ctk.CTkFrame(dots, fg_color=col, width=8, height=8,
                         corner_radius=4).pack(side="left", padx=2)

        ctk.CTkLabel(log_hdr, text="console",
                     font=("Segoe UI", 8), text_color=C["muted"]).pack(side="left", padx=4)

        self.log_text = scrolledtext.ScrolledText(
            log_card,
            bg=C["log_bg"], fg=C["log_fg"],
            font=FONT_MONO if self._font_exists("Cascadia Code") else FONT_LOG,
            insertbackground=C["log_fg"],
            wrap=tk.WORD, state=tk.DISABLED,
            relief=tk.FLAT, bd=0,
            padx=14, pady=10,
            selectbackground="#2D3748",
            height=6,
        )
        self.log_text.pack(fill="both", expand=True)

        self.log_text.tag_config("ok",    foreground=C["log_ok"])
        self.log_text.tag_config("err",   foreground=C["log_err"])
        self.log_text.tag_config("warn",  foreground=C["log_warn"])
        self.log_text.tag_config("info",  foreground=C["log_info"])
        self.log_text.tag_config("muted", foreground=C["log_muted"])

        bottom_right = ctk.CTkFrame(right, fg_color="transparent")
        bottom_right.pack(fill="x", pady=(8, 0))

        self._section_title(bottom_right, "Validade do Laudo (dias)")
        card_val = self._card(bottom_right)

        val_row = ctk.CTkFrame(card_val, fg_color="transparent")
        val_row.pack(fill="x")

        self.entry_validade = ctk.CTkEntry(
            val_row, font=("Segoe UI", 13, "bold"),
            fg_color=C["surface2"], border_color=C["border"],
            text_color=C["accent"],
            border_width=1, corner_radius=8, height=44,
            placeholder_text="Ex: 12",
            placeholder_text_color=C["muted"],
            justify="center",
        )
        self.entry_validade.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.entry_validade.bind("<Return>",   lambda e: self._salvar_validade())
        self.entry_validade.bind("<FocusOut>", lambda e: self._salvar_validade())
        self.entry_validade.bind("<KeyRelease>", lambda e: self._salvar_validade())

        self.lbl_val_status = ctk.CTkLabel(
            val_row, text="", font=("Segoe UI", 10),
            text_color=C["ok"], width=80,
        )
        self.lbl_val_status.pack(side="left")

        self._section_title(bottom_right, "Configurações")
        self.btn_salvar = ctk.CTkButton(
            bottom_right, text="💾  Salvar Config",
            font=FONT_BTN,
            fg_color=C["neut_bg"], hover_color=C["blue_lite"],
            text_color=C["text"],
            border_width=1, border_color=C["neut_brd"],
            corner_radius=10, height=42,
            command=self._salvar_config_manual,
        )
        self.btn_salvar.pack(fill="x")

        self._dot_phase = 0
        self._animate_dot()

    def _section_title(self, parent, text, top_pad=6):
        ctk.CTkLabel(
            parent, text=text,
            font=("Segoe UI", 11, "bold"),
            text_color=C["muted"], anchor="w",
        ).pack(anchor="w", pady=(top_pad, 3))

    def _card(self, parent):
        card = ctk.CTkFrame(
            parent,
            fg_color=C["surface"],
            border_color=C["border"],
            border_width=1,
            corner_radius=8,
        )
        card.pack(fill="x", pady=(0, 4))

        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=10)
        return inner

    def _animate_dot(self):
        """Pulsa o dot no header enquanto executando."""
        if self.executando:
            import math
            self._dot_phase = (self._dot_phase + 0.12) % (2 * math.pi)
            brightness = int(180 + 75 * math.sin(self._dot_phase))
            r, g = brightness, int(brightness * 0.78)
            color = f"#{min(r,255):02x}{min(g,255):02x}00"
            self._dot_canvas.itemconfig(self._dot, fill=color)
        else:
            self._dot_canvas.itemconfig(self._dot, fill=C["muted"])
        self.root.after(16, self._animate_dot)

    def _font_exists(self, name):
        try:
            import tkinter.font as tkfont
            return name in tkfont.families()
        except Exception:
            return False

    def _log_ui(self, msg: str):
        def _append():
            self.log_text.configure(state=tk.NORMAL)

            if "✓" in msg:
                tag = "ok"
            elif "✗" in msg or "ERRO" in msg or "ERROR" in msg:
                tag = "err"
            elif "AVISO" in msg or "CANCELAD" in msg:
                tag = "warn"
            elif msg.startswith("=") or msg.startswith("ROBÔ") or msg.startswith("EXECUÇÃO"):
                tag = "info"
            elif not msg.strip():
                tag = "muted"
            else:
                tag = None

            self.log_text.insert(tk.END, msg + "\n", tag if tag else "")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)

            if msg.startswith("PROGRESSO:"):
                partes = msg.replace("PROGRESSO:", "").strip().split("/")
                if len(partes) == 2:
                    self.progress_var.set(f"Processando {partes[0].strip()} de {partes[1].strip()} matrículas...")
                else:
                    self.progress_var.set(msg)
            elif msg.startswith("__UUID__:"):
                uuid = msg.replace("__UUID__:", "").strip()
                self._ultimo_uuid = uuid
                if hasattr(self, "_validade_pendente") and self._validade_pendente:
                    if hasattr(self, "api") and self.robo and self.robo.api:
                        self.api = self.robo.api
                    if hasattr(self, "api"):
                        self._enviar_validade(uuid, self._validade_pendente)

        self.root.after(16, _append)

    def _selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta Comparativos")
        if pasta:
            self.entry_pasta.delete(0, tk.END)
            self.entry_pasta.insert(0, pasta)
            self._atualizar_subpastas(pasta)

    def _selecionar_credentials(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione credentials.json",
            filetypes=[("JSON", "*.json")],
        )
        if arquivo:
            self.entry_creds.delete(0, tk.END)
            self.entry_creds.insert(0, arquivo)

    def _atualizar_subpastas(self, pasta: str):
        try:
            p = Path(pasta)
            subs = []
            for d in sorted(p.iterdir()):
                if d.is_dir():
                    imgs = [f for f in d.iterdir() if f.suffix.lower() in IMAGE_EXTENSIONS]
                    if imgs:
                        subs.append(f"{d.name} ({len(imgs)} imgs)")

            if subs:
                self.lbl_subpastas.configure(
                    text="✓  " + "   ·   ".join(subs),
                    text_color=C["ok"])
            else:
                self.lbl_subpastas.configure(
                    text="Nenhuma subpasta com imagens encontrada",
                    text_color="#991B1B")
        except Exception:
            pass

    def _confirmar_dialog(self, msg: str) -> bool:
        result = [False]
        def _ask():
            result[0] = messagebox.askyesno("Confirmar", msg)
        self.root.after(0, _ask)
        time.sleep(0.5)
        return result[0]

    def _executar(self):
        if self.executando:
            return

        pasta = self.entry_pasta.get().strip()
        if not pasta or not os.path.isdir(pasta):
            messagebox.showerror("Erro", "Selecione uma pasta válida")
            return

        creds = self.entry_creds.get().strip()
        if not creds or not os.path.isfile(creds):
            messagebox.showerror("Erro", "Selecione o arquivo credentials.json")
            return

        self.config.sistema = self.var_sistema.get()
        self.config.pasta_comparativos = pasta
        self.config.credentials_path = creds
        self.config.planilha_id = self.entry_planilha.get().strip()
        self.config.excluir_imagens = self.var_excluir.get()
        self.config.gerar_laudo = self.var_gerar_laudo.get()

        import re as _re
        from pathlib import Path as _Path
        matriculas = []
        try:
            for d in sorted(_Path(pasta).iterdir()):
                if d.is_dir():
                    imgs = [f for f in d.iterdir() if f.suffix.lower() in IMAGE_EXTENSIONS]
                    if imgs:
                        matriculas.append(d.name)
        except Exception:
            pass

        mapa_assinaturas = {}
        if self.config.gerar_laudo:
            for matricula in matriculas:
                escolha = dialogo_assinaturas(matricula, self.config.sistema)
                mapa_assinaturas[matricula] = escolha

        self.executando = True
        self.btn_executar.configure(state="disabled", fg_color=C["accent_hov"])
        self.btn_cancelar.configure(state="normal", fg_color="#8B1A1A",
                                     text_color="#FFAAAA", hover_color="#B91C1C")
        self.progress_var.set("Iniciando...")

        self._salvar_config()

        def _run():
            try:
                self.robo = RoboComparativo(
                    config=self.config,
                    logger=self.logger,
                    callback_progresso=self._log_ui,
                    callback_confirmar=self._confirmar_dialog,
                    callback_validade=lambda: self.entry_validade.get().strip(),
                    mapa_assinaturas=mapa_assinaturas,
                )
                self.robo.executar()
            finally:
                self.root.after(0, self._finalizar_execucao)

        Thread(target=_run, daemon=True).start()

    def _cancelar(self):
        if self.robo:
            self.robo.cancelar()
            self._log_ui("Cancelamento solicitado...")

    def _finalizar_execucao(self):
        self.executando = False
        self.btn_executar.configure(state="normal", fg_color=C["accent"])
        self.btn_cancelar.configure(state="disabled", fg_color=C["neut_bg"],
                                     text_color=C["muted"])
        self.progress_var.set("✓  Finalizado com sucesso")

    def _salvar_config(self):
        config_path = _base_dir() / "config.json"
        data = {
            "sistema": self.config.sistema,
            "planilha_id": self.config.planilha_id,
            "credentials_path": self.config.credentials_path,
            "pasta": self.config.pasta_comparativos,
            "excluir_imagens": self.config.excluir_imagens,
            "gerar_laudo": self.config.gerar_laudo,
        }
        with open(config_path, "w") as f:
            json.dump(data, f, indent=2)

    def _salvar_config_manual(self):
        """Salva config atual no JSON e dá feedback visual."""
        self.config.sistema = self.var_sistema.get()
        self.config.planilha_id = self.entry_planilha.get().strip()
        self.config.credentials_path = self.entry_creds.get().strip()
        self.config.pasta_comparativos = self.entry_pasta.get().strip()
        self.config.excluir_imagens = self.var_excluir.get()
        self.config.gerar_laudo = self.var_gerar_laudo.get()
        self._salvar_config()

        self.btn_salvar.configure(text="✓  Config Salva!", fg_color=C["ok"],
                                        text_color="#000000", state="disabled")
        self.root.after(1500, lambda: self.btn_salvar.configure(
            text="💾  Salvar Config", fg_color=C["neut_bg"],
            text_color=C["text"], state="normal"
        ))

    def _salvar_validade(self):
        """Envia VALIDADELAUDO para a API ao confirmar o campo."""
        valor = self.entry_validade.get().strip()
        if not valor:
            return

        if not valor.isdigit():
            self.lbl_val_status.configure(text="✗ inválido", text_color=C["err"])
            self.root.after(2000, lambda: self.lbl_val_status.configure(text=""))
            return

        self._validade_pendente = valor

        if hasattr(self, "_ultimo_uuid") and self._ultimo_uuid:
            self._enviar_validade(self._ultimo_uuid, valor)
        else:
            self.lbl_val_status.configure(text="💾 salvo", text_color=C["muted"])
            self.root.after(2000, lambda: self.lbl_val_status.configure(text=""))

    def _enviar_validade(self, uuid: str, valor: str):
        """Envia VALIDADELAUDO para a API em thread separada."""
        def _run():
            try:
                resultado = self.api.editar_avaliacao(uuid, VALIDADELAUDO=valor)
                if resultado.get("status") == "sucesso":
                    self.root.after(0, lambda: self.lbl_val_status.configure(
                        text="✓ salvo", text_color=C["ok"]))
                else:
                    self.root.after(0, lambda: self.lbl_val_status.configure(
                        text="✗ erro", text_color=C["err"]))
            except Exception:
                self.root.after(0, lambda: self.lbl_val_status.configure(
                    text="✗ erro", text_color=C["err"]))
            self.root.after(2000, lambda: self.lbl_val_status.configure(text=""))

        Thread(target=_run, daemon=True).start()

    def _carregar_config(self):
        config_path = _base_dir() / "config.json"
        if config_path.exists():
            try:
                with open(config_path) as f:
                    data = json.load(f)
                self.var_sistema.set(data.get("sistema", "MultiA Mais"))
                if data.get("planilha_id"):
                    self.entry_planilha.delete(0, tk.END)
                    self.entry_planilha.insert(0, data["planilha_id"])
                if data.get("credentials_path"):
                    self.entry_creds.insert(0, data["credentials_path"])
                if data.get("pasta") and os.path.isdir(data["pasta"]):
                    self.entry_pasta.insert(0, data["pasta"])
                    self._atualizar_subpastas(data["pasta"])
                self.var_excluir.set(data.get("excluir_imagens", True))
                self.var_gerar_laudo.set(data.get("gerar_laudo", True))
            except Exception:
                pass

    def run(self):
        self.root.mainloop()

# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    app = App()
    app.run()
