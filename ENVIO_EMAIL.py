# -*- coding: utf-8 -*-
"""
Automação: Power BI (Service) -> Exporta PDFs -> Aplica regras de páginas -> Mescla -> Envia por e-mail (Outlook)
Agendamento diário com 'schedule'.

Este arquivo está pronto para ser compartilhado: todos os caminhos e links foram
trocados por placeholders como [seu caminho] e [seu link].

Como usar:
1) Substitua [seu caminho] por diretórios do seu ambiente (downloads, lista de e-mails, msedgedriver).
2) Substitua [seu link] pelas URLs dos relatórios publicados no Power BI que você deseja exportar.
3) (Opcional) Crie um arquivo .env para configurar as variáveis sem editar o código.
4) Rode o script uma vez na semana para realizar a autenticação SSO no Edge.
5) Deixe o script rodando/schedule no servidor/PC para executar diariamente.

Dependências principais:
- selenium, PyPDF2, python-dotenv (opcional), pywin32, schedule

Autor: Gabriel Tedesque (versão para compartilhamento)
"""

import os
import time
import datetime
import shutil
import subprocess
import logging
from logging.handlers import RotatingFileHandler
from dataclasses import dataclass
from typing import Optional, List, Tuple

# -------------------------- Dependências de terceiros -------------------------
# Observação: instale via pip (ex.: pip install selenium PyPDF2 python-dotenv pywin32 schedule)
import schedule
import win32com.client as win32

from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from PyPDF2 import PdfMerger, PdfReader, PdfWriter

# ------------------------------- Suporte a .env -------------------------------
# O uso de .env é opcional. Se existir, permite configurar variáveis sem editar o código.
try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    # Se python-dotenv não estiver instalado, o script segue com defaults.
    pass

# ============================ CONFIGURAÇÕES GERAIS ============================
# ATENÇÃO: Troque [seu caminho] pelos diretórios do seu ambiente.
DEFAULT_DOWNLOAD_DIR = r"[seu caminho]\Apresentacao"     # Pasta base onde serão criadas subpastas por data para salvar PDFs
DEFAULT_EMAIL_LIST_DIR = r"[seu caminho]\ListaEmails"    # Pasta contendo um .txt/.csv com e-mails (um por linha)
DEFAULT_MSEDGEDRIVER_PATH = r"[seu caminho]\msedgedriver.exe"  # Caminho completo para o msedgedriver.exe
DEFAULT_RUN_HOUR_MINUTE = "15:50"                        # Horário diário (HH:MM) para rodar o job

# As variáveis abaixo podem ser sobrescritas via .env (se existir)
DOWNLOAD_DIR = os.getenv("DOWNLOAD_DIR", DEFAULT_DOWNLOAD_DIR)
EMAIL_LIST_DIR = os.getenv("EMAIL_LIST_DIR", DEFAULT_EMAIL_LIST_DIR)
MSEDGEDRIVER_PATH = os.getenv("MSEDGEDRIVER_PATH", DEFAULT_MSEDGEDRIVER_PATH)
RUN_HOUR_MINUTE = os.getenv("RUN_HOUR_MINUTE", DEFAULT_RUN_HOUR_MINUTE)

# Afinamentos de tempo (podem ser ajustados se o serviço estiver mais lento/rápido)
PRE_EXPORT_COOLDOWN_SEC = int(os.getenv("PRE_EXPORT_COOLDOWN_SEC", "1"))   # Pequena pausa antes de abrir o menu de exportar
IDLE_STABLE_SECONDS     = int(os.getenv("IDLE_STABLE_SECONDS", "5"))       # Quanto tempo precisa ficar "estável" sem loading
IDLE_TIMEOUT_SEC        = int(os.getenv("IDLE_TIMEOUT_SEC", "90"))         # Tempo máximo esperando estabilidade de um relatório
DOWNLOAD_TIMEOUT_SEC    = int(os.getenv("DOWNLOAD_TIMEOUT_SEC", "600"))    # Tempo máximo para um PDF finalizar download

# Estratégia: exportação "imediata" primeiro; se falhar, faz fallback esperando estabilidade
FORCE_EXPORT_IMMEDIATE  = os.getenv("FORCE_EXPORT_IMMEDIATE", "true").lower() == "true"
IMMEDIATE_TRIES         = int(os.getenv("IMMEDIATE_TRIES", "3"))           # Número de tentativas no modo imediato

# Data atual (usada para nomear pasta do dia e PDF final)
DATA_HOJE = datetime.datetime.now().strftime("%Y-%m-%d")

# Prepara pastas de saída
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
RUN_DIR = os.path.join(DOWNLOAD_DIR, DATA_HOJE)
os.makedirs(RUN_DIR, exist_ok=True)

# Caminho do PDF final (consolidado)
PDF_FINAL_PATH = os.path.join(RUN_DIR, f"IndicadoresDiarios_{DATA_HOJE}.pdf")

# ================================ LOGGING ====================================
# Log rotativo para depuração e auditoria: um arquivo por dia/execução
LOG_PATH = os.path.join(DOWNLOAD_DIR, "indicadores.log")

def setup_logging() -> logging.Logger:
    """
    Configura um logger com saída em arquivo (rotativo) e no console.
    """
    logger = logging.getLogger("indicadores")
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s", datefmt="%H:%M:%S")

    # Handler de arquivo rotativo
    fh = RotatingFileHandler(LOG_PATH, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
    fh.setFormatter(fmt)

    # Handler de console (stdout)
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)

    if not logger.handlers:
        logger.addHandler(fh)
        logger.addHandler(ch)
    return logger

log = setup_logging()

# =========================== MODELO DE RELATÓRIO ==============================
@dataclass
class Report:
    """
    Estrutura para parametrizar cada relatório:
    - name: Nome curto do relatório (apenas para arquivos/logs).
    - url: Link do relatório no Power BI Service (substituir por [seu link]).
    - extract_page: (opcional) extrair apenas esta página (1-based).
    - drop_pages: (opcional) remover páginas específicas (lista 1-based).
    - drop_last_pages: (opcional) remover as N últimas páginas.
    Obs.: Apenas uma das três regras acima deve ser usada por relatório.
    """
    name: str
    url: str
    extract_page: Optional[int] = None
    drop_pages: Optional[List[int]] = None
    drop_last_pages: Optional[int] = None

    def validate(self):
        """
        Evita configurações ambíguas: só pode haver UMA regra de página por relatório.
        """
        flags = sum([
            self.extract_page is not None,
            self.drop_pages is not None,
            self.drop_last_pages is not None
        ])
        if flags > 1:
            raise ValueError(
                f"Regras ambíguas em '{self.name}': use apenas uma entre "
                "extract_page | drop_pages | drop_last_pages."
            )

# =========================== LISTA DE RELATÓRIOS ==============================
# ATENÇÃO: Substitua [seu link] pela URL de cada relatório publicado.
REPORTS: List[Report] = [
    Report(
        name="Relatorio_A",
        url="[seu link]"  # Ex.: https://app.powerbi.com/groups/.../reports/.../ReportSection...
        # Sem regra de páginas (exporta tudo)
    ),
    Report(
        name="Relatorio_B",
        url="[seu link]",
        drop_pages=[4],  # Ex.: remove a página 4 (ajuste conforme necessidade)
    ),
    Report(
        name="Relatorio_C",
        url="[seu link]",
        extract_page=5,  # Ex.: extrai apenas a página 5
    ),
    Report(
        name="Relatorio_D",
        url="[seu link]",
        drop_last_pages=1,  # Ex.: remove a última página
    ),
    # Adicione/edite conforme sua realidade
]

# ============================== SELENIUM / EDGE ===============================
def setup_edge_driver(download_dir: str) -> webdriver.Edge:
    """
    Configura o Edge para:
    - Baixar PDFs automaticamente para 'download_dir' (sem prompt)
    - Desabilitar flag de automação visível
    - Iniciar maximizado

    Retorna uma instância de webdriver.Edge pronta para uso.
    """
    if not os.path.exists(MSEDGEDRIVER_PATH):
        raise FileNotFoundError(f"msedgedriver.exe não encontrado em:\n{MSEDGEDRIVER_PATH}")

    opts = EdgeOptions()
    opts.use_chromium = True
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--start-maximized")

    # Preferências para download automático de PDF
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,  # evita preview embutido
    }
    opts.add_experimental_option("prefs", prefs)

    service = EdgeService(executable_path=MSEDGEDRIVER_PATH)
    return webdriver.Edge(service=service, options=opts)

def switch_to_frame_with(driver, by, selector, timeout=20) -> bool:
    """
    Procura um elemento em todos os iframes da página e alterna para o iframe correto.
    Retorna True se encontrou; False caso contrário.
    """
    end = time.time() + timeout
    while time.time() < end:
        try:
            driver.switch_to.default_content()
            if driver.find_elements(by, selector):
                return True
        except Exception:
            pass
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        for fr in frames:
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(fr)
                if driver.find_elements(by, selector):
                    return True
            except Exception:
                continue
        time.sleep(0.3)
    return False

def click_anywhere(driver, by, selector, wait, must_be_clickable=True, timeout=25):
    """
    Clica em um elemento localizado pelo 'selector' (XPATH/CSS), mesmo se estiver
    dentro de iframes (usa 'switch_to_frame_with').
    """
    if not switch_to_frame_with(driver, by, selector, timeout=timeout):
        raise TimeoutError(f"Elemento não encontrado em nenhum iframe: {selector}")
    if must_be_clickable:
        el = wait.until(EC.element_to_be_clickable((by, selector)))
    else:
        el = wait.until(EC.presence_of_element_located((by, selector)))
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.05)
        el.click()
    except Exception:
        # Fallback: tenta clicar via JS direto no elemento
        driver.execute_script("arguments[0].click();", el)
    return el

def pbix_ready(driver, wait, extra_sleep=0.5):
    """
    Ajuda a garantir que a estrutura base do Power BI carregou (iframes presentes).
    """
    driver.switch_to.default_content()
    try:
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
    except Exception:
        pass
    time.sleep(extra_sleep)

def _has_loading_indicators(driver) -> bool:
    """
    Heurística para detectar indicadores de loading no Power BI.
    Ajuste os seletores conforme mudanças de layout.
    """
    sels = [
        "//*[@aria-busy='true']",
        "//*[contains(@class,'busy') or contains(@class,'loading') or contains(@class,'spinner')]",
        "//*[@role='progressbar']",
        "//div[contains(@class,'powerbi-loader')]",
        "//div[contains(@class,'waitSpinner')]",
    ]
    return any(driver.find_elements(By.XPATH, s) for s in sels)

def wait_report_idle(driver, timeout=IDLE_TIMEOUT_SEC, stable_required=IDLE_STABLE_SECONDS, rep_name="") -> bool:
    """
    Espera o relatório ficar 'estável' (sem indicadores de loading) por um período contínuo.
    Retorna True se conseguiu, senão dispara TimeoutError.
    """
    end = time.time() + timeout
    stable_start = None
    while time.time() < end:
        busy = _has_loading_indicators(driver)
        if not busy:
            if stable_start is None:
                stable_start = time.time()
                log.info(f"[{rep_name}] Sem loading — iniciando janela de estabilidade…")
            if time.time() - stable_start >= stable_required:
                log.info(f"[{rep_name}] Estável por {stable_required}s.")
                return True
        else:
            stable_start = None
        time.sleep(0.4)
    raise TimeoutError(f"[{rep_name}] Não ficou estável a tempo.")

# ======================= EXPORTAÇÃO (IMEDIATO / FALLBACK) ====================
def open_export_menu_unified(driver, wait, rep_name) -> str:
    """
    Abre o menu de exportação do Power BI por diferentes caminhos:
    1) Botão 'Exportar' direto;
    2) Menu 'Mais opções' (overflow) -> 'Exportar para PDF';
    3) Menu 'Arquivo' -> 'Exportar' -> 'PDF'.
    Retorna uma string indicando qual caminho funcionou.
    """
    pbix_ready(driver, wait)

    # Tenta o botão 'Exportar' direto
    export_btn = ("//button[@id='exportMenuBtn'] | "
                  "//button[contains(@title,'Export') or contains(@aria-label,'Export') or contains(@aria-label,'Exportar')]")
    try:
        log.info(f"[{rep_name}] Procurando botão 'Exportar' direto…")
        click_anywhere(driver, By.XPATH, export_btn, wait)
        log.info(f"[{rep_name}] Menu 'Exportar' (direto).")
        return "direct"
    except Exception as e:
        log.info(f"[{rep_name}] 'Exportar' direto não disponível: {e}")

    # Tenta via menu de 'Mais opções'
    try:
        overflow = ("//button[contains(@aria-label,'Mais opções') or contains(@aria-label,'More options') "
                    "or contains(@title,'Mais opções') or contains(@title,'More options')]")
        log.info(f"[{rep_name}] Tentando overflow…")
        click_anywhere(driver, By.XPATH, overflow, wait)
        time.sleep(0.4)
        pdf_item = (
            "//button[.//span[contains(text(),'Exportar para PDF') or contains(text(),'Export to PDF')]]"
            " | //div[.//span[contains(text(),'Exportar para PDF') or contains(text(),'Export to PDF')]]//button"
            " | //li[.//span[contains(text(),'Exportar para PDF') or contains(text(),'Export to PDF')]]"
        )
        click_anywhere(driver, By.XPATH, pdf_item, wait, must_be_clickable=False)
        log.info(f"[{rep_name}] 'Exportar para PDF' via overflow.")
        return "overflow"
    except Exception as e:
        log.info(f"[{rep_name}] Overflow indisponível: {e}")

    # Tenta via menu 'Arquivo'
    try:
        file_menu = ("//button[contains(@aria-label,'Arquivo') or contains(@aria-label,'File') "
                     "or contains(@title,'Arquivo') or contains(@title,'File')]")
        log.info(f"[{rep_name}] Tentando Arquivo → Exportar → PDF…")
        click_anywhere(driver, By.XPATH, file_menu, wait)
        time.sleep(0.4)
        export_sub = ("//button[.//span[contains(text(),'Exportar') or contains(text(),'Export')]] | "
                      "//div[.//span[contains(text(),'Exportar') or contains(text(),'Export')]]//button")
        click_anywhere(driver, By.XPATH, export_sub, wait, must_be_clickable=False)
        time.sleep(0.4)
        pdf_item2 = "//button[.//span[contains(text(),'PDF')]] | //div[.//span[contains(text(),'PDF')]]//button"
        click_anywhere(driver, By.XPATH, pdf_item2, wait, must_be_clickable=False)
        log.info(f"[{rep_name}] 'PDF' via Arquivo.")
        return "filemenu"
    except Exception as e:
        log.info(f"[{rep_name}] Arquivo→Exportar indisponível: {e}")

    # Se nenhum caminho funcionou, dispare erro
    raise RuntimeError(f"[{rep_name}] Não consegui abrir o menu de exportação.")

def confirm_pdf_export(driver, wait, rep_name) -> bool:
    """
    Confirma o diálogo de exportação (OK/Exportar), independente da variação de idioma/layout.
    """
    driver.switch_to.default_content()
    selectors = [
        "//*[@id='okButton']",
        "//button[@data-testid='dialog-ok-btn']",
        "//button[normalize-space(.)='OK']",
        "//button[normalize-space(.)='Exportar']",
        "//button[normalize-space(.)='Export']",
    ]
    for sel in selectors:
        try:
            click_anywhere(driver, By.XPATH, sel, wait, must_be_clickable=True, timeout=15)
            log.info(f"[{rep_name}] Confirmação do diálogo acionada.")
            return True
        except Exception:
            continue
    # Fallback: tenta clicar por texto direto
    try:
        driver.switch_to.default_content()
        for b in driver.find_elements(By.TAG_NAME, "button"):
            t = (b.text or "").strip().lower()
            if t in ("ok", "exportar", "export"):
                driver.execute_script("arguments[0].click();", b)
                log.info(f"[{rep_name}] Confirmação via fallback de texto.")
                return True
    except Exception as e:
        log.error(f"[{rep_name}] Falha na confirmação: {e}")
    raise RuntimeError(f"[{rep_name}] Não consegui confirmar exportação.")

def export_to_pdf_immediate(driver, wait, rep_name) -> None:
    """
    Executa a sequência para exportação em PDF no modo 'imediato' (sem aguardar idle).
    """
    if PRE_EXPORT_COOLDOWN_SEC > 0:
        time.sleep(PRE_EXPORT_COOLDOWN_SEC)
    path_used = open_export_menu_unified(driver, wait, rep_name)
    if path_used == "direct":
        # Em alguns layouts, há um botão adicional "Export to PDF"
        try:
            click_anywhere(driver, By.XPATH, "//button[@data-testid='export-to-pdf-btn']", wait)
        except Exception:
            click_anywhere(driver, By.XPATH, "//button[.//span[contains(text(),'PDF')]]", wait)
    confirm_pdf_export(driver, wait, rep_name)
    log.info(f"[{rep_name}] Exportação acionada (modo imediato).")

def export_to_pdf_with_idle(driver, wait, rep_name) -> None:
    """
    Fallback: aguarda o relatório ficar estável (sem loading) e então exporta.
    """
    switch_to_frame_with(driver, By.TAG_NAME, "button", timeout=10)
    wait_report_idle(driver, timeout=IDLE_TIMEOUT_SEC, stable_required=IDLE_STABLE_SECONDS, rep_name=rep_name)
    export_to_pdf_immediate(driver, wait, rep_name)

# =============================== PDF HELPERS =================================
def merge_pdfs(paths: List[str], output: str):
    """
    Mescla múltiplos PDFs em um único arquivo, na ordem recebida.
    """
    log.info(f"Mesclando {len(paths)} PDFs em {output}…")
    m = PdfMerger()
    for p in paths:
        m.append(p)
    m.write(output)
    m.close()
    log.info("Mesclagem concluída.")

def extract_single_page(src: str, dst: str, page_no: int):
    """
    Cria um novo PDF contendo apenas a página 'page_no' (1-based) do arquivo 'src'.
    """
    with open(src, "rb") as fsrc:
        reader = PdfReader(fsrc, strict=False)
        writer = PdfWriter()
        writer.add_page(reader.pages[page_no - 1])
        with open(dst, "wb") as fdst:
            writer.write(fdst)

def strip_last_pages(src: str, dst: str, n_last: int = 1):
    """
    Remove as 'n_last' últimas páginas do PDF 'src' e salva em 'dst'.
    """
    with open(src, "rb") as fsrc:
        reader = PdfReader(fsrc, strict=False)
        total = len(reader.pages)
        writer = PdfWriter()
        for i in range(max(0, total - n_last)):
            writer.add_page(reader.pages[i])
        with open(dst, "wb") as fdst:
            writer.write(fdst)

def strip_specific_pages(src: str, dst: str, pages_to_remove: List[int]):
    """
    Remove páginas específicas (1-based) de 'src' e salva em 'dst'.
    """
    with open(src, "rb") as fsrc:
        reader = PdfReader(fsrc, strict=False)
        writer = PdfWriter()
        total = len(reader.pages)
        to_remove = set(pages_to_remove)
        for i in range(total):
            if (i + 1) not in to_remove:
                writer.add_page(reader.pages[i])
        with open(dst, "wb") as fdst:
            writer.write(fdst)

# ================================ UTIL GERAIS ================================
def _is_temp_file(fn: str) -> bool:
    """
    Heurística para identificar arquivos temporários de download (Edge/Chromium).
    """
    f = fn.lower()
    return f.endswith((".crdownload", ".tmp", ".part", ".partial"))

def wait_for_pdf_download(download_dir: str, existing_files: set, rep_name: str,
                          timeout: int = DOWNLOAD_TIMEOUT_SEC) -> str:
    """
    Observa a pasta 'download_dir' até detectar um novo PDF completo (sem extensão temporária).
    Retorna o caminho do PDF finalizado.
    """
    deadline = time.time() + timeout
    started = False
    last_base = None
    log.info(f"[{rep_name}] Aguardando download…")
    while time.time() < deadline:
        new = set(os.listdir(download_dir)) - existing_files
        # detecta início do download (qualquer arquivo .pdf/.crdownload/.tmp etc.)
        if not started and any(n.lower().endswith((".pdf", ".crdownload", ".tmp", ".part", ".partial")) for n in new):
            started = True
            log.info(f"[{rep_name}] Download detectado: {list(new)}")
        if started:
            pdfs = [n for n in new if n.lower().endswith(".pdf")]
            temps = [n for n in new if _is_temp_file(n)]
            # só termina quando há PDF e nenhum temporário (indicando finalização)
            if pdfs and not temps:
                paths = [os.path.join(download_dir, f) for f in pdfs]
                latest = max(paths, key=os.path.getmtime)
                # dupla checagem de tamanho para garantir que o arquivo estável
                s1 = os.path.getsize(latest)
                time.sleep(1.2)
                s2 = os.path.getsize(latest)
                if s1 == s2 and s2 > 0:
                    log.info(f"[{rep_name}] Download concluído: {os.path.basename(latest)} ({s2} bytes)")
                    return latest
                else:
                    base = os.path.basename(latest)
                    if base != last_base:
                        log.info(f"[{rep_name}] Aguardando finalizar ({s1} → {s2})…")
                        last_base = base
        time.sleep(0.8)
    raise TimeoutError(f"[{rep_name}] Não baixou PDF a tempo.")

def read_recipients(folder: str) -> List[str]:
    """
    Lê o primeiro .txt ou .csv encontrado em 'folder' e retorna a lista de e-mails.
    Cada linha deve conter um endereço de e-mail.
    """
    for fn in os.listdir(folder):
        if fn.lower().endswith((".txt", ".csv")):
            with open(os.path.join(folder, fn), "r", encoding="utf-8") as f:
                return [l.strip() for l in f if l.strip()]
    return []

def send_email_outlook_html(to_list: List[str], subject: str, html_body: str, attach: str):
    """
    Envia e-mail via Microsoft Outlook (COM) com corpo HTML e anexo.
    - Abre o Outlook se não houver instância ativa.
    """
    try:
        ol = win32.DispatchEx("Outlook.Application")
    except Exception:
        # Tenta abrir o executável do Outlook se não estiver rodando
        candidates = [
            r"[seu caminho]\OUTLOOK.EXE",  # Ex.: C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE
            r"[seu caminho]\OUTLOOK.EXE",  # Ex.: C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE
        ]
        for path in candidates:
            if os.path.exists(path):
                subprocess.Popen([path])
                time.sleep(8)
                break
        ol = win32.DispatchEx("Outlook.Application")

    # Cria e envia o e-mail
    mail = ol.CreateItem(0)
    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.To = ";".join(to_list)
    mail.Attachments.Add(attach)
    try:
        mail.Send()
        log.info("E-mail enviado.")
    except Exception as e:
        log.error(f"Falha ao enviar e-mail: {e}\nDestinatários: {to_list}", exc_info=True)

# ============================== FLUXO PRINCIPAL ==============================
def process_single_report(driver: webdriver.Edge, rep: Report, wait: WebDriverWait) -> Optional[str]:
    """
    Fluxo para um único relatório:
    1) Abre URL;
    2) Tenta exportação imediata (retries);
    3) Se falhar, fallback aguardando estabilidade e exporta;
    4) Espera download finalizar;
    5) Aplica regra de páginas (se houver);
    6) Retorna caminho do PDF final do relatório.
    """
    rep.validate()
    rep_name = rep.name
    try:
        log.info(f"==== Iniciando: {rep_name} ====")
        # Mapeia arquivos existentes para detectar somente os novos
        existing = set(os.listdir(RUN_DIR))

        # Abre a URL do relatório publicado
        driver.get(rep.url)
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        pbix_ready(driver, wait, extra_sleep=0.5)

        # Tenta exportação imediata com N tentativas
        immediate_ok = False
        if FORCE_EXPORT_IMMEDIATE:
            for attempt in range(1, IMMEDIATE_TRIES + 1):
                try:
                    log.info(f"[{rep_name}] Export imediato (tentativa {attempt}/{IMMEDIATE_TRIES})…")
                    export_to_pdf_immediate(driver, wait, rep_name)
                    immediate_ok = True
                    break
                except Exception as e:
                    log.warning(f"[{rep_name}] Imediato falhou: {e}")
                    driver.refresh()
                    time.sleep(1.0)

        # Se não conseguiu no modo imediato, faz fallback aguardando estabilidade
        if not immediate_ok:
            log.info(f"[{rep_name}] Fallback com idle…")
            export_to_pdf_with_idle(driver, wait, rep_name)

        # Aguarda o download do PDF novo finalizar
        pdf_tmp = wait_for_pdf_download(RUN_DIR, existing, rep_name, timeout=DOWNLOAD_TIMEOUT_SEC)

        # Move/renomeia para um arquivo padrão do relatório no RUN_DIR
        destino = os.path.join(RUN_DIR, f"{rep_name}_{DATA_HOJE}.pdf")
        if os.path.exists(destino):
            os.remove(destino)
        shutil.move(pdf_tmp, destino)
        log.info(f"[{rep_name}] Salvo: {destino}")

        # Aplica regra de páginas (se configurada)
        if rep.extract_page:
            only = os.path.join(RUN_DIR, f"{rep_name}_{DATA_HOJE}_pg{rep.extract_page}.pdf")
            log.info(f"[{rep_name}] Extraindo página {rep.extract_page}…")
            extract_single_page(destino, only, rep.extract_page)
            final_path = only
        elif rep.drop_last_pages:
            trimmed = os.path.join(RUN_DIR, f"{rep_name}_{DATA_HOJE}_sem_ult.pdf")
            log.info(f"[{rep_name}] Removendo últimas {rep.drop_last_pages} pág.…")
            strip_last_pages(destino, trimmed, rep.drop_last_pages)
            final_path = trimmed
        elif rep.drop_pages:
            trimmed = os.path.join(RUN_DIR, f"{rep_name}_{DATA_HOJE}_sem_paginas.pdf")
            log.info(f"[{rep_name}] Removendo páginas: {rep.drop_pages}…")
            strip_specific_pages(destino, trimmed, rep.drop_pages)
            final_path = trimmed
        else:
            final_path = destino

        log.info(f"==== Finalizado: {rep_name} ====")
        return final_path

    except Exception:
        log.error(f"❌ Erro em {rep_name}", exc_info=True)
        return None

def run_all_reports(driver: webdriver.Edge, reports: List[Report]) -> Tuple[List[str], List[str]]:
    """
    Percorre todos os relatórios definidos em REPORTS, processando cada um.
    Retorna (lista_de_pdfs_ok, lista_de_nomes_falhados).
    """
    wait = WebDriverWait(driver, 60)  # Timeout padrão para operações do Selenium
    ok_paths: List[str] = []
    falhas: List[str] = []
    for rep in reports:
        out = process_single_report(driver, rep, wait)
        if out:
            ok_paths.append(out)
        else:
            falhas.append(rep.name)
        time.sleep(0.5)  # Pequena pausa entre relatórios
    return ok_paths, falhas

def build_email_html(ok_reports: List[Report], failed_names: List[str]) -> str:
    """
    Monta o corpo do e-mail em HTML:
    - Lista de relatórios processados com links (substituir [seu link] pelos reais).
    - Se houver falhas, lista nominal.
    """
    lis_ok = [f'<li>{r.name} — <a href="{r.url}">abrir</a></li>' for r in ok_reports]

    failed_html = ""
    if failed_names:
        failed_html = "<p><b>Falhas:</b></p><ul>" + "".join(f"<li>{n}</li>" for n in failed_names) + "</ul>"

    html = f"""
    <p>Bom dia,</p>
    <p>Segue o PDF diário consolidado (data {DATA_HOJE}) em anexo.</p>
    <p>Relatórios processados:</p>
    <ul>{''.join(lis_ok)}</ul>
    {failed_html}
    <p>Atenciosamente,<br>Sua Equipe</p>
    """
    return html

def main():
    """
    Função principal:
    - Abre o Edge,
    - Processa todos os relatórios,
    - Mescla PDFs,
    - Lê destinatários,
    - Envia e-mail com HTML e anexo.
    """
    log.info("Iniciando execução principal…")
    driver = setup_edge_driver(RUN_DIR)
    try:
        # Processa todos os relatórios
        ok_paths, falhas = run_all_reports(driver, REPORTS)

        if not ok_paths:
            log.error("Nenhum PDF baixado com sucesso. Abortando mescla/envio.")
            return

        # Mescla PDFs gerados em um único consolidado do dia
        merge_pdfs(ok_paths, PDF_FINAL_PATH)

        # Lê destinatários a partir da pasta configurada
        recipients = read_recipients(EMAIL_LIST_DIR)
        if not recipients:
            log.warning("Nenhum destinatário encontrado. Não haverá envio.")
            return

        # Prepara HTML do e-mail
        ok_reports = [r for r in REPORTS if any(os.path.basename(p).startswith(r.name + "_") for p in ok_paths)]
        html = build_email_html(ok_reports, falhas)

        # Assunto padrão
        subject = f"Indicadores diários ({DATA_HOJE})"

        # Envia e-mail via Outlook
        send_email_outlook_html(recipients, subject, html, PDF_FINAL_PATH)
    finally:
        # Fecha o navegador mesmo em caso de erro
        try:
            driver.quit()
        except Exception:
            pass
        log.info("Execução finalizada.")

# ============================== AGENDAMENTO ==================================
if __name__ == "__main__":
    # Agenda a execução diária no horário configurado (RUN_HOUR_MINUTE).
    schedule.every().day.at(RUN_HOUR_MINUTE).do(main)
    log.info(f"⏱️ Aguardando {RUN_HOUR_MINUTE} para enviar indicadores…")

    # Loop simples que mantém o agendamento ativo.
    while True:
        try:
            schedule.run_pending()
        except Exception:
            log.error("Erro no loop de agendamento.", exc_info=True)
        time.sleep(30)
