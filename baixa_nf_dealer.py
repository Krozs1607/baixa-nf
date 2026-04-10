"""
Automação de Baixa de NF no Dealer.net (Playwright)
Lê o Excel 'Mandarim Iguatemi.xlsx' e faz a baixa de cada NF no site.

Estrutura de iframes do site:
  - Página principal (default.html)
    - iframe[1] (menucontarecebertitulo.aspx) => formulário principal
      - iframe#gxp0_ifrm (wp_titulomov.aspx) => popup Movimento do Título
"""

import time
import logging
import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# ============================================================
# CONFIGURAÇÕES
# ============================================================
EXCEL_FILE = "Mandarim Iguatemi.xlsx"
URL_LOGIN = "https://workflow.grupoindiana.com.br/Portal/default.html"
USUARIO = "IGORFERREIRA"
SENHA = "123@Mudar"

# Campos do Movimento do Título
TIPO_CREDITO_VALUE = "64"       # RECEBIMENTO DE TÍTULO
AGENTE_COBRADOR_VALUE = "55"    # CONTA MOVIMENTO FABRICA (3.06.60)
TIPO_DOCUMENTO_VALUE = "2"      # AVISO DE LANCAMENTO
HISTORICO_TEXTO = "Baixa Garantia"

# Começa a partir de qual NF (1 = primeira, útil para retomar)
INICIAR_A_PARTIR_DE = 1

# ============================================================
# LOGGING
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("baixa_nf_log.txt", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger(__name__)


def ler_notas_do_excel(caminho: str) -> list[dict]:
    """Lê as NFs do Excel e retorna lista de dicts."""
    wb = openpyxl.load_workbook(caminho, read_only=True)
    ws = wb.active
    notas = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        nf = row[1].value   # Coluna B = Referência
        valor = row[3].value  # Coluna D = Total Geral
        if nf is not None:
            notas.append({"nf": str(int(nf)), "valor": valor})
    wb.close()
    return notas


def get_main_frame(page):
    """Retorna o frame principal (menucontarecebertitulo.aspx)."""
    for frame in page.frames:
        if "menucontarecebertitulo" in frame.url:
            return frame
    raise Exception("Frame menucontarecebertitulo não encontrado!")


def get_popup_frame(main_frame):
    """Retorna o frame do popup (wp_titulomov.aspx) dentro do frame principal."""
    for frame in main_frame.child_frames:
        if "wp_titulomov" in frame.url:
            return frame
    return None


def login(page):
    """Faz login no Dealer.net."""
    log.info("Navegando para o site...")
    page.goto(URL_LOGIN, wait_until="networkidle", timeout=60000)
    page.wait_for_timeout(3000)

    try:
        # Tenta preencher campos de login
        page.locator("input[type='text']").first.fill(USUARIO)
        page.wait_for_timeout(300)
        page.locator("input[type='password']").first.fill(SENHA)
        page.wait_for_timeout(300)
        page.locator("input[type='submit'], button[type='submit']").first.click()
        page.wait_for_timeout(5000)
        log.info("Login realizado!")
    except Exception as e:
        log.warning(f"Login automático falhou: {e}")
        log.info(">>> FAÇA O LOGIN MANUALMENTE NO NAVEGADOR <<<")
        input(">>> Pressione ENTER após fazer login...")


def navegar_titulo_a_receber(page):
    """Navega até a tela de Título a Receber."""
    log.info("Navegando para Título a Receber...")
    try:
        page.locator("text=Financeiro").first.click()
        page.wait_for_timeout(1500)
        page.locator("text=Contas a Receber").first.click()
        page.wait_for_timeout(1500)
        page.locator("text=Título a Receber").last.click()
        page.wait_for_timeout(3000)
        log.info("Tela de Título a Receber aberta!")
    except Exception as e:
        log.warning(f"Navegação automática falhou: {e}")
        log.info(">>> NAVEGUE MANUALMENTE ATÉ 'Título a Receber' <<<")
        input(">>> Pressione ENTER quando estiver na tela...")


def expandir_filtro_avancado(main_frame):
    """Expande o Filtro Avançado se estiver fechado."""
    try:
        fieldset = main_frame.locator("#GROUPFILTROAVANCADO")
        if not fieldset.is_visible(timeout=2000):
            # Clica no botão de expandir (o "+" ao lado de Filtro Avançado)
            main_frame.evaluate("""
                const fs = document.getElementById('GROUPFILTROAVANCADO');
                if (fs) fs.style.display = 'block';
            """)
            main_frame.wait_for_timeout(500)
            # Se não funcionou, tenta clicar na legenda
            if not fieldset.is_visible(timeout=1000):
                log.info(">>> EXPANDA O 'Filtro Avançado' MANUALMENTE <<<")
                input(">>> Pressione ENTER após expandir...")
        log.info("Filtro Avançado visível")
    except:
        log.info(">>> EXPANDA O 'Filtro Avançado' MANUALMENTE <<<")
        input(">>> Pressione ENTER após expandir...")


def processar_nf(page, main_frame, nf_numero: str, indice: int, total: int) -> str:
    """
    Processa uma NF. Retorna:
    - 'sucesso' se deu certo
    - 'pago' se já estava paga
    - 'nao_encontrada' se não encontrou
    - 'erro' se deu erro
    """
    log.info(f"[{indice}/{total}] Processando NF: {nf_numero}")

    # === PASSO 1: Limpar e preencher NFS-e ===
    try:
        nfse_field = main_frame.locator("#vTITULO_NOTAFISCALNRONFSE")
        nfse_field.click()
        nfse_field.fill("")
        main_frame.wait_for_timeout(200)
        nfse_field.fill(nf_numero)
        main_frame.wait_for_timeout(300)
    except Exception as e:
        log.error(f"  ERRO ao preencher NFS-e: {e}")
        return "erro"

    # === PASSO 2: Clicar Consultar ===
    try:
        main_frame.locator("#BTNCONSULTAR").click()
        main_frame.wait_for_timeout(4000)
    except Exception as e:
        log.error(f"  ERRO ao clicar Consultar: {e}")
        return "erro"

    # === PASSO 3: Verificar se encontrou resultado ===
    try:
        row = main_frame.locator("#GridContainerRow_0001")
        if not row.is_visible(timeout=3000):
            log.warning(f"  NF {nf_numero} NÃO ENCONTRADA")
            return "nao_encontrada"
    except:
        log.warning(f"  NF {nf_numero} NÃO ENCONTRADA (timeout)")
        return "nao_encontrada"

    # === PASSO 4: Verificar Status (pular se Pago) ===
    try:
        status_text = main_frame.locator("#span_vGRID_TITULO_STATUS_0001").text_content(timeout=2000)
        status_text = status_text.strip()
        log.info(f"  Status: {status_text}")
        if status_text.lower() == "pago":
            log.info(f"  NF {nf_numero} já está PAGA - pulando")
            return "pago"
    except Exception as e:
        log.warning(f"  Não conseguiu ler status: {e}")

    # === PASSO 5: Clicar Movimento ===
    try:
        main_frame.locator("#vBMPMOVIMENTO_0001").click()
        main_frame.wait_for_timeout(3000)
    except Exception as e:
        log.error(f"  ERRO ao clicar Movimento: {e}")
        return "erro"

    # === PASSO 6: No popup, verificar se o botão INSERT existe e clicar ===
    try:
        popup_frame = get_popup_frame(main_frame)
        if popup_frame is None:
            log.error("  Popup de movimento não abriu")
            return "erro"

        insert_btn = popup_frame.locator("#INSERT")
        if not insert_btn.is_visible(timeout=3000):
            log.warning(f"  NF {nf_numero} - botão INSERT não visível (pode já estar baixada)")
            # Fecha o popup
            try:
                popup_frame.locator("#FECHAR").click()
                main_frame.wait_for_timeout(2000)
            except:
                main_frame.evaluate("document.getElementById('gxp0_cls')?.click()")
                main_frame.wait_for_timeout(2000)
            return "pago"

        insert_btn.click()
        main_frame.wait_for_timeout(3000)
        log.info("  Formulário de novo movimento aberto")
    except Exception as e:
        log.error(f"  ERRO ao clicar INSERT: {e}")
        return "erro"

    # === PASSO 7: Preencher Tipo Crédito/Débito ===
    try:
        # O popup pode ter recarregado, pegar o frame novamente
        popup_frame = get_popup_frame(main_frame)
        if popup_frame is None:
            log.error("  Popup de formulário não encontrado")
            return "erro"

        popup_frame.locator("#TITULOMOV_TIPOCDCOD").select_option(value=TIPO_CREDITO_VALUE)
        popup_frame.wait_for_timeout(1000)
        log.info(f"  Tipo Crédito/Débito: RECEBIMENTO DE TÍTULO")
    except Exception as e:
        log.error(f"  ERRO ao selecionar Tipo Crédito/Débito: {e}")
        return "erro"

    # === PASSO 7b: Preencher Agente Cobrador ===
    try:
        popup_frame.locator("#TITULOMOV_AGENTECOBRADORCOD").select_option(value=AGENTE_COBRADOR_VALUE)
        popup_frame.wait_for_timeout(1000)
        log.info(f"  Agente Cobrador: CONTA MOVIMENTO FABRICA (3.06.60)")
    except Exception as e:
        log.error(f"  ERRO ao selecionar Agente Cobrador: {e}")
        return "erro"

    # === PASSO 8: Preencher Tipo de Documento ===
    try:
        popup_frame.locator("#TITULOMOV_TIPODOCUMENTOCOD").select_option(value=TIPO_DOCUMENTO_VALUE)
        popup_frame.wait_for_timeout(500)
        log.info(f"  Tipo Documento: AVISO DE LANCAMENTO")
    except Exception as e:
        log.error(f"  ERRO ao selecionar Tipo Documento: {e}")
        return "erro"

    # === PASSO 9: Preencher Histórico ===
    try:
        popup_frame.locator("#TITULOMOV_HISTORICO").fill(HISTORICO_TEXTO)
        popup_frame.wait_for_timeout(500)
        log.info(f"  Histórico: {HISTORICO_TEXTO}")
    except Exception as e:
        log.error(f"  ERRO ao preencher Histórico: {e}")
        return "erro"

    # === PASSO 10: Clicar Confirmar ===
    try:
        popup_frame.locator("#TRN_ENTER").click()
        main_frame.wait_for_timeout(4000)
        log.info(f"  ✓ NF {nf_numero} CONFIRMADA!")
    except Exception as e:
        log.error(f"  ERRO ao confirmar: {e}")
        return "erro"

    # === PASSO 11: Fechar popup de movimento ===
    try:
        popup_frame = get_popup_frame(main_frame)
        if popup_frame:
            fechar = popup_frame.locator("#FECHAR")
            if fechar.is_visible(timeout=3000):
                fechar.click()
                main_frame.wait_for_timeout(2000)
    except:
        try:
            main_frame.evaluate("document.getElementById('gxp0_cls')?.click()")
            main_frame.wait_for_timeout(2000)
        except:
            pass

    return "sucesso"


def main():
    log.info("=" * 60)
    log.info("AUTOMAÇÃO DE BAIXA DE NF - DEALER.NET")
    log.info("=" * 60)

    # Lê as NFs do Excel
    notas = ler_notas_do_excel(EXCEL_FILE)
    total = len(notas)
    log.info(f"Total de NFs no Excel: {total}")
    log.info(f"Iniciando a partir da NF #{INICIAR_A_PARTIR_DE}")

    contadores = {"sucesso": 0, "pago": 0, "nao_encontrada": 0, "erro": 0}
    nfs_com_erro = []

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            slow_mo=300,
        )
        context = browser.new_context(
            viewport={"width": 1366, "height": 768},
        )
        page = context.new_page()
        page.set_default_timeout(15000)

        # Login
        login(page)

        # Navega até Título a Receber
        navegar_titulo_a_receber(page)

        # Pega o frame principal
        page.wait_for_timeout(3000)
        try:
            main_frame = get_main_frame(page)
        except:
            log.info(">>> Aguardando frame principal carregar... <<<")
            page.wait_for_timeout(5000)
            main_frame = get_main_frame(page)

        # Expande filtro avançado
        expandir_filtro_avancado(main_frame)

        # Processa cada NF
        erros_seguidos = 0
        for i, nota in enumerate(notas, start=1):
            if i < INICIAR_A_PARTIR_DE:
                continue

            resultado = processar_nf(page, main_frame, nota["nf"], i, total)
            contadores[resultado] += 1

            if resultado == "erro":
                nfs_com_erro.append(nota["nf"])
                erros_seguidos += 1

                # Tenta recuperar: fecha popups e volta ao estado inicial
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(1000)
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(1000)
                except:
                    pass

                if erros_seguidos >= 3:
                    log.error("3 erros seguidos! Verifique o navegador.")
                    resp = input(">>> Digite 'c' para continuar ou 'q' para parar: ")
                    if resp.lower() == 'q':
                        break
                    erros_seguidos = 0
            else:
                erros_seguidos = 0

            # Log de progresso a cada 10 NFs
            if i % 10 == 0:
                processadas = sum(contadores.values())
                log.info(f"--- PROGRESSO: {processadas}/{total} | "
                         f"OK: {contadores['sucesso']} | "
                         f"Pago: {contadores['pago']} | "
                         f"Não encontrada: {contadores['nao_encontrada']} | "
                         f"Erro: {contadores['erro']} ---")

        # Relatório final
        log.info("=" * 60)
        log.info("RELATÓRIO FINAL")
        log.info(f"  Sucesso:         {contadores['sucesso']}")
        log.info(f"  Já pagas:        {contadores['pago']}")
        log.info(f"  Não encontradas: {contadores['nao_encontrada']}")
        log.info(f"  Erros:           {contadores['erro']}")
        if nfs_com_erro:
            log.info(f"  NFs com erro: {', '.join(nfs_com_erro)}")
        log.info("=" * 60)

        input(">>> Pressione ENTER para fechar o navegador...")
        browser.close()


if __name__ == "__main__":
    main()
