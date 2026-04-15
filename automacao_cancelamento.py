"""
Módulo de CANCELAMENTO de baixas de NF para GAULESA no Dealer.net
Busca por CHASSI, encontra o movimento pelo VALOR e cancela com motivo "Erro".
"""

import time
import logging
import openpyxl
from playwright.sync_api import sync_playwright

URL_DEALER = "https://workflow.grupoindiana.com.br/Portal/default.html"

log = logging.getLogger("cancelamento")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


class AutomacaoCancelamento:
    def __init__(self, caminho_excel: str, estado: dict):
        self.caminho_excel = caminho_excel
        self.estado = estado
        self.notas = []
        self.parar = False
        self.pausado = False

    def _log(self, msg: str):
        log.info(msg)
        self.estado["log_mensagens"].append(msg)

    def carregar_notas(self) -> int:
        wb = openpyxl.load_workbook(self.caminho_excel, read_only=True)
        ws = wb.active
        self.notas = []
        for row in ws.iter_rows(min_row=2, values_only=False):
            chassi = row[0].value
            valor = row[1].value
            if chassi and valor:
                self.notas.append({"chassi": str(chassi).strip(), "valor": float(valor)})
        wb.close()
        self._log(f"Excel carregado: {len(self.notas)} chassis para cancelar")
        return len(self.notas)

    def _get_main_frame(self, page):
        for tentativa in range(20):
            for frame in page.frames:
                try:
                    if frame.query_selector("#BTNCONSULTAR"):
                        return frame
                except:
                    pass
            time.sleep(3)
        raise Exception("Frame nao encontrado!")

    def _get_popup_frame(self, main_frame, tentativas=8):
        """Encontra gxp0_ifrm (popup de Movimento)."""
        for t in range(tentativas):
            for frame in main_frame.page.frames:
                try:
                    if "titulomov" in frame.url.lower():
                        return frame
                except:
                    pass
            time.sleep(2)
        return None

    def _get_cancel_frame(self, main_frame, tentativas=8):
        """Encontra gxp1_ifrm (popup de Cancelamento)."""
        for t in range(tentativas):
            for frame in main_frame.page.frames:
                try:
                    if frame.query_selector("#vHISTORICO_OBSERVACAO"):
                        return frame
                    if frame.query_selector("#IMGCONFIRMAR"):
                        return frame
                except:
                    pass
            time.sleep(2)
        return None

    def _expandir_filtro_avancado(self, main_frame):
        try:
            chassi = main_frame.query_selector("#vTITULO_VEICULOCHASSI")
            if chassi and chassi.is_visible():
                return
            main_frame.evaluate("""
                const fs = document.getElementById('GROUPFILTROAVANCADO');
                if (fs) {
                    fs.style.display = 'block';
                    fs.style.visibility = 'visible';
                    fs.style.height = 'auto';
                    fs.style.overflow = 'visible';
                    let parent = fs.parentElement;
                    while (parent) {
                        parent.style.display = parent.style.display === 'none' ? 'block' : parent.style.display;
                        parent.style.overflow = 'visible';
                        parent = parent.parentElement;
                    }
                }
            """)
            main_frame.wait_for_timeout(1000)
        except:
            pass

    def _parse_valor_br(self, texto: str) -> float:
        try:
            return float(texto.replace(".", "").replace(",", "."))
        except:
            return 0.0

    def _formatar_valor_br(self, valor: float) -> str:
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def _encontrar_linha_grid_por_valor(self, main_frame, valor_excel: float):
        """Encontra a linha no grid principal que tem o valor igual ao Excel."""
        for r in range(1, 20):
            idx = str(r).zfill(4)
            try:
                row = main_frame.locator(f"#GridContainerRow_{idx}")
                if not row.is_visible(timeout=500):
                    break
                valor_span = main_frame.locator(f"#span_vGRID_TITULO_VALOR_{idx}")
                valor_texto = valor_span.text_content(timeout=1000).strip()
                valor_dealer = self._parse_valor_br(valor_texto)
                if abs(valor_dealer - valor_excel) < 0.02:
                    self._log(f"    Linha grid {r}: Valor {valor_texto} MATCH!")
                    return r
            except:
                break
        return 0

    def _encontrar_movimento_por_valor(self, popup_frame, valor_excel: float):
        """Dentro do popup Movimento, encontra a linha com valor igual e retorna o índice."""
        for r in range(1, 20):
            idx = str(r).zfill(4)
            try:
                valor_span = popup_frame.locator(f"#span_TITULOMOV_VALOR_{idx}")
                if not valor_span.is_visible(timeout=500):
                    break
                valor_texto = valor_span.text_content(timeout=1000).strip()
                valor_mov = self._parse_valor_br(valor_texto)
                if abs(valor_mov - valor_excel) < 0.02:
                    self._log(f"    Movimento {r}: Valor {valor_texto} MATCH!")
                    return r
            except:
                break
        return 0

    def _processar_cancelamento(self, main_frame, nota, indice, total):
        chassi = nota["chassi"]
        valor_excel = nota["valor"]
        self.estado["nf_atual"] = chassi
        self._log(f"[{indice}/{total}] CANCELAR Chassi: {chassi} | Valor: {self._formatar_valor_br(valor_excel)}")

        self._expandir_filtro_avancado(main_frame)

        # 1. Buscar por Chassi
        try:
            campo = main_frame.locator("#vTITULO_VEICULOCHASSI")
            campo.click()
            campo.fill("")
            main_frame.wait_for_timeout(200)
            campo.fill(chassi)
            main_frame.wait_for_timeout(300)
            main_frame.locator("#BTNCONSULTAR").click()
            main_frame.wait_for_timeout(4000)
        except Exception as e:
            self._log(f"  ERRO consulta: {e}")
            return "erro"

        # 2. Verificar resultado
        try:
            row = main_frame.locator("#GridContainerRow_0001")
            if not row.is_visible(timeout=3000):
                self._log(f"  Chassi {chassi} NAO ENCONTRADO")
                return "nao_encontrada"
        except:
            self._log(f"  Chassi {chassi} NAO ENCONTRADO")
            return "nao_encontrada"

        # 3. Encontrar linha com valor igual no grid
        self._log(f"  Buscando NF com valor {self._formatar_valor_br(valor_excel)}...")
        linha_grid = self._encontrar_linha_grid_por_valor(main_frame, valor_excel)
        if linha_grid == 0:
            self._log(f"  Nenhuma NF com valor {self._formatar_valor_br(valor_excel)} - pulando")
            return "nao_encontrada"

        idx_grid = str(linha_grid).zfill(4)

        # 4. Clicar Movimento na linha correta
        try:
            main_frame.locator(f"#vBMPMOVIMENTO_{idx_grid}").click()
            main_frame.wait_for_timeout(3000)
        except Exception as e:
            self._log(f"  ERRO Movimento: {e}")
            return "erro"

        # 5. No popup de Movimento, encontrar a linha com valor igual
        popup = self._get_popup_frame(main_frame)
        if not popup:
            self._log("  ERRO: popup Movimento nao abriu")
            return "erro"

        mov_linha = self._encontrar_movimento_por_valor(popup, valor_excel)
        if mov_linha == 0:
            self._log(f"  Movimento com valor {self._formatar_valor_br(valor_excel)} nao encontrado")
            try:
                popup.locator("#FECHAR").click()
                main_frame.wait_for_timeout(2000)
            except:
                pass
            return "nao_encontrada"

        idx_mov = str(mov_linha).zfill(4)

        # 6. Clicar no ícone Cancelar (lixeira) dessa linha
        try:
            popup.locator(f"#vCANCELA_{idx_mov}").click()
            main_frame.wait_for_timeout(3000)
            self._log(f"  Cancelar clicado na linha {mov_linha}")
        except Exception as e:
            self._log(f"  ERRO ao clicar cancelar: {e}")
            return "erro"

        # 7. Preencher Motivo = "Erro" no popup de cancelamento (gxp1_ifrm)
        cancel_frame = self._get_cancel_frame(main_frame)
        if not cancel_frame:
            self._log("  ERRO: popup de cancelamento nao abriu")
            return "erro"

        try:
            cancel_frame.locator("#vHISTORICO_OBSERVACAO").fill("Erro")
            cancel_frame.wait_for_timeout(500)
            self._log(f"  Motivo: Erro")
        except Exception as e:
            self._log(f"  ERRO preencher motivo: {e}")
            return "erro"

        # 8. Confirmar cancelamento
        try:
            cancel_frame.locator("#IMGCONFIRMAR").click()
            main_frame.wait_for_timeout(4000)
            self._log(f"  >> Chassi {chassi} CANCELADO com sucesso!")
        except Exception as e:
            self._log(f"  ERRO confirmar cancelamento: {e}")
            return "erro"

        # 9. Fechar popup de Movimento
        try:
            popup = self._get_popup_frame(main_frame, tentativas=5)
            if popup:
                fechar = popup.locator("#FECHAR")
                if fechar.is_visible(timeout=5000):
                    fechar.click()
                    main_frame.wait_for_timeout(2000)
        except:
            try:
                for frame in main_frame.page.frames:
                    try:
                        btn = frame.query_selector("#FECHAR")
                        if btn and btn.is_visible():
                            btn.click()
                            main_frame.wait_for_timeout(2000)
                            break
                    except:
                        pass
            except:
                pass

        return "sucesso"

    def executar_cancelamento(self):
        """Abre browser e cancela as notas da Gaulesa."""
        MENSAGENS = {
            "sucesso": "Cancelada com sucesso",
            "nao_encontrada": "Chassi/valor nao encontrado",
            "erro": "Erro ao cancelar",
        }
        total = len(self.notas)
        erros_seguidos = 0

        with sync_playwright() as pw:
            self._log("Abrindo navegador para CANCELAMENTO Gaulesa...")
            browser = pw.chromium.launch(headless=False, slow_mo=200, args=["--start-maximized"])
            context = browser.new_context(no_viewport=True)
            page = context.new_page()
            page.set_default_timeout(15000)

            page.goto(URL_DEALER, wait_until="networkidle", timeout=60000)
            self._log("FACA LOGIN, selecione GAULESA e va em Titulo a Receber.")

            main_frame = None
            for tentativa in range(120):
                if self.parar:
                    browser.close()
                    self.estado["rodando"] = False
                    return
                for frame in page.frames:
                    try:
                        if frame.query_selector("#BTNCONSULTAR"):
                            main_frame = frame
                            break
                    except:
                        pass
                if main_frame:
                    self._log("Dealer pronto!")
                    break
                if tentativa % 10 == 0 and tentativa > 0:
                    self._log(f"Aguardando... ({tentativa*3}s)")
                time.sleep(3)

            if not main_frame:
                self._log("ERRO: Timeout.")
                browser.close()
                self.estado["rodando"] = False
                return

            self.estado["dealer_pronto"] = True
            self._log("DEALER PRONTO! Configure e clique INICIAR CANCELAMENTO.")

            while not self.estado.get("inicio_confirmado", False):
                if self.parar:
                    browser.close()
                    self.estado["rodando"] = False
                    return
                time.sleep(1)

            self._log("Iniciando cancelamentos...")
            self._expandir_filtro_avancado(main_frame)

            for i, nota in enumerate(self.notas, start=1):
                if self.parar:
                    self._log("PARADO pelo usuario")
                    break
                while self.pausado and not self.parar:
                    time.sleep(0.5)
                if self.parar:
                    break

                entrada = {
                    "cnpj": "",
                    "nf": nota["chassi"],
                    "nf_original": nota["chassi"],
                    "valor": nota["valor"],
                    "status": "processando",
                    "mensagem": "Cancelando...",
                }
                self.estado["tabela_nfs"].append(entrada)

                resultado = self._processar_cancelamento(main_frame, nota, i, total)
                self.estado["progresso"][resultado] = self.estado["progresso"].get(resultado, 0) + 1
                entrada["status"] = resultado
                entrada["mensagem"] = MENSAGENS.get(resultado, resultado)

                if resultado == "erro":
                    erros_seguidos += 1
                    try:
                        page.keyboard.press("Escape")
                        page.wait_for_timeout(1000)
                        page.keyboard.press("Escape")
                        page.wait_for_timeout(1000)
                    except:
                        pass
                    if erros_seguidos >= 5:
                        self._log("MUITOS ERROS (5)! Parando.")
                        break
                else:
                    erros_seguidos = 0

                if i % 10 == 0:
                    p = self.estado["progresso"]
                    self._log(f"--- PROGRESSO {i}/{total} | OK:{p['sucesso']} Erro:{p['erro']} ---")

            p = self.estado["progresso"]
            self._log("=" * 50)
            self._log(f"CANCELAMENTO FINALIZADO! Canceladas:{p['sucesso']} | Nao encontradas:{p.get('nao_encontrada',0)} | Erros:{p['erro']}")
            self._log("=" * 50)

            try:
                while not self.parar:
                    time.sleep(1)
            except:
                pass
            browser.close()
            self.estado["rodando"] = False
