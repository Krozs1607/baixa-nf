"""
Módulo de automação de baixa de NF no Dealer.net
Usado pelo servidor Flask (servidor.py)
Tudo roda numa única thread para evitar problemas com Playwright.
"""

import time
import logging
import openpyxl
from playwright.sync_api import sync_playwright

# Configurações dos campos
URL_DEALER = "https://workflow.grupoindiana.com.br/Portal/default.html"
TIPO_CREDITO_VALUE = "64"       # RECEBIMENTO DE TÍTULO
AGENTE_COBRADOR_VALUE = "55"    # CONTA MOVIMENTO FABRICA (3.06.60)
TIPO_DOCUMENTO_VALUE = "2"      # AVISO DE LANCAMENTO
EMPRESA_IGUATEMI_VALUE = "32"   # MANDARIM IGUATEMI (sempre matriz)
HISTORICO_TEXTO = "Baixa Garantia"

log = logging.getLogger("automacao")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


class AutomacaoBaixa:
    def __init__(self, caminho_excel: str, estado: dict, loja_key: str = ""):
        self.caminho_excel = caminho_excel
        self.estado = estado
        self.loja_key = loja_key
        self.notas = []  # Lista de {"nf": str, "valor": float}
        self.parar = False
        self.pausado = False

    def _log(self, msg: str):
        log.info(msg)
        self.estado["log_mensagens"].append(msg)

    def _formatar_nf(self, nf_numero: str) -> str:
        """Formata o número da NF. Para Itabuna e Lauro: 15 dígitos com prefixo 2026."""
        if self.loja_key in ("mandarim_itabuna", "mandarim_lauro"):
            return "2026" + nf_numero.zfill(11)
        return nf_numero

    def _formatar_nf_fallback(self, nf_numero: str) -> str:
        """Formato alternativo com prefixo 2025 (notas do ano anterior)."""
        return "2025" + nf_numero.zfill(11)

    def _formatar_valor(self, valor: float) -> str:
        """Converte valor do Excel (negativo) para formato brasileiro positivo. Ex: -2167.20 → '2167,20'"""
        valor_abs = abs(valor)
        # Formata com 2 casas decimais e vírgula
        return f"{valor_abs:.2f}".replace(".", ",")

    def carregar_notas(self) -> int:
        wb = openpyxl.load_workbook(self.caminho_excel, read_only=True)
        ws = wb.active
        self.notas = []
        for row in ws.iter_rows(min_row=2, values_only=False):
            cnpj = row[0].value  # Coluna A = CNPJ
            nf = row[1].value    # Coluna B = Referência
            valor = row[3].value  # Coluna D = Total Geral
            if nf is not None:
                nf_str = str(int(nf))
                nf_formatada = self._formatar_nf(nf_str)
                self.notas.append({
                    "cnpj": str(cnpj) if cnpj else "",
                    "nf": nf_formatada,
                    "nf_original": nf_str,
                    "valor": valor if valor else 0
                })
        wb.close()
        self._log(f"=== LOJA SELECIONADA: {self.loja_key} ===")
        if self.loja_key in ("mandarim_itabuna", "mandarim_lauro"):
            self._log(f"Excel carregado: {len(self.notas)} NFs (formato 15 digitos com prefixo 2026)")
            if self.notas:
                self._log(f"  Exemplo: NF original={self.notas[0]['nf_original']} -> formatada={self.notas[0]['nf']}")
        else:
            self._log(f"Excel carregado: {len(self.notas)} NFs encontradas")
        return len(self.notas)

    def _get_main_frame(self, page):
        """Encontra o frame que contém o formulário de Título a Receber."""
        for tentativa in range(20):
            for frame in page.frames:
                url = frame.url.lower()
                if "menucontarecebertitulo" in url or "contarecebertitulo" in url:
                    self._log(f"Frame encontrado por URL: {frame.url[:80]}")
                    return frame
            for frame in page.frames:
                try:
                    if frame.query_selector("#BTNCONSULTAR"):
                        self._log(f"Frame encontrado pelo Consultar: {frame.url[:80]}")
                        return frame
                except:
                    pass
            self._log(f"Tentativa {tentativa+1}/20 - {len(page.frames)} frames. Aguardando...")
            frame_urls = [f.url[:80] for f in page.frames]
            self._log(f"  URLs: {frame_urls}")
            time.sleep(3)
        raise Exception("Frame nao encontrado!")

    def _get_popup_frame(self, main_frame, procurar_formulario=False, tentativas=5):
        """Encontra o popup frame. Se procurar_formulario=True, busca o frame do formulário de inserção."""
        for t in range(tentativas):
            # Procura em child frames do main_frame
            for frame in main_frame.child_frames:
                try:
                    if "wp_titulomov" in frame.url or "titulomov" in frame.url.lower():
                        if procurar_formulario:
                            # Verifica se o formulário já carregou
                            if frame.query_selector("#TITULOMOV_TIPOCDCOD"):
                                return frame
                        else:
                            return frame
                except:
                    pass

            # Procura em TODOS os frames da página (pode ser popup aninhado)
            for frame in main_frame.page.frames:
                try:
                    url = frame.url.lower()
                    if "titulomov" in url:
                        if procurar_formulario:
                            if frame.query_selector("#TITULOMOV_TIPOCDCOD"):
                                return frame
                        else:
                            return frame
                except:
                    pass

            # Se procurando formulário, tenta encontrar por conteúdo
            if procurar_formulario:
                for frame in main_frame.page.frames:
                    try:
                        if frame.query_selector("#TITULOMOV_TIPOCDCOD"):
                            return frame
                    except:
                        pass

            time.sleep(2)

        return None

    def _expandir_filtro_avancado(self, main_frame):
        try:
            nfse = main_frame.query_selector("#vTITULO_NOTAFISCALNRONFSE")
            if nfse and nfse.is_visible():
                self._log("Filtro Avancado ja expandido")
                return

            # Força exibição do fieldset e todos os filhos via CSS
            main_frame.evaluate("""
                const fs = document.getElementById('GROUPFILTROAVANCADO');
                if (fs) {
                    fs.style.display = 'block';
                    fs.style.visibility = 'visible';
                    fs.style.height = 'auto';
                    fs.style.overflow = 'visible';
                    // Também garante que o pai esteja visível
                    let parent = fs.parentElement;
                    while (parent) {
                        parent.style.display = parent.style.display === 'none' ? 'block' : parent.style.display;
                        parent.style.overflow = 'visible';
                        parent = parent.parentElement;
                    }
                }
            """)
            main_frame.wait_for_timeout(1000)

            nfse = main_frame.query_selector("#vTITULO_NOTAFISCALNRONFSE")
            if nfse and nfse.is_visible():
                self._log("Filtro Avancado expandido via CSS")
                return

            # Tenta clicar no ícone "+" ao lado do Filtro Avançado
            main_frame.evaluate("""
                // Procura imagens/links próximos ao fieldset
                const imgs = document.querySelectorAll('img');
                for (const img of imgs) {
                    const src = (img.src || '').toLowerCase();
                    if (src.includes('plus') || src.includes('expand') || src.includes('add')) {
                        const rect = img.getBoundingClientRect();
                        if (rect.y > 200) { img.click(); break; }
                    }
                }
            """)
            main_frame.wait_for_timeout(1000)

            nfse = main_frame.query_selector("#vTITULO_NOTAFISCALNRONFSE")
            if nfse and nfse.is_visible():
                self._log("Filtro Avancado expandido via clique")
                return

            self._log("AVISO: Nao conseguiu expandir Filtro Avancado. Expanda manualmente!")
            # Espera até 60s para o usuário expandir manualmente
            for i in range(20):
                nfse = main_frame.query_selector("#vTITULO_NOTAFISCALNRONFSE")
                if nfse and nfse.is_visible():
                    self._log("Filtro Avancado expandido pelo usuario")
                    return
                time.sleep(3)

            self._log("ERRO: Campo NFS-e nao ficou visivel. Expanda o Filtro Avancado!")
        except Exception as e:
            self._log(f"Aviso filtro: {e}")

    def _processar_nf(self, main_frame, nota, indice, total):
        nf = nota["nf"]
        valor = nota["valor"]
        valor_excel = abs(valor) if valor else 0
        valor_formatado = self._formatar_valor(valor)
        self.estado["nf_atual"] = nf
        self._log(f"[{indice}/{total}] Processando NF: {nf} | Valor: {valor_formatado}")

        # Garante que o Filtro Avançado esteja expandido antes de cada NF
        self._expandir_filtro_avancado(main_frame)

        # Tenta buscar a NF (para Itabuna/Lauro, tenta 2026 primeiro, depois 2025)
        nf_buscar = nf
        encontrou = False

        for tentativa_nf in range(2):
            try:
                nfse = main_frame.locator("#vTITULO_NOTAFISCALNRONFSE")
                nfse.click()
                nfse.fill("")
                main_frame.wait_for_timeout(200)
                nfse.fill(nf_buscar)
                main_frame.wait_for_timeout(300)
            except Exception as e:
                self._log(f"  ERRO NFS-e: {e}")
                return "erro"

            try:
                main_frame.locator("#BTNCONSULTAR").click()
                main_frame.wait_for_timeout(4000)
            except Exception as e:
                self._log(f"  ERRO consultar: {e}")
                return "erro"

            try:
                row = main_frame.locator("#GridContainerRow_0001")
                if row.is_visible(timeout=3000):
                    encontrou = True
                    break
            except:
                pass

            # Se não encontrou e é Itabuna/Lauro, tenta com 2025
            if not encontrou and tentativa_nf == 0 and self.loja_key in ("mandarim_itabuna", "mandarim_lauro"):
                nf_buscar = self._formatar_nf_fallback(nota["nf_original"])
                self._log(f"  NF {nf} nao encontrada com 2026, tentando 2025: {nf_buscar}")
                self._expandir_filtro_avancado(main_frame)
            else:
                break

        if not encontrou:
            self._log(f"  NF {nf} NAO ENCONTRADA (tentou 2026 e 2025)")
            return "nao_encontrada"

        if nf_buscar != nf:
            self._log(f"  NF encontrada com prefixo 2025: {nf_buscar}")

        # Captura Valor Total e Saldo da NF no Dealer
        try:
            valor_total_dealer = main_frame.locator("#span_vGRID_TITULO_VALOR_0001").text_content(timeout=2000).strip()
            if valor_total_dealer:
                if self.estado["tabela_nfs"]:
                    self.estado["tabela_nfs"][-1]["valor_total_nota"] = valor_total_dealer
                self._log(f"  Valor Total da Nota (Dealer): {valor_total_dealer}")
        except Exception as e:
            self._log(f"  Aviso: nao capturou valor total: {e}")

        try:
            status_text = main_frame.locator("#span_vGRID_TITULO_STATUS_0001").text_content(timeout=2000).strip()
            self._log(f"  Status: {status_text}")
            if status_text.lower() == "pago":
                self._log(f"  NF {nf} PAGA - pulando")
                return "pago"
        except:
            pass

        # Valida Saldo >= Valor Excel (se saldo < valor, ja foi baixada anteriormente)
        try:
            saldo_texto = main_frame.locator("#span_vGRID_TITULO_SALDO_0001").text_content(timeout=2000).strip()
            saldo_dealer = float(saldo_texto.replace(".", "").replace(",", "."))
            self._log(f"  Saldo (Dealer): {saldo_texto} | Valor Excel: {valor_formatado}")
            if saldo_dealer < valor_excel:
                self._log(f"  SALDO INSUFICIENTE ({saldo_texto} < {valor_formatado}) - Nota baixada anteriormente")
                return "baixada_anteriormente"
        except Exception as e:
            self._log(f"  Aviso: nao conseguiu validar saldo: {e}")

        try:
            main_frame.locator("#vBMPMOVIMENTO_0001").click()
            main_frame.wait_for_timeout(3000)
        except Exception as e:
            self._log(f"  ERRO Movimento: {e}")
            return "erro"

        try:
            popup = self._get_popup_frame(main_frame, tentativas=8)
            if not popup:
                self._log("  ERRO: popup nao abriu")
                return "erro"
            insert_btn = popup.locator("#INSERT")
            if not insert_btn.is_visible(timeout=5000):
                self._log(f"  NF {nf} - sem botao + (ja baixada)")
                try:
                    popup.locator("#FECHAR").click()
                    main_frame.wait_for_timeout(2000)
                except:
                    main_frame.evaluate("document.getElementById('gxp0_cls')?.click()")
                    main_frame.wait_for_timeout(2000)
                return "pago"
            insert_btn.click()
            # Aguarda o formulário carregar (o frame recarrega após INSERT)
            main_frame.wait_for_timeout(4000)
        except Exception as e:
            self._log(f"  ERRO INSERT: {e}")
            return "erro"

        try:
            # Re-encontra o popup pois ele recarrega após INSERT
            popup = self._get_popup_frame(main_frame, procurar_formulario=True, tentativas=8)
            if not popup:
                self._log("  ERRO: formulario nao encontrado apos INSERT")
                return "erro"
            # 1. Tipo Crédito/Débito
            popup.locator("#TITULOMOV_TIPOCDCOD").select_option(value=TIPO_CREDITO_VALUE)
            popup.wait_for_timeout(1500)
            # 2. Tipo de Documento (precisa vir antes do Agente Cobrador)
            popup.locator("#TITULOMOV_TIPODOCUMENTOCOD").select_option(value=TIPO_DOCUMENTO_VALUE)
            popup.wait_for_timeout(500)
            # Clica no Histórico para forçar o blur/update da página
            popup.locator("#TITULOMOV_HISTORICO").click()
            popup.wait_for_timeout(2000)
            # 3. Agente Cobrador (só aparece após o blur do Tipo de Documento)
            popup.locator("#TITULOMOV_AGENTECOBRADORCOD").select_option(value=AGENTE_COBRADOR_VALUE)
            popup.wait_for_timeout(1000)
            # 4. Histórico (com número da NF original)
            nf_original = nota["nf_original"]
            texto_historico = "Baixa Garantia Nota: " + nf_original
            popup.evaluate(f"""
                const h = document.getElementById('TITULOMOV_HISTORICO');
                h.value = '{texto_historico}';
                h.dispatchEvent(new Event('change', {{bubbles: true}}));
                h.dispatchEvent(new Event('input', {{bubbles: true}}));
            """)
            popup.wait_for_timeout(500)
            self._log(f"  Historico: {texto_historico}")
            popup.wait_for_timeout(500)
            # 5. Valor (preenche com o valor do Excel)
            valor_field = popup.locator("#TITULOMOV_VALOR")
            valor_field.click()
            valor_field.fill("")
            popup.wait_for_timeout(200)
            valor_field.fill(valor_formatado)
            popup.wait_for_timeout(500)
            # 6. Empresa - sempre MANDARIM IGUATEMI (matriz), mesmo para Itabuna/Lauro
            try:
                # Usa JavaScript + chamada direta ao gx.evt.onchange do GeneXus
                resultado_empresa = popup.evaluate(f"""
                    (() => {{
                        const sel = document.getElementById('TITULOMOV_EMPRESACOD_MOVIMENTO');
                        if (!sel) return 'nao_encontrado';
                        const valorAntes = sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text : '?';
                        sel.value = '{EMPRESA_IGUATEMI_VALUE}';
                        // Chama o handler do GeneXus diretamente
                        const gxFound = !!(window.gx && window.gx.evt && window.gx.evt.onchange);
                        try {{
                            if (gxFound) {{
                                window.gx.evt.onchange(sel);
                            }}
                        }} catch(e) {{}}
                        sel.dispatchEvent(new Event('change', {{bubbles: true}}));
                        const valorDepois = sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text : '?';
                        return 'antes=' + valorAntes + ' | depois=' + valorDepois + ' | gx=' + gxFound;
                    }})()
                """)
                popup.wait_for_timeout(2500)
                self._log(f"  Empresa: {resultado_empresa}")

                # Aparece popup de confirmacao "Empresa do Titulo e diferente... Sim/Nao"
                # Busca e clica em SIM a partir do TOP DOCUMENT (page.evaluate)
                sim_clicado = main_frame.page.evaluate("""
                    (() => {
                        function clicarSim(docRef, label) {
                            let total = 0;
                            let encontrados = [];
                            try {
                                const btns = docRef.querySelectorAll('button, input[type="button"]');
                                for (const b of btns) {
                                    const texto = (b.textContent || b.value || '').trim();
                                    if (texto === 'Sim') {
                                        encontrados.push(label);
                                        try { b.click(); total++; } catch(e) {}
                                    }
                                }
                                const subIframes = docRef.querySelectorAll('iframe');
                                for (let i = 0; i < subIframes.length; i++) {
                                    try {
                                        if (subIframes[i].contentDocument) {
                                            const r = clicarSim(subIframes[i].contentDocument, label + '>iframe[' + i + ']');
                                            total += r.total;
                                            encontrados = encontrados.concat(r.encontrados);
                                        }
                                    } catch(e) {}
                                }
                            } catch(e) {}
                            return {total, encontrados};
                        }
                        return clicarSim(document, 'top');
                    })()
                """)
                popup.wait_for_timeout(2000)
                self._log(f"  Popup SIM clicado: {sim_clicado}")

                # Verifica valor final da Empresa
                try:
                    valor_final = popup.evaluate("""
                        (() => {
                            const sel = document.getElementById('TITULOMOV_EMPRESACOD_MOVIMENTO');
                            return sel ? sel.value + ':' + (sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text : '?') : 'nao_encontrado';
                        })()
                    """)
                    self._log(f"  Empresa FINAL: {valor_final}")
                except:
                    pass
            except Exception as e_emp:
                self._log(f"  AVISO: nao alterou Empresa: {e_emp}")
            self._log(f"  Formulario preenchido | Valor: {valor_formatado}")
        except Exception as e:
            self._log(f"  ERRO formulario: {e}")
            return "erro"

        try:
            # Re-busca popup pois pode ter perdido referência
            popup = self._get_popup_frame(main_frame, procurar_formulario=True, tentativas=3)
            if not popup:
                self._log("  ERRO: popup perdido antes de confirmar")
                return "erro"
            popup.locator("#TRN_ENTER").click()
            main_frame.wait_for_timeout(5000)
            self._log(f"  >> NF {nf} CONFIRMADA!")
        except Exception as e:
            self._log(f"  ERRO confirmar: {e}")
            return "erro"

        try:
            # Após confirmar, o popup volta pra lista de movimentos. Re-encontra e fecha.
            main_frame.wait_for_timeout(2000)
            popup = self._get_popup_frame(main_frame, tentativas=5)
            if popup:
                fechar = popup.locator("#FECHAR")
                if fechar.is_visible(timeout=5000):
                    fechar.click()
                    main_frame.wait_for_timeout(2000)
        except:
            try:
                # Fallback: fecha pelo X do popup
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

    def _processar_analise(self, main_frame, nota, indice, total):
        """Apenas consulta a NF e extrai valores - não faz baixa."""
        nf = nota["nf"]
        self.estado["nf_atual"] = nf
        self._log(f"[{indice}/{total}] Analisando NF: {nf}")

        self._expandir_filtro_avancado(main_frame)

        # Preenche NFS-e e consulta (para Itabuna/Lauro, tenta 2026 depois 2025)
        nf_buscar = nf
        encontrou = False

        for tentativa_nf in range(2):
            try:
                nfse = main_frame.locator("#vTITULO_NOTAFISCALNRONFSE")
                nfse.click()
                nfse.fill("")
                main_frame.wait_for_timeout(200)
                nfse.fill(nf_buscar)
                main_frame.wait_for_timeout(300)
                main_frame.locator("#BTNCONSULTAR").click()
                main_frame.wait_for_timeout(3500)
            except Exception as e:
                self._log(f"  ERRO consultar: {e}")
                return {"encontrada": False, "erro": str(e)}

            try:
                row = main_frame.locator("#GridContainerRow_0001")
                if row.is_visible(timeout=3000):
                    encontrou = True
                    break
            except:
                pass

            if not encontrou and tentativa_nf == 0 and self.loja_key in ("mandarim_itabuna", "mandarim_lauro"):
                nf_buscar = self._formatar_nf_fallback(nota["nf_original"])
                self._log(f"  NF {nf} nao encontrada com 2026, tentando 2025: {nf_buscar}")
                self._expandir_filtro_avancado(main_frame)
            else:
                break

        if not encontrou:
            self._log(f"  NF {nf} NAO ENCONTRADA (tentou 2026 e 2025)")
            return {"encontrada": False}

        if nf_buscar != nf:
            self._log(f"  NF encontrada com prefixo 2025: {nf_buscar}")

        # Extrai Valor Total e Saldo
        resultado = {"encontrada": True}
        try:
            valor_total = main_frame.locator("#span_vGRID_TITULO_VALOR_0001").text_content(timeout=2000).strip()
            saldo = main_frame.locator("#span_vGRID_TITULO_SALDO_0001").text_content(timeout=2000).strip()
            status = main_frame.locator("#span_vGRID_TITULO_STATUS_0001").text_content(timeout=2000).strip()
            resultado["valor_total"] = valor_total
            resultado["saldo"] = saldo
            resultado["status"] = status
            self._log(f"  Valor Total: {valor_total} | Saldo: {saldo} | Status: {status}")
        except Exception as e:
            self._log(f"  ERRO extrair valores: {e}")
            resultado["erro"] = str(e)

        return resultado

    def executar_analise(self):
        """Executa apenas consultas para análise - NÃO faz baixas."""
        total = len(self.notas)

        with sync_playwright() as pw:
            self._log("Abrindo navegador para ANALISE...")
            browser = pw.chromium.launch(headless=False, slow_mo=200, args=["--start-maximized"])
            context = browser.new_context(no_viewport=True)
            page = context.new_page()
            page.set_default_timeout(15000)

            self._log("Navegando para Dealer.net...")
            page.goto(URL_DEALER, wait_until="networkidle", timeout=60000)
            self._log("FACA LOGIN e va em Titulo a Receber.")

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
                    self._log(f"Dealer pronto!")
                    break
                if tentativa % 10 == 0 and tentativa > 0:
                    self._log(f"Aguardando... ({tentativa*3}s)")
                time.sleep(3)

            if not main_frame:
                self._log("ERRO: Timeout aguardando login.")
                browser.close()
                self.estado["rodando"] = False
                return

            self.estado["dealer_pronto"] = True
            self._log("DEALER PRONTO! Configure o Filtro de Selecao e clique INICIAR ANALISE.")

            while not self.estado.get("inicio_confirmado", False):
                if self.parar:
                    browser.close()
                    self.estado["rodando"] = False
                    return
                time.sleep(1)

            self._log("Iniciando analise...")
            self._expandir_filtro_avancado(main_frame)

            for i, nota in enumerate(self.notas, start=1):
                if self.parar:
                    self._log("Analise PARADA pelo usuario")
                    break

                while self.pausado and not self.parar:
                    time.sleep(0.5)

                if self.parar:
                    break

                # Valor da Baixa Excel (positivo)
                valor_baixa_excel = abs(nota["valor"]) if nota["valor"] else 0

                entrada = {
                    "cnpj": nota["cnpj"],
                    "nf": nota["nf"],
                    "nf_original": nota["nf_original"],
                    "valor_total": "",
                    "saldo": "",
                    "valor_baixa_dealer": 0,
                    "valor_baixa_excel": valor_baixa_excel,
                    "status": "analisando",
                    "mensagem": "Consultando...",
                }
                self.estado["tabela_analise"].append(entrada)

                resultado = self._processar_analise(main_frame, nota, i, total)

                if not resultado.get("encontrada"):
                    entrada["status"] = "nao_encontrada"
                    entrada["mensagem"] = "NF nao encontrada no Dealer"
                else:
                    entrada["valor_total"] = resultado.get("valor_total", "")
                    entrada["saldo"] = resultado.get("saldo", "")
                    # Calcula Valor da Baixa Dealer = Valor Total - Saldo
                    try:
                        vt = float(entrada["valor_total"].replace(".", "").replace(",", "."))
                        sd = float(entrada["saldo"].replace(".", "").replace(",", "."))
                        entrada["valor_baixa_dealer"] = round(vt - sd, 2)
                    except:
                        entrada["valor_baixa_dealer"] = 0
                    entrada["status"] = "analisada"
                    entrada["mensagem"] = "OK"

                # Atualiza progresso
                self.estado["progresso"]["processadas"] = i

                if i % 10 == 0:
                    self._log(f"--- PROGRESSO {i}/{total} ---")

            self._log("=" * 50)
            self._log(f"ANALISE FINALIZADA! {len(self.estado['tabela_analise'])} NFs analisadas")
            self._log("=" * 50)

            self._log("Navegador ficara aberto. Feche manualmente.")
            try:
                while not self.parar:
                    time.sleep(1)
            except:
                pass

            browser.close()
            self.estado["rodando"] = False

    def executar_tudo(self):
        """Abre browser, espera login, e processa todas as NFs. Tudo numa thread."""
        MENSAGENS = {
            "sucesso": "Nota baixada com sucesso",
            "pago": "Nota ja estava paga - pulou",
            "nao_encontrada": "Nota nao encontrada no Dealer",
            "baixada_anteriormente": "Saldo insuficiente - baixada anteriormente",
            "erro": "Erro ao processar nota",
        }

        total = len(self.notas)
        erros_seguidos = 0

        with sync_playwright() as pw:
            self._log("Abrindo navegador...")
            browser = pw.chromium.launch(headless=False, slow_mo=200, args=["--start-maximized"])
            context = browser.new_context(no_viewport=True)
            page = context.new_page()
            page.set_default_timeout(15000)

            self._log("Navegando para Dealer.net...")
            page.goto(URL_DEALER, wait_until="networkidle", timeout=60000)
            self._log("FACA LOGIN no navegador, va em Titulo a Receber e clique OK aqui.")
            self._log("Aguardando voce preparar o Dealer...")

            # Aguarda o usuário fazer login (espera até encontrar o frame correto)
            # O loop vai ficar tentando até encontrar o frame com BTNCONSULTAR
            main_frame = None
            for tentativa in range(120):  # 6 minutos max
                if self.parar:
                    browser.close()
                    self.estado["rodando"] = False
                    return

                # Tenta encontrar o frame
                for frame in page.frames:
                    try:
                        if frame.query_selector("#BTNCONSULTAR"):
                            main_frame = frame
                            break
                    except:
                        pass

                if main_frame:
                    self._log(f"Dealer.net pronto! Frame encontrado: {main_frame.url[:80]}")
                    break

                if tentativa % 10 == 0 and tentativa > 0:
                    self._log(f"Ainda aguardando... ({tentativa*3}s). Faca login e va em Titulo a Receber.")

                time.sleep(3)

            if not main_frame:
                self._log("ERRO: Timeout aguardando login. Tente novamente.")
                browser.close()
                self.estado["rodando"] = False
                return

            # Sinaliza que o Dealer está pronto e espera o usuário confirmar
            self.estado["dealer_pronto"] = True
            self._log("DEALER PRONTO! Configure tudo e clique INICIAR BAIXAS no painel.")
            self._log("Aguardando sua confirmacao...")

            # Aguarda o usuário clicar "INICIAR BAIXAS" no painel web
            while not self.estado.get("inicio_confirmado", False):
                if self.parar:
                    browser.close()
                    self.estado["rodando"] = False
                    return
                time.sleep(1)

            self._log("Confirmacao recebida! Iniciando baixas...")
            self._log(f"Loja: {self.loja_key} | Total: {len(self.notas)} NFs")
            if self.notas:
                self._log(f"Primeira NF: {self.notas[0]}")
                self._log(f"Ultima NF: {self.notas[-1]}")

            # Expande Filtro Avançado
            self._expandir_filtro_avancado(main_frame)

            # Processa NFs
            for i, nota in enumerate(self.notas, start=1):
                if self.parar:
                    self._log("Automacao PARADA pelo usuario")
                    break

                while self.pausado and not self.parar:
                    time.sleep(0.5)

                if self.parar:
                    self._log("Automacao PARADA pelo usuario")
                    break

                entrada = {
                    "cnpj": nota["cnpj"],
                    "nf": nota["nf"],
                    "nf_original": nota["nf_original"],
                    "valor": abs(nota["valor"]) if nota["valor"] else 0,
                    "status": "processando",
                    "mensagem": "Baixando..."
                }
                self.estado["tabela_nfs"].append(entrada)

                resultado = self._processar_nf(main_frame, nota, i, total)
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
                        self._log("MUITOS ERROS SEGUIDOS (5)! Parando.")
                        break
                else:
                    erros_seguidos = 0

                if i % 10 == 0:
                    p = self.estado["progresso"]
                    self._log(f"--- PROGRESSO {i}/{total} | OK:{p['sucesso']} Pago:{p['pago']} Erro:{p['erro']} ---")

            p = self.estado["progresso"]
            self._log("=" * 50)
            self._log(f"FINALIZADO! Sucesso:{p['sucesso']} | Pagas:{p['pago']} | Erros:{p['erro']} | Nao encontradas:{p['nao_encontrada']}")
            self._log("=" * 50)

            self._log("Navegador ficara aberto. Feche manualmente quando quiser.")
            # Mantém o browser aberto
            try:
                while not self.parar:
                    time.sleep(1)
            except:
                pass

            browser.close()
            self.estado["rodando"] = False
