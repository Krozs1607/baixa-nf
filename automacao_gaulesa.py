"""
Módulo de automação de baixa de NF para GAULESA no Dealer.net
Busca por CHASSI e localiza a NF pelo VALOR correspondente.
"""

import time
import logging
import openpyxl
from playwright.sync_api import sync_playwright

# Configurações dos campos
URL_DEALER = "https://workflow.grupoindiana.com.br/Portal/default.html"
TIPO_CREDITO_VALUE = "64"       # RECEBIMENTO DE TÍTULO
AGENTE_COBRADOR_VALUE = "21"    # FABRICA CONTA MOVIMENTO (3.04.60) - APENAS GAULESA
TIPO_DOCUMENTO_VALUE = "2"      # AVISO DE LANCAMENTO
EMPRESA_IGUATEMI_VALUE = "32"   # MANDARIM IGUATEMI (sempre matriz)
HISTORICO_TEXTO = "Baixa Garantia"

log = logging.getLogger("gaulesa")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")


class AutomacaoGaulesa:
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
            chassi = row[0].value   # Coluna A = Chassi
            valor = row[1].value    # Coluna B = Valor Total
            if chassi and valor:
                self.notas.append({
                    "chassi": str(chassi).strip(),
                    "valor": float(valor),
                })
        wb.close()
        total_excel = sum(n["valor"] for n in self.notas)
        self.estado["valor_total_excel"] = total_excel
        self._log(f"Excel Gaulesa carregado: {len(self.notas)} chassis | Total acumulado: R$ {total_excel:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        return len(self.notas)

    def _get_main_frame(self, page):
        for tentativa in range(20):
            for frame in page.frames:
                try:
                    if frame.query_selector("#BTNCONSULTAR"):
                        self._log(f"Frame encontrado: {frame.url[:80]}")
                        return frame
                except:
                    pass
            self._log(f"Tentativa {tentativa+1}/20 - aguardando frames...")
            time.sleep(3)
        raise Exception("Frame nao encontrado!")

    def _get_popup_frame(self, main_frame, procurar_formulario=False, tentativas=5):
        for t in range(tentativas):
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
            chassi = main_frame.query_selector("#vTITULO_VEICULOCHASSI")
            if chassi and chassi.is_visible():
                self._log("Filtro Avancado expandido")
                return
            self._log("AVISO: Expanda o Filtro Avancado manualmente!")
            for i in range(20):
                chassi = main_frame.query_selector("#vTITULO_VEICULOCHASSI")
                if chassi and chassi.is_visible():
                    return
                time.sleep(3)
        except Exception as e:
            self._log(f"Aviso filtro: {e}")

    def _formatar_valor_br(self, valor: float) -> str:
        """Formata valor para comparação com texto do Dealer. Ex: 1503.21 → '1.503,21'"""
        return f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def _parse_valor_br(self, texto: str) -> float:
        """Converte texto BR para float. Ex: '1.503,21' → 1503.21"""
        try:
            return float(texto.replace(".", "").replace(",", "."))
        except:
            return 0.0

    def _varrer_pagina(self, main_frame, valor_excel: float):
        """Percorre as linhas da PAGINA ATUAL. Retorna (linha_match, status, encontrou_pago_nesta_pagina)."""
        encontrou_pago_local = False
        for r in range(1, 20):
            idx = str(r).zfill(4)
            try:
                row = main_frame.locator(f"#GridContainerRow_{idx}")
                if not row.is_visible(timeout=500):
                    break

                valor_span = main_frame.locator(f"#span_vGRID_TITULO_VALOR_{idx}")
                status_span = main_frame.locator(f"#span_vGRID_TITULO_STATUS_{idx}")
                saldo_span = main_frame.locator(f"#span_vGRID_TITULO_SALDO_{idx}")

                valor_texto = valor_span.text_content(timeout=1000).strip()
                status_texto = status_span.text_content(timeout=1000).strip()
                saldo_texto = saldo_span.text_content(timeout=1000).strip()
                valor_dealer = self._parse_valor_br(valor_texto)

                if abs(valor_dealer - valor_excel) < 0.02:
                    if status_texto.lower() == "pago":
                        # Valor bate E status Pago → ja foi pago, retorna imediatamente (nao continua buscando)
                        self._log(f"    Linha {r}: Valor {valor_texto} MATCH | Saldo {saldo_texto} | Status PAGO - ja foi pago, pulando")
                        return 0, "pago", True
                    saldo_dealer = self._parse_valor_br(saldo_texto)
                    if saldo_dealer < valor_excel:
                        self._log(f"    Linha {r}: Valor {valor_texto} MATCH mas SALDO INSUFICIENTE ({saldo_texto} < {self._formatar_valor_br(valor_excel)}) - baixada anteriormente")
                        return 0, "baixada_anteriormente", encontrou_pago_local
                    self._log(f"    Linha {r}: Valor {valor_texto} MATCH! | Saldo {saldo_texto} | Status: {status_texto}")
                    return r, "match", encontrou_pago_local
                else:
                    self._log(f"    Linha {r}: Valor {valor_texto} | Saldo {saldo_texto} | Status {status_texto} (diferente)")
            except:
                break
        return 0, None, encontrou_pago_local

    def _ir_proxima_pagina(self, main_frame) -> bool:
        """Clica em IMGPAGENEXT. Retorna True se avançou para próxima página."""
        try:
            btn_next = main_frame.locator("#IMGPAGENEXT")
            if not btn_next.is_visible(timeout=500):
                return False
            # Verifica se o botão está habilitado (não desabilitado)
            disabled = btn_next.get_attribute("disabled")
            if disabled:
                return False
            btn_next.click()
            main_frame.wait_for_timeout(3000)
            return True
        except:
            return False

    def _encontrar_linha_por_valor(self, main_frame, valor_excel: float):
        """
        Percorre linhas do grid (multiplas paginas) e compara o VALOR com o valor do Excel.
        Se nao encontrar match na pagina atual, vai para a proxima pagina.
        Retorna (numero_linha, status).
        """
        encontrou_pago_global = False
        max_paginas = 10

        for pagina in range(1, max_paginas + 1):
            if pagina > 1:
                self._log(f"  --- Buscando na pagina {pagina} ---")

            linha, status, pago_local = self._varrer_pagina(main_frame, valor_excel)
            if pago_local:
                encontrou_pago_global = True

            # Se achou match, baixada_anteriormente ou pago, retorna imediatamente
            if status in ("match", "baixada_anteriormente", "pago"):
                return linha, status

            # Tenta próxima página
            if not self._ir_proxima_pagina(main_frame):
                break

        if encontrou_pago_global:
            return 0, "pago"
        return 0, "nao_encontrada"

    def _selecionar_documento_controlado(self, main_frame, valor_total_excel):
        """
        Clica na seta azul (IMAGEDOCUMENTOCONTROLADO) abre popup Conta Gerencial,
        encontra o lançamento com valor IGUAL ao valor total do Excel, e clica na setinha verde.
        Retorna True se conseguiu selecionar.
        """
        page = main_frame.page
        self._log(f"  Abrindo selecao de Documento Controlado (total Excel: R$ {self._formatar_valor_br(valor_total_excel)})")

        # Clica na seta azul dentro do popup gxp0_ifrm do formulário
        try:
            popup = self._get_popup_frame(main_frame, procurar_formulario=True, tentativas=3)
            if not popup:
                self._log("  ERRO: popup formulario nao encontrado para clicar seta azul")
                return False
            popup.locator("#IMAGEDOCUMENTOCONTROLADO").click()
            main_frame.wait_for_timeout(3500)
        except Exception as e:
            self._log(f"  ERRO ao clicar seta azul: {e}")
            return False

        # Agora procurar o popup interno com o grid de lançamentos (sel_documentocontrolado.aspx)
        popup_lanc = None
        for t in range(10):
            for frame in page.frames:
                try:
                    url = (frame.url or "").lower()
                    if "sel_documentocontrolado" in url or "documentocontrolado" in url:
                        popup_lanc = frame
                        break
                except:
                    pass
            if popup_lanc:
                break
            time.sleep(1)

        if not popup_lanc:
            # Fallback: procurar por frame que tenha vLINKSELECTION_0001
            for frame in page.frames:
                try:
                    if frame.query_selector("#vLINKSELECTION_0001"):
                        popup_lanc = frame
                        break
                except:
                    pass

        if not popup_lanc:
            self._log("  ERRO: popup de lancamentos nao abriu")
            return False

        self._log(f"  Popup lancamentos aberto")

        # Ajusta filtro de data: 01/01/ANO_ATUAL até hoje (pra pegar lancamentos antigos)
        try:
            import datetime
            hoje = datetime.date.today()
            data_ini = f"01/01/{hoje.year}"
            data_fim = hoje.strftime("%d/%m/%Y")
            campo_ini = popup_lanc.locator("#vFILTRO_DATAINICIO")
            campo_ini.click()
            campo_ini.fill("")
            popup_lanc.wait_for_timeout(200)
            campo_ini.fill(data_ini)
            popup_lanc.wait_for_timeout(300)
            campo_fim = popup_lanc.locator("#vFILTRO_DATAFIM")
            campo_fim.click()
            campo_fim.fill("")
            popup_lanc.wait_for_timeout(200)
            campo_fim.fill(data_fim)
            popup_lanc.wait_for_timeout(300)
            # Clica Consultar
            popup_lanc.locator("#IMGCONSULTA").click()
            popup_lanc.wait_for_timeout(3500)
            self._log(f"  Filtro de data ajustado: {data_ini} ate {data_fim}")
        except Exception as e:
            self._log(f"  AVISO ao ajustar filtro de data: {e}")

        # Percorre linhas e compara valor com valor_total_excel
        linha_match = 0
        for r in range(1, 30):
            idx = str(r).zfill(4)
            try:
                row = popup_lanc.locator(f"#GridContainerRow_{idx}")
                if not row.is_visible(timeout=500):
                    break
                valor_span = popup_lanc.locator(f"#span_vTESOURARIA_VALOR_{idx}")
                valor_texto = valor_span.text_content(timeout=1000).strip()
                valor_lanc = self._parse_valor_br(valor_texto)
                if abs(valor_lanc - valor_total_excel) < 0.02:
                    self._log(f"    Lancamento {r}: Valor {valor_texto} MATCH!")
                    linha_match = r
                    break
                else:
                    self._log(f"    Lancamento {r}: Valor {valor_texto} (diferente de R$ {self._formatar_valor_br(valor_total_excel)})")
            except:
                break

        if linha_match == 0:
            self._log(f"  Nenhum lancamento com valor R$ {self._formatar_valor_br(valor_total_excel)} encontrado")
            return False

        # Clica na setinha verde (vLINKSELECTION_XXXX) da linha correta
        try:
            idx = str(linha_match).zfill(4)
            popup_lanc.locator(f"#vLINKSELECTION_{idx}").click()
            main_frame.wait_for_timeout(3000)
            self._log(f"  >> Documento Controlado selecionado (linha {linha_match})")
            return True
        except Exception as e:
            self._log(f"  ERRO ao selecionar linha: {e}")
            return False

    def _coletar_autorizadas_todas_paginas(self, main_frame):
        """Percorre todas as paginas coletando NFs com status Autorizado (saldo > 0)."""
        todas = []
        for pagina in range(1, 11):
            for r in range(1, 20):
                idx = str(r).zfill(4)
                try:
                    row = main_frame.locator(f"#GridContainerRow_{idx}")
                    if not row.is_visible(timeout=500):
                        break
                    valor_texto = main_frame.locator(f"#span_vGRID_TITULO_VALOR_{idx}").text_content(timeout=1000).strip()
                    status_texto = main_frame.locator(f"#span_vGRID_TITULO_STATUS_{idx}").text_content(timeout=1000).strip()
                    saldo_texto = main_frame.locator(f"#span_vGRID_TITULO_SALDO_{idx}").text_content(timeout=1000).strip()
                    valor = self._parse_valor_br(valor_texto)
                    saldo = self._parse_valor_br(saldo_texto)
                    if status_texto.lower() != "pago" and saldo > 0.01:
                        todas.append({"valor": valor, "saldo": saldo, "pagina": pagina, "linha": r, "valor_texto": valor_texto})
                except:
                    break
            if not self._ir_proxima_pagina(main_frame):
                break
        return todas

    def _encontrar_combinacao_soma(self, autorizadas, valor_alvo, max_tamanho=4):
        """Tenta combinacoes de 2 ate max_tamanho NFs que somam ao valor_alvo."""
        from itertools import combinations
        n = len(autorizadas)
        for tamanho in range(2, min(max_tamanho + 1, n + 1)):
            for combo in combinations(autorizadas, tamanho):
                soma = sum(nf["valor"] for nf in combo)
                if abs(soma - valor_alvo) < 0.02:
                    return list(combo)
        return None

    def _buscar_chassi_reset(self, main_frame, chassi):
        """Refaz busca por chassi para resetar estado (volta pra pagina 1)."""
        try:
            self._expandir_filtro_avancado(main_frame)
            campo = main_frame.locator("#vTITULO_VEICULOCHASSI")
            campo.click()
            campo.fill("")
            main_frame.wait_for_timeout(200)
            campo.fill(chassi)
            main_frame.wait_for_timeout(300)
            main_frame.locator("#BTNCONSULTAR").click()
            main_frame.wait_for_timeout(4000)
            return True
        except Exception as e:
            self._log(f"  ERRO rebusca chassi: {e}")
            return False

    def _fazer_baixa_em_nf(self, main_frame, chassi, valor_baixa, valor_total_excel):
        """
        Faz a baixa numa NF com valor_baixa especifico.
        Assume que estamos na tela de resultados da busca pelo chassi.
        Procura linha com status Autorizado e valor==valor_baixa em todas as paginas.
        Retorna True se baixou com sucesso.
        """
        # Encontra a linha
        linha_match = 0
        for pagina in range(1, 11):
            for r in range(1, 20):
                idx = str(r).zfill(4)
                try:
                    row = main_frame.locator(f"#GridContainerRow_{idx}")
                    if not row.is_visible(timeout=500):
                        break
                    valor_texto = main_frame.locator(f"#span_vGRID_TITULO_VALOR_{idx}").text_content(timeout=1000).strip()
                    status_texto = main_frame.locator(f"#span_vGRID_TITULO_STATUS_{idx}").text_content(timeout=1000).strip()
                    saldo_texto = main_frame.locator(f"#span_vGRID_TITULO_SALDO_{idx}").text_content(timeout=1000).strip()
                    valor_dealer = self._parse_valor_br(valor_texto)
                    saldo_dealer = self._parse_valor_br(saldo_texto)
                    if abs(valor_dealer - valor_baixa) < 0.02 and status_texto.lower() != "pago" and saldo_dealer >= valor_baixa - 0.02:
                        linha_match = r
                        break
                except:
                    break
            if linha_match:
                break
            if not self._ir_proxima_pagina(main_frame):
                break

        if linha_match == 0:
            self._log(f"    ERRO: linha com valor {self._formatar_valor_br(valor_baixa)} nao achada no grid")
            return False

        idx = str(linha_match).zfill(4)
        self._log(f"    Processando baixa linha {linha_match} | Valor: {self._formatar_valor_br(valor_baixa)}")

        # Clicar Movimento
        try:
            main_frame.locator(f"#vBMPMOVIMENTO_{idx}").click()
            main_frame.wait_for_timeout(3000)
        except Exception as e:
            self._log(f"    ERRO Movimento: {e}")
            return False

        # Clicar INSERT
        try:
            popup = self._get_popup_frame(main_frame, tentativas=8)
            if not popup:
                self._log("    ERRO: popup nao abriu")
                return False
            insert_btn = popup.locator("#INSERT")
            if not insert_btn.is_visible(timeout=5000):
                self._log(f"    sem botao + (ja baixada)")
                try:
                    popup.locator("#FECHAR").click()
                    main_frame.wait_for_timeout(2000)
                except:
                    pass
                return False
            insert_btn.click()
            main_frame.wait_for_timeout(4000)
        except Exception as e:
            self._log(f"    ERRO INSERT: {e}")
            return False

        # Preencher formulário
        try:
            popup = self._get_popup_frame(main_frame, procurar_formulario=True, tentativas=8)
            if not popup:
                self._log("    ERRO: formulario nao encontrado")
                return False
            popup.locator("#TITULOMOV_TIPOCDCOD").select_option(value=TIPO_CREDITO_VALUE)
            popup.wait_for_timeout(1500)
            popup.locator("#TITULOMOV_TIPODOCUMENTOCOD").select_option(value=TIPO_DOCUMENTO_VALUE)
            popup.wait_for_timeout(500)
            popup.locator("#TITULOMOV_HISTORICO").click()
            popup.wait_for_timeout(2000)
            popup.locator("#TITULOMOV_AGENTECOBRADORCOD").select_option(value=AGENTE_COBRADOR_VALUE)
            popup.wait_for_timeout(1000)
            texto_historico = f"Baixa Garantia Chassi: {chassi}"
            popup.evaluate(f"""
                (() => {{
                    const h = document.getElementById('TITULOMOV_HISTORICO');
                    h.value = '{texto_historico}';
                    h.dispatchEvent(new Event('change', {{bubbles: true}}));
                    h.dispatchEvent(new Event('input', {{bubbles: true}}));
                }})()
            """)
            popup.wait_for_timeout(500)
            valor_formatado = f"{valor_baixa:.2f}".replace(".", ",")
            valor_field = popup.locator("#TITULOMOV_VALOR")
            valor_field.click()
            valor_field.fill("")
            popup.wait_for_timeout(200)
            valor_field.fill(valor_formatado)
            popup.wait_for_timeout(500)
            self._log(f"    Formulario preenchido | Valor: {valor_formatado}")
        except Exception as e:
            self._log(f"    ERRO formulario: {e}")
            return False

        # Documento Controlado
        if valor_total_excel and valor_total_excel > 0:
            ok_doc = self._selecionar_documento_controlado(main_frame, valor_total_excel)
            if not ok_doc:
                self._log(f"    ERRO: nao selecionou Documento Controlado")
                return False

        # Confirmar
        try:
            popup = self._get_popup_frame(main_frame, procurar_formulario=True, tentativas=3)
            if not popup:
                self._log("    ERRO: popup perdido antes de confirmar")
                return False
            popup.locator("#TRN_ENTER").click()
            main_frame.wait_for_timeout(5000)
            self._log(f"    >> Baixa CONFIRMADA!")
        except Exception as e:
            self._log(f"    ERRO confirmar: {e}")
            return False

        # Fechar popup
        try:
            main_frame.wait_for_timeout(2000)
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

        return True

    def _processar_chassi(self, main_frame, nota, indice, total):
        chassi = nota["chassi"]
        valor_excel = nota["valor"]
        self.estado["nf_atual"] = chassi
        self._log(f"[{indice}/{total}] Chassi: {chassi} | Valor Excel: {self._formatar_valor_br(valor_excel)}")

        self._expandir_filtro_avancado(main_frame)

        # PASSO 1: Preencher Chassi e consultar
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

        # PASSO 2: Verificar se encontrou algum resultado
        try:
            row = main_frame.locator("#GridContainerRow_0001")
            if not row.is_visible(timeout=3000):
                self._log(f"  Chassi {chassi} NAO ENCONTRADO no Dealer")
                return "nao_encontrada"
        except:
            self._log(f"  Chassi {chassi} NAO ENCONTRADO")
            return "nao_encontrada"

        # PASSO 3: Percorrer linhas e encontrar a que tem valor igual ao Excel
        self._log(f"  Buscando NF com valor {self._formatar_valor_br(valor_excel)}...")
        linha_match, match_status = self._encontrar_linha_por_valor(main_frame, valor_excel)

        if match_status == "pago":
            self._log(f"  Valor {self._formatar_valor_br(valor_excel)} encontrado mas JA PAGO")
            return "pago"

        if match_status == "baixada_anteriormente":
            self._log(f"  Saldo insuficiente - nota baixada anteriormente")
            return "baixada_anteriormente"

        if match_status == "nao_encontrada":
            # Tenta encontrar combinacao de NFs Autorizadas que somem ao valor_excel
            self._log(f"  Valor unico nao encontrado. Tentando combinacoes de NFs Autorizadas...")
            if not self._buscar_chassi_reset(main_frame, chassi):
                return "nao_encontrada"
            autorizadas = self._coletar_autorizadas_todas_paginas(main_frame)
            self._log(f"  Coletadas {len(autorizadas)} NFs Autorizadas para analise")
            if not autorizadas:
                self._log(f"  Sem NFs Autorizadas - pulando")
                return "nao_encontrada"
            combinacao = self._encontrar_combinacao_soma(autorizadas, valor_excel, max_tamanho=4)
            if not combinacao:
                self._log(f"  Nenhuma combinacao soma R$ {self._formatar_valor_br(valor_excel)} - pulando")
                return "nao_encontrada"

            valores_combo = [self._formatar_valor_br(nf["valor"]) for nf in combinacao]
            self._log(f"  COMBINACAO ENCONTRADA ({len(combinacao)} NFs): {' + '.join(valores_combo)} = R$ {self._formatar_valor_br(valor_excel)}")

            # Faz baixa em cada NF da combinacao
            valor_total_excel = self.estado.get("valor_total_excel", 0)
            baixas_ok = 0
            for i, nf_combo in enumerate(combinacao, start=1):
                self._log(f"  --- Baixa {i}/{len(combinacao)} da combinacao (Valor: {self._formatar_valor_br(nf_combo['valor'])}) ---")
                # Rebusca chassi para resetar estado
                if not self._buscar_chassi_reset(main_frame, chassi):
                    self._log(f"    ERRO: nao conseguiu rebuscar chassi")
                    break
                ok = self._fazer_baixa_em_nf(main_frame, chassi, nf_combo["valor"], valor_total_excel)
                if ok:
                    baixas_ok += 1
                else:
                    self._log(f"    ERRO na baixa {i} da combinacao")
                    break

            if baixas_ok == len(combinacao):
                self._log(f"  >> TODAS as {baixas_ok} baixas da combinacao CONFIRMADAS!")
                return "sucesso"
            else:
                self._log(f"  ERRO: {baixas_ok}/{len(combinacao)} baixas feitas - combinacao parcial")
                return "erro"

        idx = str(linha_match).zfill(4)

        # Captura valor total da nota
        try:
            valor_total = main_frame.locator(f"#span_vGRID_TITULO_VALOR_{idx}").text_content(timeout=2000).strip()
            if self.estado["tabela_nfs"]:
                self.estado["tabela_nfs"][-1]["valor_total_nota"] = valor_total
        except:
            pass

        # PASSO 4: Clicar Movimento na linha correta
        try:
            main_frame.locator(f"#vBMPMOVIMENTO_{idx}").click()
            main_frame.wait_for_timeout(3000)
        except Exception as e:
            self._log(f"  ERRO Movimento: {e}")
            return "erro"

        # PASSO 5: Clicar INSERT
        try:
            popup = self._get_popup_frame(main_frame, tentativas=8)
            if not popup:
                self._log("  ERRO: popup nao abriu")
                return "erro"
            insert_btn = popup.locator("#INSERT")
            if not insert_btn.is_visible(timeout=5000):
                self._log(f"  Chassi {chassi} - sem botao + (ja baixada)")
                try:
                    popup.locator("#FECHAR").click()
                    main_frame.wait_for_timeout(2000)
                except:
                    pass
                return "pago"
            insert_btn.click()
            main_frame.wait_for_timeout(4000)
        except Exception as e:
            self._log(f"  ERRO INSERT: {e}")
            return "erro"

        # PASSO 6: Preencher formulário
        try:
            popup = self._get_popup_frame(main_frame, procurar_formulario=True, tentativas=8)
            if not popup:
                self._log("  ERRO: formulario nao encontrado")
                return "erro"

            # 1. Tipo Crédito/Débito
            popup.locator("#TITULOMOV_TIPOCDCOD").select_option(value=TIPO_CREDITO_VALUE)
            popup.wait_for_timeout(1500)
            # 2. Tipo de Documento
            popup.locator("#TITULOMOV_TIPODOCUMENTOCOD").select_option(value=TIPO_DOCUMENTO_VALUE)
            popup.wait_for_timeout(500)
            # Clica no Histórico para forçar blur
            popup.locator("#TITULOMOV_HISTORICO").click()
            popup.wait_for_timeout(2000)
            # 3. Agente Cobrador
            popup.locator("#TITULOMOV_AGENTECOBRADORCOD").select_option(value=AGENTE_COBRADOR_VALUE)
            popup.wait_for_timeout(1000)
            # 4. Histórico
            texto_historico = f"Baixa Garantia Chassi: {chassi}"
            popup.evaluate(f"""
                (() => {{
                    const h = document.getElementById('TITULOMOV_HISTORICO');
                    h.value = '{texto_historico}';
                    h.dispatchEvent(new Event('change', {{bubbles: true}}));
                    h.dispatchEvent(new Event('input', {{bubbles: true}}));
                }})()
            """)
            popup.wait_for_timeout(500)
            # 5. Valor
            valor_formatado = f"{valor_excel:.2f}".replace(".", ",")
            valor_field = popup.locator("#TITULOMOV_VALOR")
            valor_field.click()
            valor_field.fill("")
            popup.wait_for_timeout(200)
            valor_field.fill(valor_formatado)
            popup.wait_for_timeout(500)
            # Gaulesa: NAO altera a Empresa (usa a empresa original da NF)

            self._log(f"  Formulario preenchido | Valor: {valor_formatado}")
        except Exception as e:
            self._log(f"  ERRO formulario: {e}")
            return "erro"

        # PASSO 6b: Selecionar Documento Controlado pelo valor TOTAL do Excel
        valor_total_excel = self.estado.get("valor_total_excel", 0)
        if valor_total_excel and valor_total_excel > 0:
            ok_doc = self._selecionar_documento_controlado(main_frame, valor_total_excel)
            if not ok_doc:
                self._log(f"  ERRO: nao conseguiu selecionar Documento Controlado")
                return "erro"
        else:
            self._log(f"  AVISO: valor_total_excel nao disponivel - pulando Documento Controlado")

        # PASSO 7: Confirmar
        try:
            popup = self._get_popup_frame(main_frame, procurar_formulario=True, tentativas=3)
            if not popup:
                self._log("  ERRO: popup perdido antes de confirmar")
                return "erro"
            popup.locator("#TRN_ENTER").click()
            main_frame.wait_for_timeout(5000)
            self._log(f"  >> Chassi {chassi} CONFIRMADA!")
        except Exception as e:
            self._log(f"  ERRO confirmar: {e}")
            return "erro"

        # PASSO 8: Fechar popup
        try:
            main_frame.wait_for_timeout(2000)
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

    def executar_tudo(self):
        """Abre browser, espera login, e processa todas as notas Gaulesa."""
        MENSAGENS = {
            "sucesso": "Baixada com sucesso",
            "pago": "Ja estava paga - pulou",
            "nao_encontrada": "Chassi/valor nao encontrado",
            "baixada_anteriormente": "Saldo insuficiente - baixada anteriormente",
            "erro": "Erro ao processar",
        }
        total = len(self.notas)
        erros_seguidos = 0

        with sync_playwright() as pw:
            self._log("Abrindo navegador GAULESA...")
            browser = pw.chromium.launch(headless=False, slow_mo=200, args=["--start-maximized"])
            context = browser.new_context(no_viewport=True)
            page = context.new_page()
            page.set_default_timeout(15000)

            self._log("Navegando para Dealer.net...")
            page.goto(URL_DEALER, wait_until="networkidle", timeout=60000)
            self._log("FACA LOGIN, selecione GAULESA IGUATEMI e va em Titulo a Receber.")

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
                self._log("ERRO: Timeout aguardando login.")
                browser.close()
                self.estado["rodando"] = False
                return

            self.estado["dealer_pronto"] = True
            self._log("DEALER PRONTO! Configure Filtro de Selecao e clique INICIAR.")

            while not self.estado.get("inicio_confirmado", False):
                if self.parar:
                    browser.close()
                    self.estado["rodando"] = False
                    return
                time.sleep(1)

            self._log("Iniciando baixas Gaulesa...")
            self._expandir_filtro_avancado(main_frame)

            for i, nota in enumerate(self.notas, start=1):
                if self.parar:
                    self._log("Automacao PARADA")
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
                    "valor_total_nota": "",
                    "status": "processando",
                    "mensagem": "Baixando...",
                }
                self.estado["tabela_nfs"].append(entrada)

                resultado = self._processar_chassi(main_frame, nota, i, total)
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
                    self._log(f"--- PROGRESSO {i}/{total} | OK:{p['sucesso']} Pago:{p['pago']} Erro:{p['erro']} ---")

            p = self.estado["progresso"]
            self._log("=" * 50)
            self._log(f"GAULESA FINALIZADA! Sucesso:{p['sucesso']} | Pagas:{p['pago']} | Erros:{p['erro']} | Nao encontradas:{p['nao_encontrada']}")
            self._log("=" * 50)

            try:
                while not self.parar:
                    time.sleep(1)
            except:
                pass
            browser.close()
            self.estado["rodando"] = False
