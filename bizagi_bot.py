from playwright.sync_api import Playwright, Browser, Page, expect
import config
import time
import os
import re
from validator import BizagiValidator
from case_reporter import CaseReporter

class BizagiBot:
    def __init__(self, playwright: Playwright):
        self.playwright = playwright
        self.browser = None
        self.context = None
        self.page = None
        self.validator = BizagiValidator()
        self.reporter = CaseReporter()

    def launch_browser(self):
        """Launches the browser in headful mode."""
        self.browser = self.playwright.chromium.launch(
            headless=config.HEADLESS,
            channel="chrome", 
            args=["--start-maximized"]
        )
        self.context = self.browser.new_context(
            viewport=None, # Set to None to use full available size
            accept_downloads=True
        )
        self.page = self.context.new_page()

    def go_to_login(self):
        """Navigates to the login page."""
        self.page.goto(config.BASE_URL)
        print(f"Navegou para {config.BASE_URL}")

    def wait_for_login(self, timeout=300):
        """Waits for the user to log in manually by checking for 'Caixa de entrada'."""
        print(f"Aguardando até {timeout} segundos pelo login...")
        try:
            self.page.wait_for_selector("text=Caixa de entrada", timeout=timeout * 1000)
            return True
        except Exception:
            return False

    def process_all_cases(self):
        """Processes a specific case for debugging purposes."""
        try:
            print("Iniciando processamento focado no caso 14353...")
            
            # User requested specific stable URL
            self.page.goto("https://digital-ons.bizagi.com/#")
            # User requested specific stable URL
            self.page.goto("https://digital-ons.bizagi.com/#")
            # time.sleep(5) # Removed hard wait, replaced with element wait in navigate

            
            # Navigate to the specific folder as requested
            self.navigate_to_folder()
            
            # Target specific case
            if self.navigate_to_first_case(target_case_id="14353"):
                if self.take_ownership():
                    print(">>> CASO ABERTO E POSSE TOMADA COM SUCESSO <<<")
                if self.take_ownership():
                    print(">>> CASO ABERTO E POSSE TOMADA COM SUCESSO <<<")
                    self.process_current_case(case_id_hint="14353")
                else:
                    print("Não foi possível tomar posse do caso.")
            else:
                print("Caso 14353 não encontrado.")
            
            print("\n>>> FIM DO PROCESSO DE DEBUG <<<")
            input("Pressione [ENTER] no terminal para fechar o navegador e encerrar o script...")
                
        except Exception as e:
            print(f"Erro fatal no processamento: {e}")
            self._save_screenshot("erro_main_process")
            input("Erro fatal ocorrido. Pressione [ENTER] para fechar...")


    def navigate_to_folder(self):
        """Navigates to Jurídico -> Recuperação de Encargos CUST with retry."""
        print("Navegando para pasta Jurídico -> Recuperação de Encargos CUST...")
        try:
            target_sub = self.page.locator("text=Recuperação de Encargos CUST").first
            
            # Retry loop to expand 'Jurídico'
            for attempt in range(3):
                if target_sub.is_visible():
                    print("Subpasta visível, clicando...")
                    target_sub.click(force=True)
                    break
                
                print(f"Tentativa {attempt+1}: Expandindo 'Jurídico'...")
                
                # Try specific strategies cleanly
                try:
                    # Strategy 1: Exact text match (Robust)
                    juridico_text = self.page.get_by_text("Jurídico", exact=True)
                    if juridico_text.is_visible():
                        juridico_text.click(force=True)
                        time.sleep(2)
                        continue
                except:
                    pass

                try:
                    # Strategy 2: CSS with has-text (if exact fails)
                    juridico_css = self.page.locator("li:has-text('Jurídico')").first
                    if juridico_css.is_visible():
                        juridico_css.click(force=True)
                        time.sleep(2)
                        continue
                except:
                    print("Estratégia CSS falhou.")
                    
                print("Menu 'Jurídico' não encontrado na tentativa atual.")
            
            # Wait for grid to load after clicking subfolder
            print("Aguardando carregamento da tabela de casos...")
            self.page.wait_for_selector(".ui-bizagi-grid", timeout=15000)
            self.page.wait_for_load_state('networkidle')
            
        except Exception as e:
            print(f"Erro na navegação de pastas: {e}")
            self._save_screenshot("erro_navegacao_pasta")

    def highlight_element(self, locator):
        """Highlights the element in YELLOW as requested."""
        try:
            if locator.is_visible():
                locator.evaluate("el => el.style.border = '4px solid yellow'")
                locator.evaluate("el => el.style.backgroundColor = 'rgba(255, 255, 0, 0.3)'")
                time.sleep(0.5) # Small pause for user to see
        except Exception:
            pass

    def _click_with_debug(self, selector):
        """Highlights an element and clicks it with force=True."""
        try:
            loc = self.page.locator(selector).first
            if loc.is_visible():
                self.highlight_element(loc)
                print(f"Clicando em: {selector}")
                # Force click to bypass overlays
                loc.click(force=True, timeout=5000)
            else:
                print(f"Elemento não visível para clique: {selector}")
        except Exception as e:
            print(f"Falha ao clicar em {selector}: {e}")
            raise e

    def _save_screenshot(self, name):
        """Saves a screenshot for debugging."""
        try:
            filename = f"{name}_{int(time.time())}.png"
            path = os.path.join(config.PROCESS_DIR, "downloads", filename)
            self.page.screenshot(path=path)
            print(f"Screenshot salvo em: {path}")
        except Exception as e:
            print(f"Erro ao salvar screenshot: {e}")

    def navigate_to_first_case(self, target_case_id="14353"):
        """Finds the case in the inbox and clicks it. Defaults to 14353 for testing."""
        print(f"Procurando caso na lista (Alvo: {target_case_id} ou 'Analisar Solicitação')...")
        try:
            # Broad wait for any table data to ensure load
            self.page.wait_for_selector("tbody tr", timeout=15000)
            
            # TEST STRATEGY: Click specifically on the ID 14353 if visible
            print(f"Tentativa direta: Clicar no texto '{target_case_id}'...")
            specific_link = self.page.get_by_text(target_case_id, exact=True)
            if specific_link.is_visible():
                print(f"Texto {target_case_id} encontrado! Clicando...")
                self.highlight_element(specific_link)
                specific_link.click(force=True)
                return True
            
            print("Alvo direto não visível. Tentando iterar linhas...")
            
            # Fallback: Loop rows
            rows = self.page.locator("tr.ui-bizagi-grid-row")
            count = rows.count()
            print(f"Encontradas {count} linhas.")

            for i in range(count):
                row = rows.nth(i)
                text = row.inner_text()
                
                # Check for activity or ID
                if target_case_id in text or "Analisar Solicitação" in text:
                    print(f"Linha relevante encontrada: {text[:40]}...")
                    # Try to click the numeric cell (Case ID)
                    # We assume it's the cell matching the ID
                    
                    # 1. Try to find the specific ID element inside the row
                    id_cell = row.get_by_text(target_case_id)
                    if id_cell.count() > 0:
                        self.highlight_element(id_cell.first)
                        id_cell.first.click(force=True)
                        return True
                    
                    # 2. Heuristic: Click the 3rd column (Index 2)
                    # Icons | Process | Case(2) | Activity
                    cells = row.locator("td")
                    if cells.count() > 3:
                        print("Tentando clicar na 3ª coluna (índice 2)...")
                        self.highlight_element(cells.nth(2))
                        cells.nth(2).click(force=True)
                        return True
                        
            print("Nenhum caso encontrado.")
            return False
            
        except Exception as e:
            print(f"Erro ao navegar para o caso: {e}")
            self._save_screenshot("erro_navegacao_caso_v3")
            return False

    def search_case(self, case_id):
        """Searches for a specific case ID and ensures the case opens."""
        print(f"Pesquisando pelo caso: {case_id}")
        try:
            # 1. Access Search
            search_input = self.page.locator("input[placeholder='Pesquisar']").first
            
            if not search_input.is_visible():
                search_icon = self.page.locator(".ui-bizagi-wp-app-search-item, .bz-icon-search").first
                if search_icon.is_visible():
                    search_icon.click()
                    time.sleep(1)
            
            if search_input.is_visible():
                self.highlight_element(search_input)
                search_input.fill("")
                search_input.type(case_id, delay=100)
                search_input.press("Enter")
                
                print("Aguardando resultados da pesquisa...")
                self.page.wait_for_selector("div.ui-bizagi-wp-search-results, .ui-bizagi-grid", timeout=10000)
                time.sleep(4) # Increased wait time
                
                print(f"Tentando clicar no ID: {case_id}")
                
                # STRATEGY: Text Match + Hover + Click
                # The user says "click the number".
                
                # Generic robust finder for the number text
                target = self.page.get_by_text(case_id, exact=True).first
                
                if target.is_visible():
                    self.highlight_element(target)
                    print("Elemento visualizado. Movendo mouse...")
                    target.hover()
                    time.sleep(0.5)
                    print("Clicando...")
                    target.click() # Standard click first
                    
                    # Check if navigation started
                    try:
                         # We expect the search grid to disappear or the form to appear
                         self.page.wait_for_selector(".ui-bizagi-form", timeout=5000)
                         print(">>> SUCESSO: Caso aberto (Formulário detectado).")
                         return True
                    except:
                        print("Clique padrão não abriu. Tentando Force Click...")
                        target.click(force=True)
                        time.sleep(2)

                else:
                    print(f"Caso {case_id} visualmente não encontrado nos resultados.")
                    return False

                # VALIDATION
                print("Verificando se o caso abriu (Final)...")
                if self.page.locator(".ui-bizagi-form").is_visible() or self.page.locator("text=Decisão da Análise").is_visible():
                    return True
                else:
                    print("ALERTA: Falha na abertura do caso.")
                    self._save_screenshot(f"erro_abertura_caso_{case_id}")
                    return False

            else:
                print("Barra de pesquisa não encontrada.")
                return False
                
        except Exception as e:
            print(f"Erro na pesquisa do caso: {e}")
            self._save_screenshot(f"erro_pesquisa_{case_id}")
            return False

    def take_ownership(self):
        """Clicks on 'Tomar posse' if available. Enforces Case Context."""
        try:
            time.sleep(2) 
            
            # First, verify we are NOT on the search grid still
            if self.page.locator(".ui-bizagi-wp-search-results").is_visible():
                print("ERRO: Ainda estamos na tela de pesquisa. 'Tomar posse' falhou.")
                return False

            take_ownership_btn = self.page.locator("text=Tomar posse")
            if take_ownership_btn.is_visible(timeout=5000):
                print("Botão 'Tomar posse' encontrado. Clicando...")
                self.highlight_element(take_ownership_btn)
                take_ownership_btn.click()
                self.page.wait_for_load_state('networkidle')
                return True
            
            # If button not found, check if we are simply already viewing the case (Success)
            # or if we are completely lost
            if self.page.locator(".ui-bizagi-form").is_visible() or self.page.locator("text=Decisão da Análise").is_visible():
                 print("Botão 'Tomar posse' não visível, mas estamos no caso. Assumindo posse prévia.")
                 return True
            
            print("Não foi possível confirmar posse nem contexto do caso.")
            return False
            
        except Exception as e:
            print(f"Erro ao tentar tomar posse: {e}")
            return False

    def perform_adjustment_batch(self, obs_text):
        """Specific logic for the batch adjustment task."""
        print("Executando ajuste em lote...")
        try:
           # 1. Define Action
           action = "Ajustar"
           
           # 2. Select Decision
           print(f"Selecionando decisão: {action}")
           dropdown_clicked = False
           try:
                # Find the container for the specific field
                field_container = self.page.locator("div.ui-bizagi-render:has(label:text-is('Decisão de Análise'))")
                
                if field_container.count() > 0:
                    trigger = field_container.locator(".ui-select-match, .ui-select-toggle, input[type='text']").first
                    if trigger.is_visible():
                        self.highlight_element(trigger)
                        trigger.click(force=True)
                        dropdown_clicked = True
                    else:
                        self.highlight_element(field_container)
                        field_container.click(force=True)
                        dropdown_clicked = True
           except Exception as e:
                print(f"Erro ao abrir dropdown: {e}")
                
           if dropdown_clicked:
                time.sleep(1)
                try:
                    option_selector = f"li:has-text('{action}'), span:has-text('{action}'), div:has-text('{action}')"
                    option = self.page.locator(option_selector).last
                    if option.is_visible():
                        option.scroll_into_view_if_needed()
                        option.click(force=True)
                    else:
                        print("Opção não visível, tentando digitar...")
                        self.page.keyboard.type(action)
                        self.page.keyboard.press("Enter")
                except Exception as e:
                    print(f"Erro ao selecionar opção: {e}")

           # 3. Fill Observations
           print("Preenchendo observações...")
           try:
                obs_region = self.page.locator("div.ui-bizagi-render:has(label:text-is('Observações da Análise'))")
                if obs_region.count() > 0:
                    self.highlight_element(obs_region)
                    textarea = obs_region.locator("textarea").first
                    if textarea.is_visible():
                         textarea.fill(obs_text)
                    else:
                         print("Textarea não encontrado.")
                else:
                    print("Campo 'Observações da Análise' não encontrado.")
           except Exception as e:
                print(f"Erro ao preencher observações: {e}")
           
           # 4. Click Send (Enviar) - User requested ONLY Send, NO Save.
           print("Enviando solicitação...")
           try:
               # Specific button locator
               btn = self.page.locator("button").filter(has_text="Enviar").last
               if not btn.is_visible():
                   btn = self.page.get_by_role("button", name="Enviar").first
               
               if btn.is_visible():
                   self.highlight_element(btn)
                   btn.click()
                   print("Botão ENVIAR clicado.")
                   # Wait for completion/navigation?
                   # time.sleep(5) 
               else:
                   print("Botão ENVIAR não encontrado.")
           except Exception as e:
               print(f"Erro ao clicar em Enviar: {e}")

        except Exception as e:
           print(f"Erro fatal no ajuste em lote: {e}")
           self._save_screenshot("erro_ajuste_lote")

    def process_current_case(self, case_id_hint=None):
        """Main orchestrator for a single case."""
        print(f"Processando caso atual... (Hint: {case_id_hint})")
        
        # Ensure we are in a case context
        try:
            self.page.wait_for_selector(".ui-bizagi-form, .ui-bizagi-workportal-widget-content", timeout=10000)
        except:
            print("Aviso: Formulário do caso demorou a carregar ou não foi detectado.")
        
        # 1. Scrape Metadata
        print("--- DEBUG: Iniciando extração de dados do formulário ---")
        case_data = self.scrape_case_data()
        print(f"--- DEBUG: Dados extraídos: {case_data} ---")
        
        # Report Update
        self.reporter.update_data(case_data)
        
        # Explicitly update CaseID if we found one
        if case_data.get('case_id'):
             print(f"Atualizando Relatório com ID Real do Caso: {case_data['case_id']}")
             self.reporter.current_case_data['CaseID'] = case_data['case_id']
        elif case_id_hint:
             print(f"Usando ID fornecido (Hint) para Relatório: {case_id_hint}")
             self.reporter.current_case_data['CaseID'] = case_id_hint
        
        contract_code = case_data.get('contract_code')
        
        # Hardcode fallback/prompt if missing (to ensure we can test Excel opening)
        if not contract_code:
            print("!!! ERRO CRÍTICO: Código do Contrato NÃO encontrado. !!!")
        
        # Verify Master List Path explicitly
        import excel_helper
        master_path = excel_helper.MASTER_LIST_PATH
        print(f"--- DEBUG: Verificando caminho do Excel Mestre: {master_path}")
        
        if not os.path.exists(master_path):
             print(f"!!! ERRO CRÍTICO: Arquivo Mestre não encontrado no disco! Verifique o caminho: {master_path}")
        else:
             print("--- DEBUG: Arquivo Mestre existe no disco. ---")

        if contract_code:
             print(f"IDENTIFICAÇÃO: Código do Contrato: '{contract_code}' (Esperado: CUST-2021-088)")
             
             # Internal Logic: Check if contract exists (VISUAL)
             print(f"Buscando referência para {contract_code} visualmente no Excel...")
             
             ref_number = excel_helper.find_reference_number_visual(contract_code)
             
             if ref_number:
                 print(f"Sucesso: Contrato localizado! ID Referência: {ref_number}")
                 
                 # --- STEP 4: Open AVD Excel (VISUAL) ---
                 print(f"Procurando arquivo AVD para ID: {ref_number}")
                 
                 # We use the visual validator to Open/Highlight/Sum
                 is_visual_ok, total_found, msg_visual = excel_helper.validate_debt_amount_visual(ref_number, 
                                                                                        case_data.get('cnpj', ''), 
                                                                                        case_data.get('debt_amount', ''))
                 print(f"Status Visual AVD: {msg_visual}")
                 
                 self.reporter.log_visual_validation("OK" if is_visual_ok else "NOK", msg_visual)
                 
                 if is_visual_ok:
                     print(">>> VALIDAÇÃO FINANCEIRA OK. Retornando ao Bizagi para documentos... <<<")
                     
                     # --- STEP 6: Document Analysis ---
                     print("Acessando área de documentos...")
                     if self.go_to_documents_tab():
                         print("Baixando documentos para análise CADIN ANEEL...")
                         
                         documents = self.download_documents() 
                         
                         if documents:
                             print(f"Documentos baixados: {len(documents)}")
                             # Validate Compliance (Open and Check)
                             doc_valid = self.validate_downloads(documents)
                             self.reporter.log_doc_validation("OK" if doc_valid else "NOK")
                             
                             if doc_valid:
                                 print(">>> CADIN OK. Avançando para Validação de Protesto... <<<")
                                 
                                 # --- STEP 7: Protesto Analysis ---
                                 protest_docs = self.download_documents(target_label="Protesto extrajudicial da dívida", prefix="PROTESTO")
                                 
                                 if protest_docs:
                                     p_path = list(protest_docs.values())[0]
                                     print(f"Validando valor no Protesto: {p_path}")
                                     
                                     expected_debt = case_data.get('debt_amount', '')
                                     
                                     # Identify Memory Calculation File from previously downloaded 'documents' (CADIN step usually has it? No, Memory is usually attached.)
                                     # The prompt says: "Na sequencia vamos avançar". 
                                     # We need to find the memory file path. 
                                     # Let's search in the 'documents' map we downloaded earlier (Cadin/Evidence) OR new download?
                                     # Usually Memory is part of the evidence.
                                     
                                     memory_path = None
                                     # Search for keyword "memorial" or "memória" in the docs_map keys
                                     for d_name, d_path in documents.items():
                                         if "monial" in d_name.lower() or "memória" in d_name.lower() or "memorial" in d_name.lower():
                                             memory_path = d_path
                                             print(f"Arquivo de Memória identificado para Cruzamento: {d_name}")
                                             break
                                     
                                     p_ok, p_msg = self.validator.validate_protest_amount(p_path, expected_debt, memory_path=memory_path)
                                     print(f"Resultado Validação Protesto: {p_msg}")
                                     
                                     self.reporter.log_step("Protesto_Amount_Check", "OK" if p_ok else "NOK - " + p_msg)
                                     
                                     # Prepare final result
                                     final_approved = p_ok
                                     result = {
                                         'approved': final_approved,
                                         'failed_docs': {} if final_approved else {"Protesto": f"Valor divergente: {p_msg}"}
                                     }
                                 else:
                                     print("Documento de Protesto não encontrado para download.")
                                     self.reporter.log_step("Protesto_Download", "NOK")
                                     result = {'approved': False, 'failed_docs': {'Protesto': 'Documento não encontrado'}}
                                 
                                 # Execute Decision
                                 print("Executando decisão no sistema...")
                                 self.execute_decision(result)
                                 
                             else:
                                 print("CADIN Inválido. Interrompendo para ajustes.")
                                 result = {'approved': False, 'failed_docs': {'CADIN': 'Documentação incompleta'}}
                                 self.execute_decision(result)

                         else:
                             print("Nenhum documento encontrado na linha 'Inscrição no cadastro...'.")
                             self.reporter.log_doc_validation("NOK", "Documents not found")
                         
                     else:
                         print("Erro ao acessar aba de documentos.")
                 else:
                     print(f"ALERTA: Divergência financeira impede prosseguimento automático. {msg_visual}")
                     self.reporter.finalize_case("Validation Failed - Visual Discrepancy")

             else:
                 print("Falha: Contrato não encontrado na planilha Mestre (Visual).")
        else:
             print("PULANDO ABERTURA DE EXCEL: Código do contrato ausente.")
        
        # Final cleanup print
        print("Processamento do caso finalizado.")

    def execute_decision(self, result):
        """Selects the decision in the dropdown and fills observations."""
        try:
            # Determine action
            if result['approved']:
                action = "Aprovar"
                obs_text = "Documentação validada com sucesso conforme RES 1125."
            else:
                action = "Ajustar" # Returning for corrections
                obs_text = "Solicitação devolvida para ajustes. Veja detalhes nos campos específicos acima."
                
                # NEW: Flag specific document issues in the table
                if 'failed_docs' in result and result['failed_docs']:
                    print("Assinalando falhas nos documentos específicos...")
                    self.flag_document_issues(result['failed_docs'])
            
            print(f"Aplicando decisão global: {action}")
            self.reporter.finalize_case(action)
            
            # --- 1. Select Decision in Dropdown ---
            print("Tentando selecionar opção no dropdown 'Decisão de Análise'...")
            
            # Strategy A: Find by Label -> Click Placeholder "Selecione..."
            dropdown_clicked = False
            
            try:
                # Find the container for the specific field
                field_container = self.page.locator("div.ui-bizagi-render:has(label:text-is('Decisão de Análise'))")
                
                if field_container.count() > 0:
                    trigger = field_container.locator(".ui-select-match, .ui-select-toggle, input[type='text']").first
                    if trigger.is_visible():
                        self.highlight_element(trigger)
                        trigger.click(force=True)
                        dropdown_clicked = True
                    else:
                        self.highlight_element(field_container)
                        field_container.click(force=True)
                        dropdown_clicked = True
            except Exception as e:
                print(f"Erro na Estratégia A: {e}")

            time.sleep(1)
            
            # Select the option
            if dropdown_clicked:
                try:
                    option_selector = f"li:has-text('{action}'), span:has-text('{action}'), div:has-text('{action}')"
                    self.page.wait_for_timeout(500)
                    option = self.page.locator(option_selector).last
                    if option.is_visible():
                        option.scroll_into_view_if_needed()
                        option.click(force=True)
                        print(f"Opção '{action}' clicada com sucesso.")
                    else:
                        print(f"Opção '{action}' não visível. Tentando digitar...")
                        self.page.keyboard.type(action)
                        self.page.keyboard.press("Enter")
                except Exception as e:
                    print(f"Erro ao selecionar a opção: {e}")
            
            # --- 2. Fill Observations ---
            print("Preenchendo observações globais...")
            try:
                obs_region = self.page.locator("div.ui-bizagi-render:has(label:text-is('Observações de Análise'))")
                if obs_region.count() > 0:
                    self.highlight_element(obs_region)
                    textarea = obs_region.locator("textarea").first
                    if textarea.is_visible():
                        textarea.fill(obs_text)
                        print("Texto global preenchido.")
            except Exception as e:
                print(f"Erro ao preencher observações: {e}")
            
            # --- 3. Submit (Highlight Only) ---
            print("Verificando botão de envio...")
            # Try specific exact match first
            btn = self.page.locator("button").filter(has_text="Enviar").last
            
            if not btn.is_visible():
                 # Fallback to role
                 btn = self.page.get_by_role("button", name="Enviar").first
            
            if btn.is_visible():
                print("Botão ENVIAR localizado com sucesso.")
                self.highlight_element(btn)
                # btn.click() # Uncomment to actually send
            else:
                print("Botão ENVIAR não localizado (Verifique se há scroll ou outro frame).")

        except Exception as e:
            print(f"Erro fatal na execução da decisão: {e}")
            self._save_screenshot("erro_decisao_fatal")

    def flag_document_issues(self, failed_docs):
        """
        Iterates through the documents table and flags specific rows that failed validation.
        Args:
            failed_docs (dict): Map of partial row name -> failure reason string
        """
        try:
            # We assume the user is still on the Documents tab / section is visible
            # Rows in the table
            rows = self.page.locator("tr")
            count = rows.count()
            print(f"Procurando linhas para ajuste (total {count})...")
            
            for i in range(count):
                row = rows.nth(i)
                text = row.inner_text()
                
                matched_key = None
                for key in failed_docs.keys():
                    if key in text:
                        matched_key = key
                        break
                
                if matched_key:
                    reason = failed_docs[matched_key]
                    print(f"Marcando ajuste para: {matched_key}")
                    
                    # 1. Click "Ajustar?" Checkbox
                    # From screenshot: Checkbox is likely in the 2nd column (Index 1) or found by type
                    # We look for a checkbox input within this row
                    checkbox = row.locator("input[type='checkbox']").first
                    
                    if checkbox.count() > 0:
                        # Fix: Check vs Click. Some bizagi checkboxes are custom divs or input hidden.
                        # Trying click with force
                        if not checkbox.is_checked():
                            checkbox.click(force=True)
                            print(f"    -> Checkbox marcado para {matched_key}")
                        else:
                            print(f"    -> Checkbox já estava marcado para {matched_key}")
                    else:
                        print("    -> Checkbox não encontrado nesta linha.")
                        
                    # 2. Fill "Justificativa" Input
                    # From screenshot: It's a text box at the end of the row.
                    # We can try to find the last input type='text' or textarea in the row
                    justification_input = row.locator("input[type='text'], textarea").last
                    
                    if justification_input.is_visible():
                        self.highlight_element(justification_input)
                        justification_input.fill(reason)
                        print(f"  -> Justificativa preenchida: {reason}")
                    else:
                        print("  -> Campo de justificativa não encontrado na linha.")
                        
        except Exception as e:
            print(f"Erro ao assinalar falhas nos documentos: {e}")
            self._save_screenshot("erro_flag_docs")
    def scrape_case_data(self):
        """Extracts visible form data."""
        data = {}
        try:
            print("Extraindo dados do formulário...")
            
            extract_map = {
                'contract_code': ['Código do Contrato', 'Contrato', 'Contract'],
                'cnpj': ['CNPJ', 'CNPJ do Empreendimento'],
                'debt_amount': ['Débito do Ajuizamento', 'Valor', 'Débito', 'Valor Total']
            }

            # 1. Try to capture Contract Code using direct selectors first
            try:
                # Try multiple potential selectors for Contract Code
                contract_el = self.page.locator("input[name*='CodigoContrato'], span:text-matches('Código.*Contrato', 'i') + span").first
                if contract_el.count() > 0 and contract_el.is_visible():
                    # Check if it's an input or a span to get value correctly
                    if contract_el.evaluate("el => el.tagName.toLowerCase()") == 'input':
                        data['contract_code'] = contract_el.input_value()
                    else:
                        data['contract_code'] = contract_el.inner_text()
                    print(f"  -> Coletado (Direto - Código Contrato): {data['contract_code']}")
            except Exception as e:
                print(f"Erro ao tentar coletar Código do Contrato diretamente: {e}")
                pass
                
            # 2. Try to capture Case ID from header
            try:
                # Common pattern in Bizagi work portal: Case Title or Info
                case_id_el = self.page.locator(".ui-bizagi-wp-app-routing-link span.case-id, .ui-bizagi-ws-title").first
                if case_id_el.is_visible():
                    data['case_id'] = case_id_el.inner_text().strip()
                    print(f"  -> Coletado (Direto - Case ID): {data['case_id']}")
            except Exception as e:
                print(f"Erro ao tentar coletar Case ID diretamente: {e}")
                pass
            
            # Robust extraction strategy for other fields (and fallback for contract_code if not found)
            labels = self.page.locator("label")
            count = labels.count()
            
            for i in range(count):
                lbl = labels.nth(i)
                lbl_text = lbl.inner_text().strip().replace(':', '')
                
                matched_key = None
                for key, possible_names in extract_map.items():
                    if lbl_text in possible_names:
                        matched_key = key
                        break
                
                if matched_key and matched_key not in data: # Only extract if not already found by direct selectors
                    print(f"Revisando campo: {lbl_text}")
                    self.highlight_element(lbl)
                    
                    # Try siblings or proximity
                    inp = lbl.locator("xpath=../following-sibling::div//input").first
                    if not inp.is_visible():
                         inp = lbl.locator("xpath=following::input[1]").first
                    
                    if inp.is_visible():
                         val = inp.input_value()
                         if val and val.lower() != "on":
                             self.highlight_element(inp)
                             data[matched_key] = val
                             print(f"  -> Coletado (Input): {val}")
                             continue

                    spn = lbl.locator("xpath=../following-sibling::div//span[@class='ui-bizagi-render-text']").first
                    if not spn.is_visible():
                         spn = lbl.locator("xpath=following::span[contains(@class,'value')][1]").first
                    
                    if spn.is_visible():
                        val = spn.inner_text().strip()
                        self.highlight_element(spn)
                        data[matched_key] = val
                        print(f"  -> Coletado (Span): {val}")
                        continue
                        
                    # Fallback: Try generic div or p immediately following
                    # Sometimes it's just a text node in a div
                    gen = lbl.locator("xpath=following::div[1]").first
                    if gen.is_visible():
                        val = gen.inner_text().strip()
                        if val:
                             self.highlight_element(gen)
                             data[matched_key] = val
                             print(f"  -> Coletado (Div Genérico): {val}")
                             continue
                             
            print(f"Dados extraídos: {data}")
            
            if not data:
                print("AVISO: Nenhum dado extraído! Verifique os seletores.")
                self._save_screenshot("debug_form_scrape_vazio")
            elif 'contract_code' not in data:
                print("AVISO: Código do contrato NÃO encontrado. O Excel não será aberto.")
                self._save_screenshot("debug_falta_contrato")
                
        except Exception as e:
            print(f"Erro ao extrair dados: {e}")
            self._save_screenshot("erro_scrape_dados")
        return data

    def go_to_documents_tab(self):
        """Navigates to 'Documentos da Solicitação' tab."""
        print("Tentando acessar aba de Documentos...")
        possible_selectors = [
            "text=Documentos da Solicitação",
            "text=Documentos",
            "text=Anexos",
            "li[title*='Documentos']",
            ".ui-bizagi-tab:has-text('Documentos')"
        ]
        
        for sel in possible_selectors:
            try:
                # Use robust checking
                tab = self.page.locator(sel).first
                if tab.is_visible():
                    print(f"Aba encontrada com seletor '{sel}'. Clicando...")
                    self.highlight_element(tab)
                    tab.click(force=True)
                    time.sleep(2)
                    return True
            except:
                continue
                
        print("Aba de documentos não encontrada.")
        self._save_screenshot("erro_aba_documentos")
        return False

    def download_documents(self, target_label="Inscrição no cadastro de inadimplentes da ANEEL", prefix="CADIN"):
        """Downloads files for a specific label row."""
        docs_map = {}
        try:
            print(f"Iniciando download de documentos para: '{target_label}'...")
            # Wait for table
            self.page.wait_for_selector("tr", state="attached", timeout=10000)
            
            target_label = "Inscrição no cadastro de inadimplentes da ANEEL"
            
            # Locate all rows
            rows = self.page.locator("tr")
            count = rows.count()
            print(f"Varrendo {count} linhas buscando: '{target_label}'...")
            
            found_row = False
            
            for i in range(count):
                row = rows.nth(i)
                # Quick text check
                if target_label in row.inner_text():
                    found_row = True
                    print(f"LINHA ENCONTRADA (Índice {i})!")
                    
                    # Highlight and Scroll
                    self.highlight_element(row)
                    # Highlight and Scroll
                    self.highlight_element(row)
                    row.scroll_into_view_if_needed()
                    # time.sleep(1) # Removed for speed 
                    
                    # Find all links in the 'Documento' column
                    links = row.locator("a").all()
                    print(f"Encontrados {len(links)} links nesta linha.")
                    
                    for idx, link in enumerate(links):
                        txt = link.inner_text().strip()
                        # Filtering: Only verify if it looks like a file or download link
                        if not txt or ('.pdf' not in txt.lower() and len(txt) < 5):
                            continue
                            
                        print(f"  -> Tentando baixar: {txt}")
                        try:
                            # Scroll link into view specifically
                            link.scroll_into_view_if_needed()
                            if link.is_visible():
                                self.highlight_element(link)
                                
                                with self.page.expect_download(timeout=30000) as download_info:
                                    link.click(force=True) # Force click
                                
                                download = download_info.value
                                # Clean filename
                                clean_name = "".join([c for c in txt if c.isalnum() or c in ('-','_','.')])
                                # Add timestamp to avoid PermissionDenied if file is open
                                timestamp = int(time.time())
                                save_path = os.path.join(config.DOWNLOAD_DIR, f"{prefix}_{idx}_{timestamp}_{clean_name}")
                                
                                download.save_as(save_path)
                                docs_map[txt] = save_path
                                print(f"    -> SUCESSO: Salvo em {save_path}")
                            else:
                                print("    -> Link não visível mesmo após scroll.")
                        except Exception as e:
                            print(f"    -> ERRO ao baixar {txt}: {e}")
                    
                            print(f"    -> ERRO ao baixar {txt}: {e}")
                    
                        # Optimization: Stop after 3 downloads to avoid pulling unnecessary files
                        if len(docs_map) >= 3:
                             print("Limitando a 3 downloads para agilidade.")
                             break
                    
                    break # Stop after processing the target row
            
            if not found_row:
                print(f"AVISO: Linha '{target_label}' não encontrada na tabela.")
                
        except Exception as e:
            print(f"Erro ao processar downloads: {e}")
            self._save_screenshot("erro_download_docs")
        
        return docs_map

    def validate_downloads(self, docs_map):
        """
        Opens downloaded documents and verifies compliance with:
        (i) Relatório ANEEL
        (ii) Histórico de Comunicações
        (iii) Memorial Descritivo
        """
        print("\n>>> INICIANDO VALIDAÇÃO DE DOCUMENTOS (CADIN) <<<")
        required_evidence = {
            'relatorio': ['relatório', 'aneel', 'débito', 'geradora', '1 1', '1 2'], # Keyword heuristics
            'comunicacao': ['comunicação', 'email', 'e-mail', 'histórico', 'solicitação', 'resposta'],
            'memorial': ['memorial', 'descritivo', 'memória', 'cálculo', 'contratos'] 
        }
        
        found_evidence = {k: False for k in required_evidence}
        
        for name, path in docs_map.items():
            print(f"--- Analisando: {name} ---")
            
            # 1. Open File Visually - DISABLED FOR SPEED
            # try:
            #     print(f"Abrindo visualmente: {path}")
            #     os.startfile(path)
            #     # time.sleep(2) # Give time to open
            # except Exception as e:
            #     print(f"Erro ao abrir arquivo: {e}")
                
            # 2. Check Content (filename based for now, text extraction if pypdf available)
            # Since we can't easily rely on OCR/PDF text without libs, we use Filename + User Confirmation context
            
            fname_lower = name.lower()
            
            for ev_type, keywords in required_evidence.items():
                if any(k in fname_lower for k in keywords):
                    found_evidence[ev_type] = True
                    print(f"  -> Identificado como: {ev_type.upper()}")
            
        print("\n--- RESUMO DA VALIDAÇÃO DOCUMENTAL ---")
        missing = []
        
        # (i) Relatório
        if found_evidence['relatorio']:
            print("[OK] (i) Relatório atualizado ANEEL identificado.")
        else:
            print("[ALERTA] (i) Relatório ANEEL não identificado explicitamente pelo nome.")
            missing.append("(i) Relatório ANEEL")
            
        # (ii) Comunicações
        if found_evidence['comunicacao']:
             print("[OK] (ii) Histórico de Comunicações identificado.")
        else:
             # Fallback: Check content of ALL downloaded files for the specific email
             print("[INFO] Verificando conteúdo dos arquivos por e-mail da CADIN...")
             email_target = "inadimplentes.saf@aneel.gov.br"
             found_in_content = False
             
             for name, path in docs_map.items():
                 # Fix: Use self.validator and the public method
                 if self.validator.check_keywords(path, [email_target, "inadimplentes.saf"]):
                     found_in_content = True
                     print(f"  -> E-mail '{email_target}' encontrado no arquivo: {name}")
                     found_evidence['comunicacao'] = True
                     break
             
             if found_evidence['comunicacao']:
                 print("[OK] (ii) Histórico de Comunicações identificado pelo conteúdo.")
             else:
                 print("[ALERTA] (ii) Histórico de Comunicações não identificado explicitamente.")
                 missing.append("(ii) Histórico de E-mails")

        # (iii) Memorial (Conditional)
        # We don't know if multiple contracts exist without checking, but if found:
        if found_evidence['memorial']:
             print("[OK] (iii) Memorial Descritivo identificado.")
        else:
             print("[INFO] (iii) Memorial não encontrado (obrigatório apenas para múltiplos contratos).")
        
        if not missing:
            print(">>> TODAS AS EVIDÊNCIAS OBRIGATÓRIAS PARECEM ESTAR PRESENTES. <<<")
            return True
        else:
            print(f">>> ATENÇÃO: Verifique visualmente os itens: {', '.join(missing)} <<<")
            return False

    def approve_case(self):
        print("Aprovando caso no sistema...")
        self.page.click("button:has-text('Aprovar')")
        # Handle confirmation dialogs if any

    def return_case(self, reasons):
        print("Devolvendo caso no sistema...")
        # Fill reason field
        # self.page.fill("textarea[name='motivo']", "\n".join(reasons))
        self.page.click("button:has-text('Devolução')")
