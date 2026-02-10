import time
from playwright.sync_api import sync_playwright

# CONFIGURAÇÃO
BASE_URL = "https://digital-ons.bizagi.com/"
OBS_TEXT = """Prezado (a), 
Solicitamos o seguinte ajuste documental: 
Cadastro de inadimplentes da ANEEL – disponibilizar os seguintes documentos: 
(i) relatório emitido pela ANEEL contendo os débitos inscritos da transmissora; 
(ii) cópia do histórico de comunicações eletrônicas (solicitação e resposta da Agência); 
At.,"""

# Placeholder list - User will interpret this
CASE_LIST = []

def run_batch():
    # Ask for cases at runtime
    print(">>> CONFIGURAÇÃO DO LOTE <<<")
    case_input = input("Cole a lista de casos separados por vírgula (ou pressione Enter para usar o padrão): ")
    if not case_input.strip():
        # Default list (Update Step 88)
        current_list = [
            "46874", "46876", "46878", "46882", "46888",
            "46890", "46891", "46892", "46894", "46896"
        ]
        print(f"Usando lista padrão de {len(current_list)} casos.")
    else:
        current_list = [c.strip() for c in case_input.split(",") if c.strip()]
        print(f"Lista carregada com {len(current_list)} casos.")

    # Production Mode - Run all cases automatically
    print("MODO DE PRODUÇÃO: Executando lote completo.")
    # test_mode logic removed

    with sync_playwright() as p:
        # Layout Fix: viewport=None lets the browser take up the maximized window size
        browser = p.chromium.launch(headless=False, args=["--start-maximized"])
        context = browser.new_context(viewport=None) 
        page = context.new_page()

        print("\n>>> INICIANDO AUTOMAÇÃO V3 (Corrigida + Robustez Extrema) <<<")
        
        page.goto(BASE_URL)
        
        # 0. REWRITTEN AUTO-LOGIN LOGIC (Recorded Flow Strategy)
        print("Verificando necessidade de login (Estratégia Input + Enter)...")
        try:
             # Wait for either Dashboard OR Login Page
             print("Aguardando carregamento da página inicial...")
             page.wait_for_selector("#menuQuery, input[placeholder='Pesquisar'], #i0116", state="visible", timeout=15000)
             
             if page.locator("#menuQuery").is_visible() or page.locator("input[placeholder='Pesquisar']").is_visible():
                  print(">>> JÁ ESTAMOS LOGADOS! (Dashboard detectado)")
             else:
                  print(">>> TELA DE LOGIN DETECTADA. Iniciando autenticação...")
                  
                  # 1. Fill Email & Enter
                  email = "eduardo.personale@ons.org.br"
                  print(f"Preenchendo email: {email}")
                  # Try specific ID first, then fallback
                  if page.locator("#i0116").is_visible():
                      page.fill("#i0116", email)
                  else:
                      page.fill("input[type='email']", email)
                  
                  print("Pressionando Enter para avançar...")
                  page.keyboard.press("Enter")
                  
                  # 2. Fill Password & Submit
                  password = "Clara31102016@"
                  print("Aguardando campo de senha...")
                  try:
                      # Wait for password field (i0118)
                      page.wait_for_selector("#i0118, input[type='password']", state="visible", timeout=10000)
                      page.fill("#i0118, input[type='password']", password)
                      
                      page.wait_for_timeout(500)
                      print("Submetendo senha...")
                      
                      # Strategy: Press Enter first
                      page.keyboard.press("Enter")
                      page.wait_for_timeout(1500)
                      
                      # Redundancy: If button is still there, CLICK IT
                      # idSIButton9 is the standard MS login button ID
                      submit_btn = page.locator("#idSIButton9, input[type='submit'][value='Entrar'], button:has-text('Entrar')")
                      if submit_btn.is_visible():
                           print("    -> Botão 'Entrar' ainda visível. Forçando clique...")
                           submit_btn.click()
                           
                  except Exception as e:
                      print(f"Erro na etapa de senha: {e}")
                  
                  # 3. Handle 'Stay Signed In' (Sim/Yes)
                  try:
                      print("Aguardando tela 'Manter conectado'...")
                      # Wait specifically for the "Yes" button ID
                      yes_btn = page.wait_for_selector("#idSIButton9", state="visible", timeout=10000)
                      
                      if yes_btn:
                          print("Clicando em 'Sim' (Force)...")
                          yes_btn.click(force=True)
                          
                          # Redundancy: Press Enter if still visible after slight delay
                          page.wait_for_timeout(500)
                          if yes_btn.is_visible():
                               print("    -> Botão ainda visível. Pressionando Enter...")
                               page.keyboard.press("Enter")
                  except:
                      print("Tela 'Manter conectado' não apareceu ou foi pulada.")
                      
                  # 4. Final Wait for Dashboard
                  print("Aguardando carregamento final do Dashboard...")
                  page.wait_for_load_state("networkidle")
                  
        except Exception as e:
             print(f"ERRO DE LOGIN OU TIMEOUT: {e}")
             print("Tentando seguir, verifique se o navegador está aberto...")

        # Robust Login Wait (Main App)
        try:
            page.wait_for_selector("input[placeholder='Pesquisar'], a:has-text('Caixa de entrada')", timeout=60000)
            print("Login confirmado (Dashboard visível)!")
        except:
            print("Login não detectado a tempo. Verifique se o site abriu corretamente.")
            return
            
        for case_id in current_list:
            if page.is_closed():
                print("ERRO FATAL: O navegador foi fechado. Encerrando o script.")
                break

            print(f"\n==========================================")
            print(f"PROCESSANDO CASO: {case_id}")
            print(f"==========================================")
            
            try:
                # 1. Ensure Inbox/Search Ready Strategy: HARD RESET
                # We force a reload to clear any previous search filters or stuck UI states.
                print("Navegando para o Início (Reset de Estado)...")
                try:
                    page.goto(BASE_URL)
                    page.wait_for_load_state("networkidle")
                except Exception as e_nav:
                    print(f"Erro no reload: {e_nav}")
                    if page.is_closed():
                         print("Navegador fechado durante reload. Encerrando.")
                         break
                    print("Tentando prosseguir mesmo com erro de navegação...")

                # 2. Search Strategy: LOOP UNTIL FOUND
                print(f"Pesquisando {case_id}...")
                
                search_success = False
                for s_attempt in range(3):
                    # Check if already present on grid
                    if page.locator(f"tr:has-text('{case_id}')").is_visible():
                         print("Caso já visível no grid.")
                         search_success = True
                         break
                    
                    search_box = page.get_by_role("textbox", name="Pesquisar")
                    # Ensure search box is open
                    if not search_box.is_visible():
                        btn_search = page.locator(".bz-icon-search, button.bz-search-button").first
                        if btn_search.is_visible():
                             btn_search.click()
                        page.wait_for_timeout(500)
                    
                    # Fill
                    search_box.click()
                    search_box.fill("")
                    search_box.type(case_id, delay=100)
                    page.wait_for_timeout(500)
                    
                    # TRIGGER: Click Magnifying Glass (Most Reliable)
                    btn_trigger = page.locator(".bz-search-button, button.bz-icon-search").first
                    if btn_trigger.is_visible():
                        print("Clicando na lupa para pesquisar...")
                        btn_trigger.click()
                    else:
                        print("Pressionando Enter...")
                        search_box.press("Enter")
                    
                    # Wait for grid update
                    try:
                        page.wait_for_selector(f"tr:has-text('{case_id}')", timeout=5000)
                        search_success = True
                        break
                    except:
                        print(f"Pesquisa tentativa {s_attempt+1} não retornou o caso ainda. Retentando...")
                        page.wait_for_timeout(1000)
                
                if not search_success and not page.locator(f"tr:has-text('{case_id}')").is_visible():
                     print("AVISO: Caso não encontrado na pesquisa. Tentando abrir mesmo assim se aparecer...")
                
                # 3. Open Case (Standardized from run_approve_batch.py)
                print("Abrindo caso...")
                case_opened = False
                max_attempts = 4
                
                for attempt in range(1, max_attempts + 1):
                    # Check Success
                    try:
                        # Fixed Selector Syntax (Split into OR logic)
                        if page.locator(".ui-bizagi-form").or_(page.locator("text=Tomar posse")).is_visible():
                                case_opened = True
                                break
                    except: pass
                    
                    # Find and Click
                    target_link = page.get_by_text(case_id, exact=True).first
                    if not target_link.is_visible():
                         target_link = page.locator(f"tr:has-text('{case_id}')").first
                    
                    if target_link.is_visible():
                        try:
                            target_link.click(force=(attempt > 1))
                            if attempt == 3: target_link.dblclick(force=True)
                        except: pass
                        
                        # Wait for form
                        for i in range(20):
                            page.wait_for_timeout(1000)
                            if page.locator(".ui-bizagi-form").or_(page.locator("text=Tomar posse")).is_visible():
                                case_opened = True
                                break
                        if case_opened: break
                    else:
                        page.wait_for_timeout(3000)

                if not case_opened:
                    print(f"ERRO: Não foi possível abrir o caso {case_id}.")
                    continue 

                # 4. Wait & Scroll (Standardized from run_approve_batch.py)
                page.wait_for_timeout(3000)
                page.mouse.wheel(0, 500)
                
                # 5. Tomar Posse (Standardized from run_approve_batch.py)
                take_ownership_btn = page.get_by_role("link", name="Tomar posse")
                if take_ownership_btn.is_visible(timeout=5000):
                    print("Tomando posse...")
                    take_ownership_btn.click()
                    page.wait_for_load_state("networkidle")
                    time.sleep(2)

                # 6. Fill Decision - Enforce Wait for Form Reload
                print("Aguardando formulário de decisão ficar ativo...")
                form_ready = False
                
                # Loop to check if form is ready (Decisão field is the key)
                for wait_cycle in range(10): # Wait up to 10 seconds
                    if page.locator(".ui-selectmenu-btn, .ui-select-toggle").is_visible() or \
                       page.get_by_role("combobox", name="Decisão da Análise").is_visible():
                        form_ready = True
                        break
                    time.sleep(1)
                    print(f"    -> Aguardando reload do form ({wait_cycle+1}/10)...")
                
                if not form_ready:
                    print("    -> Form não carregou automaticamente. Tentando Reload da página...")
                    page.reload()
                    page.wait_for_load_state("networkidle")
                    time.sleep(3)
                
                print("Preenchendo decisão 'Ajustar' (Estratégia Menu Click)...")
                
                # RECORODED STRATEGY: Click the arrow/menu button first
                # page.locator(".ui-selectmenu-btn").click()
                try:
                    menu_btn = page.locator(".ui-selectmenu-btn, .ui-select-toggle, .biz-btn-caret").first
                    if menu_btn.is_visible():
                        menu_btn.click()
                        time.sleep(1)
                        
                        # Now try to click the option 'Ajustar' directly if visible, or use Arrows
                        option_ajustar = page.locator("li:has-text('Ajustar'), .ui-select-choices-row:has-text('Ajustar')").first
                        if option_ajustar.is_visible():
                            option_ajustar.click()
                        else:
                            # Fallback to typing in the combo if menu didn't expose options
                            page.get_by_role("combobox", name="Decisão da Análise").type("Ajustar", delay=100)
                            page.keyboard.press("Enter")
                    else:
                        print("Botão de menu dropdown não encontrado. Tentando interação direta...")
                        page.get_by_role("combobox", name="Decisão da Análise").click()
                        page.keyboard.type("Ajustar", delay=100)
                        page.keyboard.press("Enter")
                        
                    page.wait_for_timeout(1000)

                except Exception as e_dec:
                    print(f"Erro ao interagir com decisão: {e_dec}")


                # 6.5 CHECKBOX "Ajustar?" Logic (Refined Step 136)
                print("Verificando necessidade de marcar checkbox 'Ajustar?'...")
                try:
                    # Target specific document - Partial match strict
                    doc_name_part = "Inscrição no cadastro de inadimplentes da ANEEL"
                    # Use xpath or css that handles the text loosely
                    doc_row = page.locator(f"tr:has-text('{doc_name_part}')").first
                    
                    if doc_row.is_visible():
                        print(f"    -> Linha do documento encontrada.")
                        doc_row.scroll_into_view_if_needed()
                        page.wait_for_timeout(500)

                        # Locate checkbox - Generic strategy
                        # Usually it's an input with type checkbox OR a custom ui-checkbox div
                        checkbox_locator = doc_row.locator(".ui-checkbox, input[type='checkbox']").first
                        
                        if checkbox_locator.is_visible():
                            # Determine if checked
                            is_checked = False
                            
                            # Check class on label if it's the custom wrapper
                            label = doc_row.locator("label.ui-checkbox-label, .ui-checkbox label").first
                            if label.count() > 0:
                                class_attr = label.get_attribute("class") or ""
                                is_checked = "ui-checkbox-state-checked" in class_attr
                            else:
                                # Standard input check
                                is_checked = checkbox_locator.is_checked()
                            
                            if is_checked:
                                print(f"    -> Checkbox já marcado.")
                            else:
                                print("    -> Checkbox NÃO marcado. Clicando...")
                                try:
                                    if label.count() > 0:
                                        label.click(force=True)
                                    else:
                                        checkbox_locator.click(force=True)
                                    page.wait_for_timeout(500)
                                except Exception as click_err:
                                    print(f"    -> Falha ao clicar: {click_err}")
                        else:
                            print("    -> Checkbox não encontrado na linha.")
                    else:
                        print(f"    -> Documento '{doc_name_part}' não visível ou não encontrado.")
                        
                except Exception as e:
                    print(f"    -> Erro ao tentar marcar checkbox: {e}")

                # 7. Fill Observations
                print("Preenchendo observações...")
                obs_box = page.get_by_role("textbox", name="Observações da Análise")
                if obs_box.is_visible():
                    obs_box.fill(OBS_TEXT)
                    # TAB out to trigger any auto-save logic
                    obs_box.press("Tab") 
                
                # 8. Send / Save Logic
                print("Tentando Enviar...")
                
                # Often "Enviar" is hidden until "Salvar" is clicked, OR it's just off-screen.
                # Ensure we scroll to bottom
                page.mouse.wheel(0, 500)
                page.wait_for_timeout(500)
                
                btn_enviar = page.get_by_role("button", name="Enviar")
                
                if not btn_enviar.is_visible():
                    print("Botão Enviar não visível. Tentando 'Salvar' antes...")
                    btn_salvar = page.get_by_role("button", name="Salvar")
                    if btn_salvar.is_visible():
                        print("Clicando em Salvar...")
                        # Handle potential dialogs (as seen in recording)
                        page.once("dialog", lambda dialog: dialog.dismiss())
                        btn_salvar.click()
                        time.sleep(2) # Short wait for save to process
                        
                        # Check for Enviar again
                        if btn_enviar.is_visible():
                            print("Botão Enviar apareceu! Clicando...")
                            btn_enviar.click()
                            print(">>> CASO ENVIADO! <<<")
                            time.sleep(3)
                        else:
                            print("Botão Enviar AINDA não visível após Salvar.")
                    else:
                        print("Nem Botão Enviar nem Salvar encontrados.")
                else:
                    print("Clicando em Enviar...")
                    btn_enviar.click()
                    print(">>> CASO ENVIADO! <<<")
                    
                    # Post-Submission Wait: Ensure we left the page
                    print("Aguardando confirmação de saída do caso...")
                    try:
                        page.wait_for_selector(".ui-bizagi-form", state="hidden", timeout=15000)
                        print("    -> Formulário fechado. Pronto para o próximo.")
                    except:
                        print("    -> Aviso: Formulário ainda visível ou timeout. O Force Refresh no próximo ciclo resolverá.")
                    
                    time.sleep(2)

            except Exception as e:
                print(f"ERRO CRÍTICO NO CASO {case_id}: {e}")
                # Save screenshot for debugging
                try:
                    import os
                    debug_path = os.path.join(os.getcwd(), "downloads", f"error_{case_id}.png")
                    page.screenshot(path=debug_path)
                    print(f"Screenshot de erro salvo em: {debug_path}")
                except Exception as s_e:
                    print(f"Erro ao salvar screenshot: {s_e}")
                
                # Attempt to return to Home context for next case
                try:
                    page.goto(BASE_URL)
                    time.sleep(2)
                except:
                    pass
                continue
        
        print("\n>>> PROCESSO CONCLUÍDO <<<")
        input("Pressione Enter para sair...")

if __name__ == "__main__":
    run_batch()
