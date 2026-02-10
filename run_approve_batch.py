import time
from playwright.sync_api import sync_playwright

# CONFIGURAÇÃO
BASE_URL = "https://digital-ons.bizagi.com/"

# Default list from User Request
DEFAULT_CASES = [
    "46765", "46767", "46770", "46773", "46775",
    "46777", "46780", "46785", "46788", "46791"
]

def run_approval_batch():
    # Ask for cases at runtime
    # print(">>> CONFIGURAÇÃO DO LOTE DE APROVAÇÃO <<<")
    # case_input = input("Cole a lista de casos separados por vírgula (ou pressione Enter para usar o padrão): ")
    # if not case_input.strip():
    current_list = DEFAULT_CASES
    print(f"Usando lista padrão de teste: {current_list}")
  
    # else:
    #     current_list = [c.strip() for c in case_input.split(",") if c.strip()]
    #     print(f"Lista carregada com {len(current_list)} casos.")

    print("MODO DE PRODUÇÃO: Executando aprovação em lote.")

    with sync_playwright() as p:
        # VS Code / Window Size match
        browser = p.chromium.launch(headless=False, args=["--start-maximized"])
        context = browser.new_context(viewport=None) 
        page = context.new_page()

        print("\n>>> INICIANDO APROVAÇÃO AUTOMÁTICA (Engine V3) <<<")
        
        page.goto(BASE_URL)
        
        # ---------------------------------------------------------
        # 0. ROBUST AUTO-LOGIN (Recorded Flow Strategy)
        # ---------------------------------------------------------
        print("Verificando necessidade de login (Estratégia Input + Enter)...")
        try:
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
                      
                  print("Aguardando carregamento final do Dashboard...")
                  page.wait_for_load_state("networkidle")
                  
        except Exception as e:
             print(f"ERRO DE LOGIN OU TIMEOUT: {e}")
             print("Tentando seguir...")

        # Robust Login Wait
        try:
            page.wait_for_selector("#menuQuery, input[placeholder='Pesquisar'], a:has-text('Caixa de entrada')", timeout=60000)
            print("Login confirmado!")
        except:
            print("Login não detectado a tempo.")
            return
            
        # ---------------------------------------------------------
        # MAIN LOOP
        # ---------------------------------------------------------
        for case_id in current_list:
            print(f"\n==========================================")
            print(f"PROCESSANDO APROVAÇÃO: {case_id}")
            print(f"==========================================")
            
            try:
                # 1. HARD RESET Navigation
                print("Navegando para o Início (Reset)...")
                try:
                    page.goto(BASE_URL)
                    page.wait_for_load_state("networkidle")
                except:
                    pass

                # 2. Search Strategy (V3)
                print(f"Pesquisando {case_id}...")
                search_success = False
                for s_attempt in range(3):
                    if page.locator(f"tr:has-text('{case_id}')").is_visible():
                         search_success = True
                         break
                    
                    search_box = page.get_by_role("textbox", name="Pesquisar")
                    if not search_box.is_visible():
                        btn_search = page.locator(".bz-icon-search, button.bz-search-button").first
                        if btn_search.is_visible():
                             btn_search.click()
                        page.wait_for_timeout(500)
                    
                    search_box.click()
                    search_box.fill("")
                    search_box.type(case_id, delay=100)
                    page.wait_for_timeout(500)
                    
                    btn_trigger = page.locator(".bz-search-button, button.bz-icon-search").first
                    if btn_trigger.is_visible():
                        btn_trigger.click()
                    else:
                        search_box.press("Enter")
                    
                    try:
                        page.wait_for_selector(f"tr:has-text('{case_id}')", timeout=5000)
                        search_success = True
                        break
                    except:
                        print(f"Retentando pesquisa ({s_attempt+1}/3)...")
                        page.wait_for_timeout(1000)
                
                if not search_success and not page.locator(f"tr:has-text('{case_id}')").is_visible():
                     print("AVISO: Caso não encontrado na pesquisa.")

                # 3. Open Case
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

                # 4. Wait & Scroll
                page.wait_for_timeout(3000)
                page.mouse.wheel(0, 500)
                
                # 5. Tomar Posse
                take_ownership_btn = page.get_by_role("link", name="Tomar posse")
                if take_ownership_btn.is_visible(timeout=5000):
                    print("Tomando posse...")
                    take_ownership_btn.click()
                    page.wait_for_load_state("networkidle")
                    time.sleep(2)

                # ---------------------------------------------------------
                # 6. APPROVAL ACTION (Modified Step)
                # ---------------------------------------------------------
                print("Selecionando Ação: Aprovar...")
                
                # Try locating the dropdown caret or select menu
                # Recording showed: .biz-btn-caret OR .ui-selectmenu-btn
                dropdown_clicked = False
                try:
                    caret = page.locator(".biz-btn-caret, .ui-selectmenu-btn, .ui-select-toggle").first
                    if caret.is_visible():
                        caret.click()
                        dropdown_clicked = True
                    else:
                        print("Menu de ação não encontrado pelo seletor padrão.")
                except Exception as e:
                    print(f"Erro ao clicar no menu: {e}")

                if dropdown_clicked:
                    # Wait for option "Aprovar"
                    try:
                        option_aprovar = page.get_by_role("option", name="Aprovar")
                        if option_aprovar.is_visible():
                            option_aprovar.click()
                            print("Opção 'Aprovar' selecionada.")
                            page.wait_for_timeout(1000)
                        else:
                            print("Opção 'Aprovar' não visível no menu.")
                    except:
                        print("Erro ao selecionar opção Aprovar.")
                
                # ---------------------------------------------------------
                # 7. SEND (No Observations/Checkbox)
                # ---------------------------------------------------------
                print("Enviando solicitação...")
                btn_enviar = page.get_by_role("button", name="Enviar")
                
                if btn_enviar.is_visible():
                    btn_enviar.click()
                    print(">>> CASO ENVIADO (APROVADO)! <<<")
                    
                    # Wait for exit
                    try:
                        page.wait_for_selector(".ui-bizagi-form", state="hidden", timeout=15000)
                    except:
                        print("Aviso: Form ainda visível.")
                    time.sleep(2)
                else:
                    print("Botão Enviar não encontrado.")

            except Exception as e:
                print(f"ERRO CRÍTICO NO CASO {case_id}: {e}")
                try:
                    page.screenshot(path=f"downloads/erro_aprovacao_{case_id}.png")
                except: pass
                continue
        
        print("\n>>> PROCESSO DE APROVAÇÃO CONCLUÍDO <<<")
        # input("Pressione Enter para sair...")

if __name__ == "__main__":
    run_approval_batch()
