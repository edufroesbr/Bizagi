import sys
from playwright.sync_api import sync_playwright
from bizagi_bot import BizagiBot
import config

def main():
    print("Iniciando Agente BIZAGI...")
    
    with sync_playwright() as p:
        # Initialize the bot
        bot = BizagiBot(p)
        
        # 1. Start Browser and Manual Login
        print("Abrindo navegador. Por favor, faça o login manualmente.")
        bot.launch_browser()
        bot.go_to_login()
        
        # Wait for user to confirm login (Polling mechanism)
        print("Aguardando login do usuário (Verificando se 'Caixa de entrada' aparece)...")
        if bot.wait_for_login():
             print("Login detectado com sucesso!")
        else:
             print("Timeout: Login não detectado em tempo hábil.")
             return
        
        # 2. Start processing loop
        print("Iniciando loop de processamento...")
        bot.process_all_cases()
        
        print("Loop finalizado (sem mais casos).")
        # bot.close()

if __name__ == "__main__":
    main()
