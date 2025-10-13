# -*- coding: utf-8 -*-

"""
Automação PopSol - Filtragem e Envio para Google Planilhas
- Objetivo: Acessar o portal PopSol, aplicar filtro de pagamento, clicar no ícone
de ajuda para gerar um número e enviá-lo para uma Google Planilha via Web App.
- Versão com login otimizado e seletor de número preciso.
- Ajustado para funcionar com o Apps Script que adiciona o timestamp.
"""

import time
import datetime as dt
import tempfile, subprocess
import requests # Para requisições HTTP
import json     # Para formatar os dados

# ===== Imports do Selenium =====
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys

# ===== Notificação final (Opcional, mas útil) =====
import ctypes
try:
    import tkinter as tk
    from tkinter import messagebox
except Exception:
    tk = None; messagebox = None

def _vbs_escape(s: str) -> str:
    return (s or "").replace('"', '""')

def _notify_vbs(title: str, text: str) -> bool:
    try:
        vbs = f'MsgBox "{_vbs_escape(text)}", 64, "{_vbs_escape(title)}"'
        with tempfile.NamedTemporaryFile(delete=False, suffix=".vbs", mode="w", encoding="cp1252") as f:
            f.write(vbs); vbs_path = f.name
        subprocess.Popen(["wscript.exe", vbs_path], close_fds=True)
        return True
    except Exception:
        return False

def notify(title: str, text: str) -> None:
    try:
        ctypes.windll.user32.MessageBoxW(0, text, title, 0x40); return
    except Exception: pass
    if _notify_vbs(title, text): return
    if tk and messagebox:
        try:
            root = tk.Tk(); root.withdraw(); messagebox.showinfo(title, text); root.destroy(); return
        except Exception: pass
    print(f"[NOTIFICAÇÃO] {title}\n{text}")

# ================== CONFIG ==================
# Mude para False para ver o navegador em ação
HEADLESS = False
# Tempo máximo de espera para os elementos aparecerem na tela
WAIT_S = 30
# Formato de data usado pelo portal PopSol
DATE_FMT_BR = "%d/%m/%Y"

# ================== POPSOL - CREDENCIAIS E URL ==================
LOGIN_URL_POPSOL = "https://cliente.popsolenergia.com.br/#/auth/login"
USER_POPSOL = "raphael.barbosa@energiadetodos.com.br"
PWD_POPSOL = "Kon@rulind0."

# ================== GOOGLE SHEETS - WEB APP ==================
# ⚠️ COLE A URL DO SEU APP DA WEB PUBLICADO AQUI ⚠️
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbws_qHwe4PfwzMt59RIz98pa3OWBvPvJNgyZphS9Mmgr3Hz-9ZIqp_0yvVi6Fm3Ge6emw/exec"

# ================== Helpers de WebDriver ==================
def make_driver(headless: bool = True) -> webdriver.Chrome:
    """Cria e configura uma instância do Chrome WebDriver."""
    opts = Options()
    opts.page_load_strategy = "eager"
    opts.add_argument("--window-size=1366,900")
    if headless:
        opts.add_argument("--headless=new")
    return webdriver.Chrome(options=opts)

def safe_click_we(driver, we):
    """Clica em um elemento de forma segura, rolando até ele primeiro."""
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", we)
    try:
        we.click()
    except Exception:
        driver.execute_script("arguments[0].click();", we)

def safe_click(driver, locator):
    """Espera um elemento ficar clicável e usa o safe_click_we para clicar."""
    el = WebDriverWait(driver, WAIT_S).until(EC.element_to_be_clickable(locator))
    safe_click_we(driver, el)

# ================== Função de Envio para Google Planilhas ==================
def send_to_google_sheet(numero_str: str):
    """
    Envia o número extraído para o Web App do Google Apps Script.
    O Apps Script irá automaticamente adicionar a data e hora da execução.
    """
    if not APPS_SCRIPT_URL or "SUA_URL" in APPS_SCRIPT_URL:
        print("[AVISO] URL do Google Apps Script não configurada. O envio foi ignorado.")
        return

    # MENSAGEM AJUSTADA: Informa que o Apps Script registrará o timestamp
    print(f"Enviando o número '{numero_str}' (com registro de hora/data pelo Apps Script) para a Google Planilha...")
    headers = {"Content-Type": "application/json"}
    # O payload envia APENAS o número, conforme o Apps Script espera
    payload = {"text": numero_str} 

    try:
        response = requests.post(APPS_SCRIPT_URL, data=json.dumps(payload), headers=headers, timeout=15)
        response.raise_for_status() # Lança um erro se o status for 4xx ou 5xx
        
        response_data = response.json()
        if response_data.get("ok"):
            print(f"Sucesso! Número enviado e salvo na linha {response_data.get('row')}.")
        else:
            raise RuntimeError(f"O script da planilha retornou um erro: {response_data.get('error')}")

    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Falha na comunicação com a Google Planilha: {e}")


# ================== Fluxo web principal ==================
def pops_logar_e_filtrar_recebimentos(driver: webdriver.Chrome, data_de: str, data_ate: str) -> None:
    """
    Realiza o login, filtra, extrai o número e o envia para a planilha.
    """
    w = WebDriverWait(driver, WAIT_S)
    print("Acessando a página de login do PopSol...")
    driver.get(LOGIN_URL_POPSOL)

    # --- Etapa de Login ---
    email = w.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']")))
    senha = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
    email.clear(); email.send_keys(USER_POPSOL)
    senha.clear(); senha.send_keys(PWD_POPSOL)
    print("Realizando login (pressionando Enter)...")
    senha.send_keys(Keys.ENTER)
    w.until(lambda d: "auth/login" not in d.current_url)
    print("Login realizado com sucesso.")

    # --- Etapa de Navegação para Recebimentos ---
    print("Navegando para a tela de 'Recebimentos'...")
    safe_click(driver, (By.XPATH, "//a[normalize-space()='Recebimentos'] | //label[normalize-space()='Recebimentos'] | //span[normalize-space()='Recebimentos']"))

    # --- Etapa de Filtragem ---
    print("Aguardando os campos de filtro de data de pagamento...")
    campos_data_xpath = "(//h3[normalize-space()='Data de pagamento']/following::input)[position()<=2]"
    w.until(EC.presence_of_element_located((By.XPATH, campos_data_xpath)))
    inputs = driver.find_elements(By.XPATH, campos_data_xpath)
    if len(inputs) < 2:
        raise RuntimeError("PopSol: Campos de data de pagamento não encontrados.")
    
    print(f"Aplicando filtro de data: de '{data_de}' até '{data_ate}'...")
    js_set_value = """
    const el = arguments[0], val = arguments[1]; el.focus(); el.value = val;
    el.dispatchEvent(new Event('input', {bubbles:true})); el.dispatchEvent(new Event('change', {bubbles:true})); el.dispatchEvent(new Event('blur', {bubbles:true}));
    """
    driver.execute_script(js_set_value, inputs[0], data_de)
    time.sleep(0.5)
    driver.execute_script(js_set_value, inputs[1], data_ate)
    print("Clicando em 'Aplicar filtros'...")
    safe_click(driver, (By.XPATH, "//button[.//label[normalize-space()='Aplicar filtros']]"))
    time.sleep(3) # Espera a tabela de filtros recarregar

    # --- FLUXO DE EXTRAÇÃO: Clicar, Copiar e Enviar ---
    print("Clicando no ícone 'question_mark'...")
    safe_click(driver, (By.XPATH, "//mat-icon[normalize-space()='question_mark']"))

    try:
        # AJUSTE: Localizador preciso para o elemento que contém o número, baseado no seu HTML.
        locator_numero = (By.XPATH, "//div[@title='Contar novamente']")
        
        print("Aguardando o número ser gerado...")
        elemento_numero = w.until(EC.visibility_of_element_located(locator_numero))
        
        numero_copiado = elemento_numero.text.strip()
        if not numero_copiado.isdigit(): # Garante que o texto extraído é um número
            raise ValueError(f"O texto extraído ('{numero_copiado}') não parece ser um número válido.")

        print(f"Número extraído com sucesso: {numero_copiado}")
        
        # Envia para a planilha
        send_to_google_sheet(numero_copiado)

    except TimeoutException:
        raise RuntimeError("Não foi possível encontrar o elemento com o número ('title=Contar novamente'). Verifique se a página se comportou como esperado.")
    except Exception as e:
        # Repassa outros erros, como falha no envio ou texto inválido
        raise e


# ================== MAIN ==================
def main():
    hoje = dt.date.today()
    #hoje = dt.date(2025, 10, 10) # Linha para teste com data fixa
    data_inicio_str = hoje.strftime(DATE_FMT_BR)
    data_fim_str = hoje.strftime(DATE_FMT_BR)
    
    print("="*50)
    print("Iniciando automação PopSol -> Google Planilhas")
    print(f"[Data Alvo] Período de pagamento: {data_inicio_str}")
    print("="*50)

    success = False
    final_msg = "A automação encontrou um erro."
    driver = make_driver(headless=HEADLESS)
    
    try:
        pops_logar_e_filtrar_recebimentos(driver, data_inicio_str, data_fim_str)
        
        # MENSAGEM AJUSTADA: Informa que o timestamp também foi registrado
        final_msg = (
            "Processo concluído com sucesso!\n\n"
            f"- Portal: PopSol Energia\n"
            f"- Filtro de Pagamento: {data_inicio_str}\n"
            f"- O número e o registro de data/hora foram enviados para a sua Google Planilha."
        )
        success = True
        print("\n" + final_msg)

    except Exception as e:
        final_msg = f"Ocorreu um erro durante a execução:\n\n{e}"
        print(f"\n[ERRO] {final_msg}")

    finally:
        try:
            driver.quit()
        except Exception:
            pass
        
        titulo_notificacao = "Automação Concluída" if success else "Erro na Automação"
        notify(titulo_notificacao, final_msg)

if __name__ == "__main__":
    main()
