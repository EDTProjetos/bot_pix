# -*- coding: utf-8 -*-

"""
Automação PopSol - Filtragem e Envio para Google Planilhas
- Acessa PopSol, aplica filtro de pagamento (hoje), clica no ícone de ajuda
  para gerar o número e envia para um Apps Script (Google Sheets).
- Versão para CI (GitHub Actions): Chrome headless estável (sem user-data-dir).
"""

import time
import json
import datetime as dt
import requests

# ===== Selenium =====
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys


def log(title: str, text: str) -> None:
    print(f"[{title}] {text}")


# ================== CONFIG ==================
HEADLESS = True                # para rodar no Actions
WAIT_S = 45                    # mais folgado no ambiente remoto
DATE_FMT_BR = "%d/%m/%Y"

# ================== POPSOL (fixos a pedido do usuário) ==================
LOGIN_URL_POPSOL = "https://cliente.popsolenergia.com.br/#/auth/login"
USER_POPSOL = "raphael.barbosa@energiadetodos.com.br"
PWD_POPSOL  = "Kon@rulind0."

# ================== GOOGLE SHEETS (fixo a pedido do usuário) ==================
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyE8JPkff699pufz7js4xr-XWa5G5pc_6be-RmG5CQbpWHQlq6loYIceU6XB9oYcWRxfg/exec"


# ================== WebDriver ==================
def make_driver(headless: bool = True) -> webdriver.Chrome:
    """Chrome headless estável para CI (sem user-data-dir)."""
    opts = Options()
    opts.page_load_strategy = "eager"
    opts.add_argument("--window-size=1366,900")
    opts.add_argument("--lang=pt-BR")
    opts.add_argument("--disable-notifications")

    # Flags essenciais no GitHub Actions (Linux)
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--no-zygote")
    opts.add_argument("--disable-setuid-sandbox")
    opts.add_argument("--remote-debugging-pipe")  # evita conflito de porta

    if headless:
        opts.add_argument("--headless=new")

    # ⚠️ Importante: NÃO usar --user-data-dir aqui
    return webdriver.Chrome(options=opts)


def safe_click_we(driver, we):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", we)
    try:
        we.click()
    except Exception:
        driver.execute_script("arguments[0].click();", we)


def safe_click(driver, locator):
    el = WebDriverWait(driver, WAIT_S).until(EC.element_to_be_clickable(locator))
    safe_click_we(driver, el)


# ================== Google Sheets ==================
def send_to_google_sheet(numero_str: str):
    """Envia o número para o Apps Script (que grava o timestamp)."""
    if not APPS_SCRIPT_URL:
        log("AVISO", "URL do Apps Script não configurada. Envio ignorado.")
        return

    log("ENVIO", f"Enviando '{numero_str}' para a Google Planilha…")
    headers = {"Content-Type": "application/json"}
    payload = {"text": numero_str}

    try:
        resp = requests.post(APPS_SCRIPT_URL, data=json.dumps(payload), headers=headers, timeout=20)
        resp.raise_for_status()
        data = resp.json()
        if data.get("ok"):
            log("SUCESSO", f"Número salvo (linha {data.get('row')}).")
        else:
            raise RuntimeError(f"Apps Script retornou erro: {data.get('error')}")
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Falha HTTP ao enviar para a planilha: {e}")


# ================== Fluxo Web ==================
def pops_logar_e_filtrar_recebimentos(driver: webdriver.Chrome, data_de: str, data_ate: str) -> None:
    w = WebDriverWait(driver, WAIT_S)

    log("LOGIN", "Acessando a página de login do PopSol…")
    driver.get(LOGIN_URL_POPSOL)

    # Login
    email = w.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']")))
    senha = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
    email.clear(); email.send_keys(USER_POPSOL)
    senha.clear(); senha.send_keys(PWD_POPSOL)
    log("LOGIN", "Enviando credenciais…")
    senha.send_keys(Keys.ENTER)
    w.until(lambda d: "auth/login" not in d.current_url)
    log("LOGIN", "Login realizado.")

    # Recebimentos (tenta seletores alternativos)
    log("NAV", "Abrindo 'Recebimentos'…")
    try:
        safe_click(driver, (By.XPATH, "//a[normalize-space()='Recebimentos']"))
    except TimeoutException:
        try:
            safe_click(driver, (By.XPATH, "//label[normalize-space()='Recebimentos']"))
        except TimeoutException:
            safe_click(driver, (By.XPATH, "//span[normalize-space()='Recebimentos']"))

    # Filtro Data de pagamento
    log("FILTRO", "Aguardando campos de 'Data de pagamento'…")
    campos_data_xpath = "(//h3[normalize-space()='Data de pagamento']/following::input)[position()<=2]"
    w.until(EC.presence_of_element_located((By.XPATH, campos_data_xpath)))
    inputs = driver.find_elements(By.XPATH, campos_data_xpath)
    if len(inputs) < 2:
        raise RuntimeError("Campos de data de pagamento não encontrados.")

    log("FILTRO", f"Aplicando período: {data_de} até {data_ate}")
    js_set_value = """
      const el = arguments[0], val = arguments[1];
      el.focus(); el.value = val;
      el.dispatchEvent(new Event('input', {bubbles:true}));
      el.dispatchEvent(new Event('change', {bubbles:true}));
      el.dispatchEvent(new Event('blur', {bubbles:true}));
    """
    driver.execute_script(js_set_value, inputs[0], data_de)
    time.sleep(0.5)
    driver.execute_script(js_set_value, inputs[1], data_ate)

    log("FILTRO", "Clicando em 'Aplicar filtros'…")
    safe_click(driver, (By.XPATH, "//button[.//label[normalize-space()='Aplicar filtros']]"))
    time.sleep(3)

    # Ícone question_mark e captura do número
    log("EXTRAÇÃO", "Clicando no ícone 'question_mark'…")
    safe_click(driver, (By.XPATH, "//mat-icon[normalize-space()='question_mark']"))

    try:
        locator_numero = (By.XPATH, "//div[@title='Contar novamente']")
        log("EXTRAÇÃO", "Aguardando o número…")
        elemento_numero = WebDriverWait(driver, WAIT_S).until(EC.visibility_of_element_located(locator_numero))
        numero_copiado = elemento_numero.text.strip()

        if not numero_copiado.isdigit():
            raise ValueError(f"Texto extraído não numérico: '{numero_copiado}'")

        log("EXTRAÇÃO", f"Número extraído: {numero_copiado}")
        send_to_google_sheet(numero_copiado)

    except TimeoutException:
        raise RuntimeError("Não encontrei o elemento com o número (title='Contar novamente').")


# ================== MAIN ==================
def main():
    hoje = dt.date.today()
    data_inicio_str = hoje.strftime(DATE_FMT_BR)
    data_fim_str = hoje.strftime(DATE_FMT_BR)

    print("=" * 60)
    print("Iniciando automação PopSol -> Google Planilhas")
    print(f"[Data Alvo] Período de pagamento: {data_inicio_str}")
    print("=" * 60)

    try:
        driver = make_driver(headless=HEADLESS)
        pops_logar_e_filtrar_recebimentos(driver, data_inicio_str, data_fim_str)
        print("\nProcesso concluído com sucesso!\n- O número e o timestamp foram enviados para a Google Planilha.")
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
