import os
import json
import time
import keyring as kr
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class PowerBIScraper:
    def __init__(self, user, json_path, chrome_driver_path):
        self.user = user
        self.password = kr.get_password('powerbi', user)
        self.json_path = json_path
        self.chrome_driver_path = chrome_driver_path
        self.driver = None
        self.reports_data = self.load_reports()

    def load_reports(self):
        with open(self.json_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def save_reports(self, filename='powerbi_download_status.json'):
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.reports_data, f, indent=4, ensure_ascii=False)

    def setup_driver(self):
        options = Options()
        options.add_argument("--headless")  # Remova se quiser ver o navegador
        options.add_argument("--window-size=1920,1080")
        self.driver = webdriver.Chrome(executable_path=self.chrome_driver_path, options=options)

    def login(self):
        self.driver.get("https://app.powerbi.com/")
        WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, "email"))).send_keys(self.user)
        self.driver.find_element(By.ID, "email").send_keys(Keys.RETURN)

        WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.NAME, "passwd"))).send_keys(self.password)
        self.driver.find_element(By.NAME, "passwd").send_keys(Keys.RETURN)

        # Tentar clicar em "Não" na pergunta de manter login
        time.sleep(2)
        try:
            self.driver.find_element(By.ID, "idBtn_Back").click()
        except:
            pass

    def acessar_e_baixar_relatorios(self):
        for ws_id, ws_info in self.reports_data.items():
            for report_id, url in ws_info['relatorios'].items():
                print(f"Acessando: {url}")
                try:
                    self.driver.get(url)
                    time.sleep(4)
                    self.baixar_relatorio()
                    ws_info['relatorios'][report_id] = {'url': url, 'status': 'downloaded'}
                except Exception as e:
                    print(f"Erro ao processar {url}: {e}")
                    ws_info['relatorios'][report_id] = {'url': url, 'status': 'error', 'error_msg': str(e)}

    def baixar_relatorio(self):
        self.click_file_menu()
        self.click_download_option()
        self.verificar_e_fechar_mensagem_erro()
        self.confirm_download_popup()
        self.verificar_e_fechar_mensagem_erro()
        self.fechar_abas_internas()

    def click_file_menu(self):
        file_button = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="appbar-file-menu-btn"]'))
        )
        file_button.click()

    def click_download_option(self):
        download_button = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Download") or contains(text(), "Baixar este arquivo")]'))
        )
        download_button.click()

    def confirm_download_popup(self):
        confirm_button = WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((By.ID, "okButton"))
        )
        confirm_button.click()
        time.sleep(2)

    def verificar_e_fechar_mensagem_erro(self):
        try:
            WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Relatório inválido')]"))
            )
            botao_ok = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.ID, "okButton"))
            )
            botao_ok.click()
            print("Mensagem de erro detectada e botão OK clicado.")
        except:
            print("Nenhuma mensagem de erro detectada.")

    def fechar_abas_internas(self):
        botoes_fechar = self.driver.find_elements(By.CSS_SELECTOR, 'button[data-testid="tab-close"]')
        print(f"Encontrados {len(botoes_fechar)} abas abertas.")
        for botao in botoes_fechar:
            try:
                self.driver.execute_script("arguments[0].click();", botao)
                time.sleep(0.5)
            except Exception as e:
                print(f"Erro ao fechar aba: {e}")

    def run(self):
        try:
            self.setup_driver()
            self.login()
            self.acessar_e_baixar_relatorios()
            self.save_reports()
        except Exception as e:
            print(f"Erro na execução: {e}")
        finally:
            if self.driver:
                self.driver.quit()
            print("Execução finalizada.")

