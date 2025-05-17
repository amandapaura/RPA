import json
import os
import keyring as kr
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


#-- ------------ Initials configurations -------------------------------------------------------------------------------------------------------------------    

file = open(os.path.join(base_path,'workspaces_relatorios.json'), 'r', encoding='utf-8')
dados_ws = json.load(file)
chrome_drive_path = "C:\chromedriver\chromedriver.exe"

#--- user and password power bi ----

user     = 'email@email.com' ## change here your email
password = 'xxxx' ##insert your password

kr.set_password("powerbi",user,password) ## By using keyring, you only need to set up your password once. After that, your login credentials are securely stored and you don’t need to expose them again

password = kr.get_password('powerbi','email@email.com')  ## from now on you just need to get your password

#------------ functions -----------

def abrir_chrome_webdriver():
    # Configure Selenium WebDriver
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # remove se quiser ver o navegador
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")

    # Set the path to your ChromeDriver.
    return webdriver.Chrome()

def verificar_e_fechar_mensagem_erro(driver):
    try:
        # Wait up to 5 seconds for the error message to appear
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'Relatório inválido')]"))
        )
        
        # Click 'OK' if an error message appears
        botao_ok = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "okButton"))
        )
        botao_ok.click()
        print("Mensagem de erro detectada e botão OK clicado.")

    except Exception as e:
        # If not, the process continues..
        print("Nenhuma mensagem de erro detectada. Seguindo o fluxo.")

def fechar_abas_internas(driver):
    try:
        # Find all close buttons using the data-testid attribute
        botoes_fechar = driver.find_elements(By.CSS_SELECTOR, 'button[data-testid="tab-close"]')
        
        print(f"Encontrados {len(botoes_fechar)} abas abertas.")
        
        for botao in botoes_fechar:
            try:
                driver.execute_script("arguments[0].click();", botao)
                time.sleep(0.5)
            except Exception as e:
                print(f"Erro ao fechar aba: {e}")
                
    except Exception as e:
        print(f"Erro geral ao localizar abas: {e}")



# ----------------------------------------------
# ------------- EXECUTE SELENIUM -------------
# ------------------------------------------------


driver = webdriver.Chrome()
time.sleep(5)  ##sempre bom esperar um tempo antes de cada etapa

try:
    driver.get("https://app.powerbi.com/")

    time.sleep(3) 

    # Wait for the login to appear and enter the email
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "email")))
    driver.find_element(By.ID, "email").send_keys(user)
    driver.find_element(By.ID, "email").send_keys(Keys.RETURN)

    time.sleep(2)  # you can adjust if you need

    # enter the password
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.NAME, "passwd")))
    driver.find_element(By.NAME, "passwd").send_keys(password)
    driver.find_element(By.NAME, "passwd").send_keys(Keys.RETURN)

    # Confirm it if needed
    time.sleep(2)
    try:
        driver.find_element(By.ID, "idBtn_Back").click()
    except:
        pass
    
    time.sleep(2)

    for ws,rel in dados_ws.items():
        for report, url in rel['relatorios'].items():
            driver.get(url)

            time.sleep(4)

            # Click to export the PBIX 
            # Click at the button 'More Options'
            try:

                # wait the 'File' button
                print("Clicking at 'File'...")
                file_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="appbar-file-menu-btn"]'))
                )
                file_button.click()

                # wait and click the download option
                print("Clicking at 'Download'...")
                download_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, '//span[contains(text(), "Download") or contains(text(), "Baixar este arquivo")]'))
                )
                download_button.click()

                verificar_e_fechar_mensagem_erro(driver)

                time.sleep(2)

                # Wait for the 'Download' button in the window/modal and click it
                print("Clicking at the buttom 'Download' do popup...")
                confirm_download_button = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.ID, "okButton"))
                )
                confirm_download_button.click()

                time.sleep(3)

                verificar_e_fechar_mensagem_erro(driver)
                
                fechar_abas_internas(driver)

                # check if it's zero
                if len(driver.find_elements(By.CSS_SELECTOR, 'button[data-testid="tab-close"]')) == 0:
                    print("Tudo limpo, pode seguir.")
                else:
                    print("Ainda ficou alguma aba aberta, revisar.")


                print("Download started successfully!")

            except Exception as e:
                print("Error clicking the buttons:", e)
                
finally:
    driver.quit()

print('Downloads successful.')

