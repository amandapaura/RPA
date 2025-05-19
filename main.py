import os
from pbi_web_scraping import PowerBIScraper
from pbi_workdpaces_reports import PowerBIReportCollector


if __name__ == '__main__':
    base_path = 'C:/user/amandapaura/backup/PowerBi' #example 
    json_path = os.path.join(base_path, 'workspaces_relatorios.json')
    chrome_driver_path = "C:/chromedriver/chromedriver.exe"
    user_email = "email@email.com"

    # Passo 1: Gerar JSON com os relatórios a partir da API
    collector = PowerBIReportCollector(user_email)
    collector.run()
    
    # Passo 2: Executar o scraping no Power BI Web
    scraper = PowerBIScraper(
        user=user_email,
        json_path=json_path,
        chrome_driver_path=chrome_driver_path
    )
    scraper.run()
