import os
from pbi_web_scraping import PowerBIScraper
from pbi_workspaces_reports import PowerBIReportCollector

if __name__ == '__main__':
    base_path = 'C:/user/amandapaura/backup/PowerBi' #example 
    json_path = os.path.join(base_path, 'workspaces_reports.json')
    chrome_driver_path = "C:/chromedriver/chromedriver.exe"
    user_email = "email@email.com"
    
    print("Iniciando a coleta de URLS")
    # Gerar JSON com os relatórios a partir da API
    collector = PowerBIReportCollector(user_email)
    collector.run()

    print('Iniciando Scrapping')
    # Executar o scraping no Power BI Web
    scraper = PowerBIScraper(
        user=user_email,
        json_path=json_path,
        chrome_driver_path=chrome_driver_path
    )
    scraper.run()

