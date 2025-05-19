import os
import json
import traceback
from tqdm import tqdm
from api_powerbi import PowerBIExporter  # classe anterior já criada 

class PowerBIReportCollector(PowerBIExporter):
    def __init__(self, user_email, base_path='C:/user/amandapaura/backup/PowerBi', output_file='workspaces_reports.json'):
        super().__init__(kr, base_path)
        self.output_file = output_file
        self.report_links = {}

    def collect_report_urls(self):
        try:
            workspaces = self.get_workspaces().get('value', [])

            for ws in tqdm(workspaces, desc='Coletando Workspaces'):
                workspace_name = ws['name']
                workspace_id = ws['id']
                reports = self.get_reports(workspace_id).get('value', [])

                self.report_links[workspace_id] = {
                    "ws_name": workspace_name,
                    "relatorios": {}
                }

                for report in reports:
                    report_id = report['id']
                    url = report['webUrl']
                    self.report_links[workspace_id]['relatorios'][report_id] = url

        except BaseException as erro:
            print("Erro na coleta de relatórios.")
            print(traceback.format_exc())

    def save_to_json(self):
        with open(self.output_file, 'w', encoding='utf-8') as f:
            json.dump(self.report_links, f, indent=4, ensure_ascii=False)
        print(f"JSON com relatórios salvo em: {self.output_file}")

    def run(self):
        self.collect_report_urls()
        self.save_to_json()
