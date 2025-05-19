import os
import time
import datetime
import requests
import msal
import keyring as kr

class PowerBIExporter:
    def __init__(self, user_email, base_path='C:/user/amandapaura/backup/PowerBi'):
        sefl.user = user_email
        self.base_path = base_path
        self.tenant_id = kr.get_password('PowerBI', 'tenant_id')
        self.client_id = kr.get_password('PowerBI', 'client_id')
        self.client_secret = kr.get_password('PowerBI', 'client_secret')
        self.authority = f'https://login.microsoftonline.com/{self.tenant_id}'
        self.scope = ['https://analysis.windows.net/powerbi/api/.default']
        self.token = self.get_token()

    def get_token(self):
        app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            authority=self.authority,
            client_credential=self.client_secret
        )

        password = kr.get_password('PowerBI', self.user)
        result = app.acquire_token_by_username_password(self.user, password, scopes=self.scope)

        if "access_token" in result:
            return result['access_token']
        else:
            raise Exception("Erro ao adquirir token de acesso.")

    def get_workspaces(self):
        url = 'https://api.powerbi.com/v1.0/myorg/groups'
        headers = {'Authorization': f'Bearer {self.token}'}
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()

    def get_reports(self, group_id):
        url = f'https://api.powerbi.com/v1.0/myorg/groups/{group_id}/reports'
        headers = {'Authorization': f'Bearer {self.token}'}
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()

    def get_backup_path(self, workspace_name):
        today = datetime.date.today()

        if today.day == 1:
            month_folder = today.strftime("%Y-%m")
            path = os.path.join(self.base_path, 'month', month_folder, workspace_name)
        else:
            week_number = today.strftime("%Y-%W")
            day_name = today.strftime("%A").lower()
            path = os.path.join(self.base_path, 'week', week_number, day_name, workspace_name)

        os.makedirs(path, exist_ok=True)
        return path

    def is_exportable_report(self, group_id, report_id):
        headers = {'Authorization': f'Bearer {self.token}'}

        report_url = f'https://api.powerbi.com/v1.0/myorg/groups/{group_id}/reports/{report_id}'
        report_response = requests.get(report_url, headers=headers)
        report = report_response.json()

        dataset_id = report.get('datasetId')
        if not dataset_id:
            return False, "Reports without dataset - PBIX cannot be exported"

        dataset_url = f'https://api.powerbi.com/v1.0/myorg/groups/{group_id}/datasets/{dataset_id}'
        dataset_response = requests.get(dataset_url, headers=headers)
        dataset = dataset_response.json()

        if dataset.get('datasetMode') != 'Import':
            return False, f"Dataset in {dataset.get('datasetMode')} mode - PBIX can't be exported"

        return True, "OK"

    def export_report(self, group_id, report_id, report_name, output_dir):
        exportable, reason = self.is_exportable_report(group_id, report_id)

        if not exportable:
            msg = f"Relatório {report_name} não pode ser exportado: {reason}"
            return {'status': 'skipped', 'message': msg}

        url = f'https://api.powerbi.com/v1.0/myorg/reports/{report_id}/Export'
        headers = {'Authorization': f'Bearer {self.token}', 'Content-Type': 'application/json'}
        body = {"format": "PBIX"}

        response = requests.post(url, headers=headers, json=body)

        if response.status_code == 202:
            export_url = response.headers.get('Location')
            print(f'Exportando {report_name}')
            
            while True:
                r = requests.get(export_url, headers=headers, stream=True)
                if r.status_code == 200:
                    file_path = os.path.join(output_dir, f'{report_name}.pbix')
                    with open(file_path, 'wb') as f:
                        for chunk in r.iter_content(chunk_size=1024 * 1024):
                            f.write(chunk)
                    print(f'{report_name}.pbix salvo com sucesso')
                    break
                elif r.status_code == 202:
                    time.sleep(2)
                else:
                    print(f"Erro ao exportar {report_name}: {r.status_code} - {r.text}")
                    break
        else:
            print(f"Erro ao iniciar exportação: {response.status_code} - {response.text}")

