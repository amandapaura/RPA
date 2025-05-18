#-- ---------------------------------------------------------------------------------------
import requests
import json
import pbirest
import msal
from extensao import *;
from tqdm import tqdm
import os
import datetime
import keyring as kr


tenant_id     = kr.get_password('PowerBI', 'tenant_id')
client_id     = kr.get_password('PowerBI', 'client_id')
client_secret = kr.get_password('PowerBI', 'client_secret')

base_path = 'C:/user/amandapaura/backup/PowerBi'
authority = f'https://login.microsoftonline.com/{tenant_id}'
scope = ['https://analysis.windows.net/powerbi/api/.default']

# ---------------- Funções ----------------------------

# getting token
def get_token(): 
    app = msal.ConfidentialClientApplication(client_id=client_id, 
                                             authority=authority, 
                                             client_credential=client_secret)
    #result = app.acquire_token_for_client(scopes=scope)

    user     = kr.get_password('PowerBI', 'user')
    password = kr.get_password('PowerBI', 'password')
    result   = app.acquire_token_by_username_password(user,password,scopes=scope)

    if "access_token" in result:
        return result['access_token']


# listing workspaces
def get_workspaces(token):    
    url = 'https://api.powerbi.com/v1.0/myorg/groups'
    headers = {'Authorization': f'Bearer {token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    #print(response.json())
    print('\n\n')
    return response.json()

# listing reports
def get_reports(token, group_id):    
    url = f'https://api.powerbi.com/v1.0/myorg/groups/{group_id}/reports'
    headers = {'Authorization': f'Bearer {token}'}
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    #print(response.json())
    print('\n\n')

    return response.json()

## creating path for saving monthly backups (every 1st day) and daily (organized by week)
def get_backup_path(workspace_name):
    today = datetime.date.today()

    if today.day == 1:
        # Backup mensal
        month_folder = today.strftime("%Y-%m")
        path = os.path.join(base_path, 'month', month_folder, workspace_name)
    else:
        # Backup semanal
        week_number = today.strftime("%Y-%W")
        day_name = today.strftime("%A").lower()
        path = os.path.join(base_path, 'week', week_number, day_name, workspace_name)

    os.makedirs(path, exist_ok=True)
    return path


def is_exportable_report(token, group_id, report_id):
    headers = {'Authorization': f'Bearer {token}'}
    report_url = f'https://api.powerbi.com/v1.0/myorg/groups/{group_id}/reports/{report_id}'
    report_response = requests.get(report_url, headers=headers)
    report = report_response.json()

    #print(report)

    # if report.get('isFromPbix', False):
    #     return False, "Relatório não foi gerado por PBIX - não pode exportar PBIX"

    # if report.get('isPaginated', False):
    #     return False, "Relatório paginado - não pode exportar PBIX"
    

    dataset_id = report.get('datasetId')
    dataset_url = f'https://api.powerbi.com/v1.0/myorg/datasets/{dataset_id}'
    dataset_response = requests.get(dataset_url, headers=headers)
    dataset = dataset_response.json()
    #print(dataset)

    if not dataset_id:
        return False, "Reports withou dataset - PBIX can not be exported"

    dataset_url = f'https://api.powerbi.com/v1.0/myorg/groups/{group_id}/datasets/{dataset_id}'
    dataset_response = requests.get(dataset_url, headers=headers)
    dataset = dataset_response.json()

    print(dataset.get('mode'))

    if dataset.get('datasetMode') != 'Import':
        return False, f"Dataset as {dataset.get('datasetMode')} mode - PBIX can't be exported"

    return True, "OK"

## function to export pbix
def export_report(token, group_id, report_id, report_name, output_dir):
    import time
    exportable, reason = is_exportable_report(token, group_id, report_id)

    if not exportable:
        msg= f"Relatório {report_name} não pode ser exportado: {reason}"
        #print(msg)
        return {'status': 'skipped', 'message': msg}

    url = f'https://api.powerbi.com/v1.0/myorg/reports/{report_id}/Export' #f'https://api.powerbi.com/v1.0/myorg/groups/{group_id}/reports/{report_id}/ExportTo'
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
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
                    for chunk in r.iter_content(chunk_size=1024*1024):
                        f.write(chunk)
                print(f'{report_name}.pbix salvo com sucesso')
                break
            elif r.status_code == 202:
                time.sleep(2)
            else:
                print(f"Error to export {report_name}: {r.status_code} - {r.text}")
                break
    else:
        print(f"Error starting export: {response.status_code} - {response.text}")
