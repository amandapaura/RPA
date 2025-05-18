import json
import pbirest
import msal
from extensao import *;
from tqdm import tqdm
import os
import datetime
from api_powerbi import *;

# ----------------------------------------------------
# -------------- EXECUTION --------------------------
# -----------------------------------------------------

ws_rel = {}

try:
    ## ---- Connecting to the Power BI service -----------
    token = get_token()

    #---- Listing workspaces-----
    workspaces = get_workspaces(token).get('value',[])

    export_failures = {}

    for ws in tqdm(workspaces, desc='Workspaces'):
        workspace_name = ws['name']
        workspace_id = ws['id']

        ## listing reports
        reports = get_reports(token, workspace_id ).get('value',[])

        if workspace_id not in ws_rel:
            ws_rel[workspace_id] = {}
            ws_rel[workspace_id]['ws_name'] = workspace_name
            ws_rel[workspace_id]['relatorios'] = {}

        for report in reports:
            url = report['webUrl']
            rep_id = report['id']
            ws_rel[workspace_id]['relatorios'][rep_id] = url


except BaseException as erro:
       erro_traceback = traceback.format_exc()

#-- ---------------------------------------------------------------------------------------------------
# Saving as JSON
with open('workspaces_reports.json', 'w', encoding='utf-8') as f:
    json.dump(ws_rel, f, indent=4, ensure_ascii=False)

# ------------ log final ---------------
print('Json file created. Export successfuly done')
