# Databricks notebook source
# MAGIC %run "/Workspace/Users/oumaima.idaline@external.stellantis.com/qfs_classification/qfs_sharepoint/Sharepoint_utils1"
# MAGIC

# COMMAND ----------

import requests
import pyspark.sql.functions as F
import datetime
from pyspark.sql import Window
from pyspark.sql.window import Window
from pyspark.sql import Row
from pyspark.sql.functions import *
# Parameters
scope_name = 'STL-EMEA-SCOPE_QFS'
tech_user_name = 'sd20565_qfs'  # User for MFG scope
tech_user_pw = dbutils.secrets.get(scope=scope_name, key=tech_user_name)

client_id = '020a1efc-76a1-408d-a199-d9e1630c2367'
client_secret = dbutils.secrets.get(scope_name, key='STL-EMEA-LKH-LAB-SHP-QFS_PRS')
tenant_id = 'd852d5cd-724c-4128-8812-ffa5db3f8507'

# SharePoint site info
site_url = "https://shiftup.sharepoint.com/sites/QFS_PRS/"
site_id = 'shiftup.sharepoint.com,a06cecd0-c30e-467a-9c01-877adf64d816,53441b04-9613-423f-bb16-f3ee0184394d'


access_token = get_access_token(client_id, client_secret, tenant_id)
drive_id= get_drives(site_id, access_token)

# COMMAND ----------

from datetime import datetime
import json
folder_path_json = '/status' 
files = list_files_in_folder(site_id, drive_id, folder_path_json, access_token)
print(f"Files in folder '{folder_path_json}':")
for file in files:
        print(f"Name: {file['name']}, Size: {file['size']} bytes, Created: {file['createdDateTime']},LastModified: {file['lastModifiedDateTime']}")

file_id = file['id']
file_name = file['name']
file_stream = download_file_from_sharepoint(site_id, drive_id, file_id, access_token)

if file_stream:
    if file_name.endswith('.json'):
        status_list = load_status(file_stream) if file_stream else []
    else:
        raise ValueError("❗ Format de fichier non supporté")

print("Données chargées dans status file .")
status_list

# COMMAND ----------

from datetime import datetime
folder_path = '/input' 
files = list_files_in_folder(site_id, drive_id, folder_path, access_token)
print(f"Files in folder '{folder_path}':")
for file in files:
        print(f"Name: {file['name']}, Size: {file['size']} bytes, Created: {file['createdDateTime']},LastModified: {file['lastModifiedDateTime']}")
latest_file = sorted(
    files,
    key=lambda file: datetime.strptime(file['lastModifiedDateTime'], '%Y-%m-%dT%H:%M:%SZ')
)[-1]

print(f"The latest modified file in folder '{folder_path}':")
print(f"Name: {latest_file['name']}, Size: {latest_file['size']} bytes, Created: {latest_file['createdDateTime']},id: {latest_file['id']}, LastModified: {latest_file['lastModifiedDateTime']}")

# COMMAND ----------

file_id = latest_file['id']
file_name = latest_file['name']

if is_file_already_processed(file_name,status_list):
    print(f"Le fichier '{file_name}' a déjà été traité. Arrêt du traitement")
    dbutils.notebook.exit("Fichier déjà traité")  
else:
    print(f"Le fichier '{file_name}' n'a pas encore été traité. On continue le traitement")
    
    file_stream = download_file_from_sharepoint(site_id, drive_id, file_id, access_token)

    if file_stream:
        if file_name.endswith('.csv'):
            data = pd.read_csv(file_stream, encoding='ISO-8859-1')
        elif file_name.endswith(('.xls', '.xlsx')):
            data = pd.read_excel(file_stream)
        else:
            raise ValueError("❗ Format de fichier non supporté")

        print("Données chargées dans pandas.")
        display(data)


# COMMAND ----------

#processing
verbatims_df=data
data_Topic = pd.read_excel('/Workspace/Users/oumaima.idaline@external.stellantis.com/qfs_classification/qfs_sharepoint/Topic_List_VF.xlsx') 
data_Topic = data_Topic.drop(columns=['x'])
topics_df = data_Topic[['Topic', 'Explanation_Clean']]
topics_df['clean_description'] = topics_df['Explanation_Clean'].apply(cleaned_Verbatims_preprocessing)
topic_dict = topics_df.set_index('Topic')['clean_description'].to_dict()
verbatims_df['clean_verbatim'] = verbatims_df['Verb English'].apply(cleaned_Verbatims_preprocessing)
unique_verbatims_df = verbatims_df[['clean_verbatim']].drop_duplicates()
unique_verbatims_df['LLM_Result'] = unique_verbatims_df['clean_verbatim'].progress_apply(classifer_verbatim)
parsed_columns = unique_verbatims_df['LLM_Result'].apply(parse_classification_output)
parsed_df = pd.json_normalize(parsed_columns)
# Reset index before concatenation
merged_df_reset = unique_verbatims_df.reset_index(drop=True)
parsed_df_reset = parsed_df.reset_index(drop=True)
df_final = pd.concat([merged_df_reset, parsed_df_reset], axis=1)
df_final['Trust_score'] = df_final['Trust_score'].astype(str)
df_final['Feature_AI'] = df_final['Feature_AI'].apply(cleaned_Output_postprocessing)
df = verbatims_df.merge(df_final, on='clean_verbatim')

# COMMAND ----------

import requests

def refresh_access_token(client_id, client_secret, refresh_token, tenant_id):
    """
    Rafraîchit le token d'accès OAuth2 pour Microsoft Graph.
    """
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    data = {
        'grant_type': 'refresh_token',
        'client_id': client_id,
        'client_secret': client_secret,
        'refresh_token': refresh_token,
        'scope': 'https://graph.microsoft.com/.default'
    }

    response = requests.post(token_url, data=data)

    if response.status_code == 200:
        token_data = response.json()
        print("Access token refreshed.")
        return token_data.get("access_token")
    else:
        print(f"Failed to refresh token: {response.status_code} - {response.text}")
        return None


def upload_file_to_sharepoint2(
    site_id, drive_id, folder_path, access_token,
    file_stream, file_name,
    client_id=None, client_secret=None,
    refresh_token=None, tenant_id=None
):
    """
    Upload un fichier vers SharePoint via Microsoft Graph API.
    Tente de rafraîchir le token en cas d'expiration (401).
    """

    folder_path = folder_path.lstrip('/')  # Nettoyer le chemin du dossier
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path}/{file_name}:/content"

    if not access_token:
        print(" Aucun token d'accès fourni.")
        return None

    file_stream.seek(0)  # S'assurer que le stream est bien au début

    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/octet-stream'
    }

    response = requests.put(upload_url, headers=headers, data=file_stream.read())

    if response.status_code in [200, 201]:
        response_json = response.json()
        uploaded_file_id = response_json.get('id')
        print(f"Fichier '{file_name}' téléchargé avec succès. ID: {uploaded_file_id}")
        return uploaded_file_id

    elif response.status_code == 401:
        print("⚠️ Token expiré. Tentative de rafraîchissement...")

        # Si les infos de client OAuth sont disponibles, rafraîchir le token
        if all([client_id, client_secret, refresh_token, tenant_id]):
            new_access_token = refresh_access_token(client_id, client_secret, refresh_token, tenant_id)
            if new_access_token:
                # Recommencer l'upload avec le nouveau token
                file_stream.seek(0)  # Important : reset stream avant ré-essai
                return upload_file_to_sharepoint2(
                    site_id, drive_id, folder_path, new_access_token,
                    file_stream, file_name,
                    client_id, client_secret, refresh_token, tenant_id
                )
            else:
                print(" Échec du rafraîchissement du token.")
        else:
            print("Impossible de rafraîchir le token : informations manquantes.")
    else:
        print(f" Erreur lors du téléchargement : {response.status_code} - {response.text}")

    return None


# COMMAND ----------

# MAGIC %md
# MAGIC

# COMMAND ----------

import pandas as pd
from datetime import datetime, timezone
from io import BytesIO
access_token = get_access_token(client_id, client_secret, tenant_id)
basename = os.path.splitext(file_name)[0]
file_name_output = f'{basename}_{datetime.now().date()}.xlsx'

excel_buffer = BytesIO()
with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False)
excel_buffer.seek(0)

output_folder_path = "Output"

uploaded_file_id = upload_file_to_sharepoint2(
    site_id,
    drive_id,
    output_folder_path,
    access_token,
    excel_buffer,
    file_name_output
)


# COMMAND ----------

if uploaded_file_id:
    log_file=add_file_to_log(
        status_list,
        input_id=file_id,
        input_name=file_name,
        output_id=uploaded_file_id,
        output_name=file_name_output
    )

# COMMAND ----------

file_name_output = 'processed_files_log.json'
import json
json_buffer = BytesIO()
json_buffer.write(json.dumps(log_file).encode('utf-8'))
json_buffer.seek(0)

output_folder_path = "status"

uploaded_file_id = upload_file_to_sharepoint2(
    site_id,
    drive_id,
    output_folder_path,
    access_token,
    json_buffer,
    file_name_output
)
display(f"Upload du fichier '{file_name_output}' terminé.")

# COMMAND ----------

