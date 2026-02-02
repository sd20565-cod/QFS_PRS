# Databricks notebook source
# MAGIC %md
# MAGIC # sharepoint_utils notebook

# COMMAND ----------

#  Create secret scope and store credentials
# create_secret_scope(scope_name, dbutils.secrets.get(scope_name, key="databricks_token"))
# store_secret_in_databricks(scope_name, 'client_id', client_id, dbutils.secrets.get(scope_name, key="databricks_token"))
# store_secret_in_databricks(scope_name, 'client_secret', client_secret, dbutils.secrets.get(scope_name, key="databricks_token"))
# store_secret_in_databricks(scope_name, 'tenant_id', tenant_id, dbutils.secrets.get(scope_name, key="databricks_token"))

# COMMAND ----------

# MAGIC %pip install databricks-agents
# MAGIC dbutils.library.restartPython()
# MAGIC !pip install openpyxl
# MAGIC import pandas as pd

# COMMAND ----------

# MAGIC %pip install mlflow
# MAGIC %pip install tqdm
# MAGIC %pip install xlsxwriter

# COMMAND ----------

import requests
import requests

# Function to call Databricks API
def call_databricks_api(endpoint, method="GET", data=None, token=None):
    DATABRICKS_URL = 'https://dbc-ea4006c8-365b.cloud.databricks.com'
    token = 'dapid6a187088dcb75d33b84a0bc7a63c330'
    HEADERS = {'Authorization': f'Bearer {token}'}
    url = f"{DATABRICKS_URL}{endpoint}"
    
    if method == "GET":
        response = requests.get(url, headers=HEADERS)
    elif method == "POST":
        response = requests.post(url, json=data, headers=HEADERS)
    else:
        raise ValueError("Unsupported HTTP method")
    
    return response

# COMMAND ----------

def create_secret_scope(scope_name, token):
    response = call_databricks_api('/api/2.0/secrets/scopes/create', method="POST", data={'scope': scope_name}, token=token)
    
    if response.status_code == 200:
        print(f"Secret scope '{scope_name}' created successfully.")
    else:
        print(f"Error creating secret scope: {response.status_code} - {response.text}")
        

# COMMAND ----------

def store_secrets_in_databricks(scope_name, parameters, databricks_url, token):
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    for key, value in parameters.items():
        data = {
            "scope": scope_name,
            "key": key,
            "string_value": value
        }
        response = requests.post(f'{databricks_url}/api/2.0/secrets/put', headers=headers, json=data)
        if response.status_code == 200:
            print(f"Secret '{key}' stored successfully!")
        else:
            print(f"Error storing secret '{key}': {response.json()}")



# COMMAND ----------

def store_secret_for_user(scope_name, tech_user_name, tech_user_password, databricks_url, token):
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    data = {
        "scope": scope_name,
        "key": tech_user_name,
        "string_value": tech_user_password
    }
    response = requests.post(f'{databricks_url}/api/2.0/secrets/put', headers=headers, json=data)
    if response.status_code == 200:
        print(f"Secret for user '{tech_user_name}' stored successfully!")
    else:
        print(f"Error storing secret for '{tech_user_name}': {response.json()}")






# COMMAND ----------

def store_key_value_in_databricks(databricks_url, token, scope_name, secret_key, secret_value):
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    data = {
        "scope": scope_name,
        "key": secret_key,
        "string_value": secret_value
    }
    response = requests.post(f'{databricks_url}/api/2.0/secrets/put', headers=headers, json=data)
    if response.status_code == 200:
        print(f"Secret '{secret_key}' stored successfully!")
    else:
        print(f"Error storing secret '{secret_key}': {response.json()}")




# COMMAND ----------

# MAGIC %md
# MAGIC # second part

# COMMAND ----------

def get_access_token(client_id, client_secret, tenant_id):
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }
    
    response = requests.post(token_url, data=token_data)
    
    if response.status_code == 200:
        access_token = response.json().get('access_token')
        print("Access token retrieved successfully.")
        return access_token
    else:
        print(f"Error fetching token: {response.status_code} - {response.text}")
        return None

# COMMAND ----------

def get_drives(site_id,access_token):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    HEADERS_SHAREPOINT = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
   }
    response = requests.get(url, headers=HEADERS_SHAREPOINT)
    
    if response.status_code == 200:
        drives = response.json().get('value', [])
        if drives:
            for drive in drives:
                #print(f"Drive Name: {drive['name']}, Drive ID: {drive['id']}")
                print('drive id retrieved successfully')
                return drive['id']
        else:
            print("Aucun drive trouvé.")
    else:
        print(f"Erreur lors de la récupération des drives: {response.status_code} - {response.text}")

# COMMAND ----------

# Function to list folders in a SharePoint drive
def list_folders_from_drive(site_id, drive_id, access_token):
    folders_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root/children"
    folder_names = []
    
    while folders_url:
        response = requests.get(folders_url, headers={'Authorization': f'Bearer {access_token}'})
        if response.status_code == 200:
            data = response.json()
            for item in data['value']:
                if 'folder' in item:
                    folder_names.append({
                        'id': item['id'],
                        'name': item['name'],
                    })
            folders_url = data.get('@odata.nextLink')  
        else:
            print(f"Error retrieving folders: {response.status_code} - {response.text}")
            break
    
    return folder_names

# COMMAND ----------



# COMMAND ----------

def list_folders_from_drive1(drive_id, access_token):
    folders_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
    folder_names = []
    
    while folders_url:
        response = requests.get(folders_url, headers={'Authorization': f'Bearer {access_token}'})
        if response.status_code == 200:
            data = response.json()
            for item in data['value']:
                if 'folder' in item:  # Only folders
                    folder_names.append({
                        'id': item['id'],
                        'name': item['name'],
                    })
            folders_url = data.get('@odata.nextLink')  
        else:
            print(f"Error retrieving folders: {response.status_code} - {response.text}")
            break
    
    return folder_names


# COMMAND ----------

def list_all_folders_recursive(drive_id, access_token, parent_folder_id='root'):
    folders = []

    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{parent_folder_id}/children"
    while url:
        response = requests.get(url, headers={'Authorization': f'Bearer {access_token}'})
        if response.status_code != 200:
            print(f"Error retrieving folders: {response.status_code} - {response.text}")
            break

        data = response.json()
        for item in data['value']:
            if 'folder' in item:
                # Add current folder
                folders.append({
                    'id': item['id'],
                    'name': item['name']
                })
                # Recurse into subfolder
                folders.extend(list_all_folders_recursive(drive_id, access_token, item['id']))

        url = data.get('@odata.nextLink')

    return folders


# COMMAND ----------

# Function to list files in a specific folder
def list_files_in_folder(site_id, drive_id, folder_path, access_token):
    url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path}:/children'
    response = requests.get(url, headers={'Authorization': f'Bearer {access_token}'})
    
    if response.status_code == 200:
        files = response.json().get('value', [])
        return files
    else:
        print(f"Error fetching files: {response.status_code} - {response.text}")
        return []

# COMMAND ----------

def download_file_to_databricks(site_id, drive_id, file_id, file_name, token, save_path):
    headers = {'Authorization': f'Bearer {token}'}  
    download_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_id}/content"
    response = requests.get(download_url, headers=headers)
    if response.status_code == 200:
        local_file_path = f"{save_path}/{file_name}"  
        with open(local_file_path, 'wb') as f:
            f.write(response.content)  
        
        print(f"Le fichier '{file_name}' a été téléchargé avec succès dans l'espace de travail Databricks.")
    else:
        print(f"Erreur lors du téléchargement du fichier : {response.status_code} - {response.text}")




# COMMAND ----------

import requests

def upload_file_to_sharepoint(site_id, drive_id, folder_path, access_token, local_file_path, file_name):
    with open(local_file_path, 'rb') as file_data:
        file_content = file_data.read()
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path}/{file_name}:/content"
    response = requests.put(
        upload_url,
        headers={'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/octet-stream'},
        data=file_content
    )

    if response.status_code == 201:
        print(f"File '{file_name}' uploaded successfully to SharePoint folder '{folder_path}'.")
    else:
        print(f"Error uploading file: {response.status_code} - {response.text}")



# COMMAND ----------

import requests
from io import BytesIO

def upload_file_to_sharepoint1(site_id, drive_id, folder_path, access_token, file_stream, file_name):
    """
    Args:
        site_id : SharePoint site ID.
        drive_id : Drive ID within the site.
        folder_path : Path to the folder in SharePoint (e.g., "/Output")
        access_token : OAuth2 bearer token.
        file_stream : In-memory file content.
        file_name: Desired name for the uploaded file.
    """
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path}/{file_name}:/content"


    file_stream.seek(0)

    response = requests.put(
        upload_url,
        headers={
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/octet-stream'
        },
        data=file_stream.read()
    )

    if response.status_code in [200, 201]:
        print(f" file '{file_name}'  uploaded successfully to SharePoint folder  '{folder_path}'.")
    else:
        print(f"Error uploading file: : {response.status_code} - {response.text}")


# COMMAND ----------

def upload_file_to_sharepoint2(site_id, drive_id, folder_path, access_token, file_stream, file_name):
    folder_path = folder_path.lstrip('/')  # Enlever slash si besoin
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path}/{file_name}:/content"

    file_stream.seek(0)

    response = requests.put(
        upload_url,
        headers={
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/octet-stream'
        },
        data=file_stream.read()
    )

    if response.status_code in [200, 201]:
        response_json = response.json()
        uploaded_file_id = response_json.get('id')
        print(f"File '{file_name}' uploaded successfully. ID: {uploaded_file_id}")
        return uploaded_file_id
    else:
        print(f"Error uploading file: {response.status_code} - {response.text}")
        return None


# COMMAND ----------

import requests

def upload_file_to_sharepoint2(site_id, drive_id, folder_path, access_token, file_stream, file_name):
    folder_path = folder_path.lstrip('/')  # Enlever slash si besoin
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path}/{file_name}:/content"
    
    # Vérifier que le jeton est valide
    if not access_token:
        print("No access token provided.")
        return None

    # Effectuer l'upload du fichier
    file_stream.seek(0)

    response = requests.put(
        upload_url,
        headers={
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/octet-stream'
        },
        data=file_stream.read()
    )

    if response.status_code in [200, 201]:
        response_json = response.json()
        uploaded_file_id = response_json.get('id')
        print(f"File '{file_name}' uploaded successfully. ID: {uploaded_file_id}")
        return uploaded_file_id
    else:
        print(f"Error uploading file: {response.status_code} - {response.text}")
        
        # Si le jeton est expiré, tenter de le rafraîchir (si nécessaire)
        if response.status_code == 401:  # Jeton expiré
            print("Token expired. Attempting to refresh...")
            # Logique de rafraîchissement du jeton ici (voir fonction `refresh_access_token`)
            new_access_token = refresh_access_token(client_id, client_secret, refresh_token, tenant_id)
            if new_access_token:
                return upload_file_to_sharepoint2(site_id, drive_id, folder_path, new_access_token, file_stream, file_name)
        return None


# COMMAND ----------

from io import BytesIO
import pandas as pd
import requests

def download_file_from_sharepoint(site_id, drive_id, file_id, access_token):
    download_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_id}/content"
    response = requests.get(
        download_url,
        headers={'Authorization': f'Bearer {access_token}'}
    )

    if response.status_code == 200:
        print(f" Fichier téléchargé depuis SharePoint.")
        return BytesIO(response.content)
    else:
        print(f"Erreur de téléchargement : {response.status_code} - {response.text}")
        return None


# COMMAND ----------

#json log_function:
def create_log_file_if_not_exists():
    if not os.path.exists(LOG_FILE_PATH):
        os.makedirs(os.path.dirname(LOG_FILE_PATH), exist_ok=True)
        with open(LOG_FILE_PATH, 'w') as f:
            json.dump([], f, indent=2)
        print(f"Fichier log créé à l'emplacement : {LOG_FILE_PATH}")
    else:
        print(f"Fichier log déjà existant à l'emplacement : {LOG_FILE_PATH}")

# COMMAND ----------

def load_status(file_bytes_io):
    file_bytes_io.seek(0)
    content = file_bytes_io.read().decode('utf-8')
    if content.strip():
        return json.loads(content)
    else:
        return []

# COMMAND ----------

def is_file_already_processed(file_name, processed_files):
    return any(f.get('input_name') == file_name for f in processed_files)

# COMMAND ----------

def save_processed_files(processed_list):
    os.makedirs(os.path.dirname(LOG_FILE_PATH), exist_ok=True)
    with open(LOG_FILE_PATH, 'w') as f:
        json.dump(processed_list, f, indent=2)

def add_file_to_log(processed_files, input_id, input_name, output_id, output_name):
    now_iso = datetime.now(timezone.utc).isoformat()
    processed_files.append({
        'input_id': input_id,
        'input_name': input_name,
        'output_id': output_id,
        'output_name': output_name,
        'processed_date': now_iso
    })
    return processed_files

# COMMAND ----------

# MAGIC %md
# MAGIC #Classification

# COMMAND ----------

def cleaned_Verbatims_preprocessing(text):
    text = str(text).lower()
    text = re.sub(r'[_-]+', ' ', text)           # Remplacer _ et - par espace
    text = re.sub(r'\s+', ' ', text)             # Normaliser espaces
    text = re.sub(r'[^\w\s,\.]', '', text)       # Supprimer tout sauf lettres, espaces, virgules, points
    text = re.sub(r'\.+$', '', text)             # Supprimer les points à la fin de la phrase
    text = text.strip()
    return text 


def cleaned_Output_postprocessing(text):
    text = text.replace("'", "").replace('"', "")
    text = re.sub(r"\([^)]*\)", "", text)
    text = text.replace("*", "")
    text = text.strip()
    text = re.sub(r"\s{2,}", " ", text)
    return text

# COMMAND ----------

import mlflow
import pandas as pd
from tqdm.auto import tqdm
import pandas as pd
import mlflow
import re
import os
import mlflow.deployments
import pandas as pd

# Initialiser client MLflow et endpoint
client = mlflow.deployments.get_deploy_client("databricks")
endpoint = "databricks-llama-4-maverick"

def call_llm(prompt):
    try:
        response = client.predict(
            endpoint=endpoint,
            inputs={"messages": [{"role": "user", "content": prompt}]}
        )
        if "choices" in response and len(response["choices"]) > 0:
            return response["choices"][0]["message"]["content"].strip()
        else:
            print("Réponse LLM invalide:", response)
            return "LLM_INVALID_RESPONSE"
    except Exception as e:
        print("Erreur lors de l'appel LLM:", e)
        return "LLM_ERROR"


tqdm.pandas()

def classifer_verbatim(verbatim):
    prompt_template = """
You are a senior expert in Powertrain (PRS) diagnostics. Your task is to analyze a customer complaint (verbatim) related to a vehicle issue.
 
Authorized list of technical categories (features/topics):  

[{'4wd malfunction light': '4wd four wheel drive malfunction light is a warning indicator that appears on the dashboard of a vehicle equipped with a 4wd system. this light is usually accompanied by a message or symbol on the dashboard, such as 4wd or a picture of a 4wd icon',
 '4x4 issue': '4x4 issue refers to a problem with a four wheel drive 4wd or all wheel drive awd vehicle. this type of vehicle is designed to send power to all four wheels, either simultaneously or through a system that can switch between two wheel drive 2wd and 4wd modes. the 4x4 system is typically used in vehicles that need to navigate rough terrain, such as off road or in snowy conditions',
 'ac evaporation cooler issue': 'an evaporative cooling system, also known as an evaporative cooling system or evap system, is a critical component of a vehicles cooling system. its main function is to cool the air that enters the engine, which helps to regulate the engines temperature. the system uses a combination of air and water to cool the air, and its usually located in the hood of the vehicle',
 'ac fan issue': 'an ac fan issue refers to a problem with the air conditioning system in your vehicle that prevents the fan from working properly. the ac fan is responsible for circulating air through the vents and cooling the cabin. when the fan is not working, the air in the vehicle can become hot and uncomfortable',
 'ac intermittent operation issue': 'an intermittent operation issue refers to a problem where a vehicles engine or system malfunctions or fails to function properly at certain times, but not always. this means that the issue may appear and disappear randomly, making it difficult to diagnose and repair',
 'ac leak issue': 'an ac leak is a common issue that occurs when the air conditioning system in your vehicle loses refrigerant, which is the liquid that helps cool the air. this can cause the air conditioning to stop working or not work properly',
 'acceleration bucking issue': 'an acceleration bucking issue occurs when a vehicles engine hesitates or sputters when accelerating, often accompanied by a sudden loss of power or a bucking or jerking motion. this can make it difficult to smoothly accelerate from a standstill or merge onto a busy road',
 'acceleration delay issue': 'an acceleration delay issue occurs when a vehicle takes longer than usual to accelerate from a standstill or when shifting gears. this can cause the vehicle to feel sluggish, slow, or unresponsive, making it difficult to merge onto highways, climb hills, or accelerate from a stop',
 'acceleration gear shift issue': 'an acceleration gear shift issue occurs when the transmission fails to smoothly shift gears when accelerating, resulting in a jerky or abrupt change in speed. this can cause the vehicle to hesitate, stall, or jerk forward when accelerating from a standstill',
 'acceleration hesitation issue': 'acceleration hesitation, also known as hesitation or sluggish acceleration, refers to a situation where a vehicles engine or transmission fails to provide the expected power or speed when the driver presses the accelerator pedal. this can cause the vehicle to feel like its struggling to accelerate or catch up with the drivers expectations',
 'acceleration irregularity issue': 'an acceleration irregularity issue occurs when a vehicles acceleration feels uneven or jerky, making it difficult to smoothly accelerate from a standstill or merging onto a highway. this can be frustrating and affect the overall driving experience',
 'acceleration issue': 'an acceleration issue refers to a problem with a vehicles ability to accelerate, which means it doesnt accelerate as quickly or smoothly as it should. this can be a frustrating experience, especially when merging onto a highway or climbing a steep hill',
 'acceleration jerk issue': 'an acceleration jerk is a sudden, abrupt, and unpleasant feeling when accelerating from a standstill or when shifting gears. its like a jolt or a jerk that makes the vehicle feel unstable or uncomfortable to drive',
 'acceleration surge issue': 'an acceleration surge is a sudden, unexpected increase in power or speed when accelerating from a standstill or from a low speed. its like a jolt or a jerk that makes the vehicle feel like its surging or jerking forward',
 'air conditioning issue': 'an air conditioning issue refers to a problem with the vehicles air conditioning system, which is designed to cool the air inside the vehicle. this system typically consists of a compressor, condenser, evaporator, and refrigerant, which work together to cool the air',
 'air conditioning odor issue': 'an air conditioning odor issue refers to a problem where your vehicles air conditioning system is emitting a unpleasant smell, often described as musty, mildewy, or moldy. this odor can be a sign of a more serious issue with the air conditioning system, and its essential to address it promptly to prevent further damage and potential health risks',
 'air filter damage issue': 'an air filter damage issue occurs when the air filter in your vehicle becomes clogged or damaged, which can negatively impact the performance and fuel efficiency of your car',
 'air vent odor issue': 'an air vent odor issue refers to a problem where your vehicles air conditioning or heating system is emitting a strong, unpleasant smell, often described as musty, mildewy, or stale. this odor can be caused by a variety of factors, and its not just a simple matter of cleaning the vents',
 'auto stop start issue': 'an auto stopstart issue occurs when your vehicles engine fails to start or shuts off unexpectedly while its running, often accompanied by a loss of power or a stalling sensation',
 'battery charging issue': 'a battery charging issue occurs when a vehicles battery is not holding a charge or is not charging properly, which can lead to a range of problems, including a dead battery, slow engine crank, and electrical system malfunctions',
 'battery disconnection issue': 'a battery disconnection issue occurs when the electrical connection between the battery and the vehicles electrical system is broken or loose, preventing the battery from functioning properly',
 'battery issue': 'a battery issue refers to problems with the vehicles battery, which is responsible for starting the engine, powering accessories, and providing power to the electrical systems. a faulty battery can cause a range of issues, from starting problems to electrical system malfunctions',
 'battery life issue': 'a battery life issue refers to a decrease in the overall performance and lifespan of a vehicles battery, resulting in reduced starting ability, slow engine crank, or complete failure to start the engine',
 'battery mode switch issue': 'a battery mode switch issue occurs when the vehicles battery mode switch is not functioning correctly, causing the vehicles electrical system to malfunction. the battery mode switch is a crucial component that allows the vehicle to switch between different electrical modes, such as starting the engine, running accessories, and charging the battery',
 'battery performance issue': 'a battery performance issue refers to a problem with a vehicles battery that affects its ability to hold a charge, start the engine, or provide power to the electrical systems',
 'blend door issue': 'a blend door is a small, sliding panel that covers the air vent in a vehicles dashboard. its usually located on the side of the dashboard and is designed to blend in with the surrounding trim. when the blend door is not functioning properly, it can cause a rattling or squeaking noise, and may also prevent the air conditioning or heating system from working correctly',
 'brake hesitation issue': 'brake hesitation is a phenomenon where the vehicles brakes suddenly lose their effectiveness, causing the car to jerk or stutter when applying the brakes, often accompanied by a delay in stopping or a feeling of spongy braking',
 'catalytic converter issue': 'a catalytic converter is a crucial component in a vehicles exhaust system that helps reduce emissions by converting pollutants and gases into harmless substances. its essentially a special kind of filter that helps keep the air clean',
 'charge and battery issue': 'a charge and battery issue refers to problems related to the charging system of a vehicles battery, which is responsible for replenishing the batterys energy when the engine is not running. this system is crucial for starting the engine, powering accessories, and maintaining the overall health of the battery',
 'charge delay issue': 'a charge delay issue occurs when a vehicles electrical system takes longer than usual to recharge the battery or charge the battery to its full capacity. this can cause the vehicle to start slowly, stall, or not start at all',
 'charge duration issue': 'a charge duration issue refers to a problem where a vehicles battery or electrical system is not holding a charge for an extended period, often resulting in a dead battery or erratic electrical behavior',
 'charge equipment damage issue': 'a charge equipment damage issue occurs when there is a problem with the electrical system of a vehicle that prevents it from charging the battery properly. this can lead to a range of issues, including a dead battery, slow engine crank, and even complete engine failure',
 'charge operation issue': 'a charge operation issue refers to a problem that prevents a vehicles electrical system from functioning properly, which can affect the vehicles ability to start, run, or maintain a steady charge',
 'charge overheat issue': 'a charge overheat issue occurs when the electrical system in your vehicles battery, starter, or alternator becomes too hot, causing it to malfunction or fail',
 'charge schedule issue': 'a charge schedule issue refers to a problem with the vehicles onboard computer system that controls the charging system. the charging system is responsible for recharging the battery and powering the electrical systems in the vehicle. the charge schedule is a set of instructions that tells the computer when to charge the battery and when to stop',
 'charge station compatibility issue': 'a charge station compatibility issue occurs when a vehicles onboard computer obd ii system is unable to communicate with a charging station, preventing the vehicle from charging or displaying the correct charging information',
 'charge system error': 'charge system error occurs when the electrical system in your vehicle is not functioning properly, preventing the battery from charging or maintaining a stable voltage. this can lead to a range of problems, from a dead battery to a faulty electrical system',
 'check engine issue': 'a check engine issue, also known as a trouble code, is a message that appears on your dashboard when the onboard computer obd ii detects a problem with your vehicles engine or emissions system. this message is usually accompanied by a warning light on the dashboard, which is often accompanied by a symbol or a code number',
 'clutch performance issue': 'a clutch performance issue refers to problems with the clutch system in a vehicle, which is responsible for engaging and disengaging the engine from the transmission. the clutch is a critical component that allows the driver to shift gears smoothly and efficiently',
 'cold start issue': 'a cold start issue occurs when a vehicles engine struggles to start or takes a long time to start, even when the ignition is turned on. this can be frustrating and may lead to a delayed departure, especially in cold weather',
 'cold start vibration issue': 'a cold start vibration, also known as a cold start shudder or cold start shake, is a common problem that occurs when a vehicles engine vibrates or shakes when its first started, especially in cold weather. this vibration can be felt through the steering wheel, seat, and even the entire vehicle',
 'cold weather start issue': 'a cold weather start issue occurs when a vehicle has trouble starting in cold temperatures, often requiring multiple attempts or a significant amount of time to start the engine',
 'component replacement issue': 'a component replacement issue occurs when a part or component within a vehicle fails or becomes damaged, requiring replacement to ensure the vehicles safe operation and prevent further damage. this can be due to wear and tear, age, or other factors',
 'coolant leak issue': 'a coolant leak is a situation where the coolant, also known as antifreeze, is leaking out of the vehicles cooling system. coolant is a liquid that helps regulate the engines temperature, keeping it from getting too hot or too cold. its usually a mixture of water and a special liquid that prevents the water from freezing in cold temperatures',
 'cruise control issue': 'a cruise control issue occurs when the vehicles cruise control system fails to maintain a set speed, either too high or too low, or stops working altogether',
 'cylinder cutoff issue': 'a cylinder cutoff issue occurs when the engines fuel and air mixture is not properly ignited, resulting in a loss of power and potentially causing the engine to stall or run rough',
 'dashboard warning light issue': 'a dashboard warning light is a visual indicator on your vehicles dashboard that alerts you to a potential problem or issue with your vehicles systems. these lights are usually accompanied by a beep or chime to grab your attention',
 'deceleration issue': 'a deceleration issue occurs when a vehicles braking system is not functioning properly, causing the vehicle to slow down or stop unevenly, or not responding to the drivers braking input',
 'delay gear shifting issue': 'a delay gear shifting issue occurs when the transmission in your vehicle takes longer than usual to shift gears, often resulting in a hesitation or lag between gear changes. this can cause the vehicle to jerk or stutter when shifting, making it uncomfortable to drive',
 'delay start issue': 'a delay start issue occurs when a vehicle takes longer than usual to start, often requiring multiple attempts to start the engine',
 'delay throttle response issue': 'a delay in throttle response occurs when the engines throttle doesnt respond quickly enough to your pedal input, making it feel sluggish or hesitant when you press the accelerator. this can make driving uncomfortable and affect your overall driving experience',
 'dial shifter issue': 'a dial shifter is a type of transmission control system used in some vehicles, particularly in manual transmissions. its a lever or knob that you use to select the gear you want to shift into. if the dial shifter is not working properly, it can cause problems with shifting gears smoothly and efficiently',
 'downshift braking issue': 'a downshift braking issue occurs when the transmission fails to downshift properly, causing the vehicle to continue accelerating or not slow down as expected when the driver presses the brake pedal. this can lead to a loss of control, increased stopping distance, and potentially even a crash',
 'dpf issue': 'a dpf diesel particulate filter is a device that helps reduce emissions from diesel engines by trapping soot and other pollutants. its like a air filter, but for diesel engines',
 'dpf regeneration frequency issue': 'a dpf diesel particulate filter is a device that helps reduce emissions from diesel engines by trapping soot and other pollutants. the dpf regeneration process is a crucial maintenance task that helps maintain the filters effectiveness. however, some modern diesel engines have a feature called dpf regeneration frequency or dpf light on that indicates when the filter needs to be regenerated. this feature is designed to ensure the filter is functioning correctly and prevent damage to the engine',
 'drive inability issue': 'a drive inability issue refers to a problem that prevents a vehicle from moving or functioning properly while in motion. this can manifest in various ways, such as difficulty starting the engine, stalling, or struggling to accelerate. in some cases, the vehicle may not move at all',
 'drive irregularity issue': 'a drive irregularity issue refers to a problem with the way your vehicle moves or responds to your inputs while driving. this can include issues with acceleration, braking, shifting, or steering, making it difficult to control the vehicle or maintain a steady pace',
 'drive mode transition issue': 'a drive mode transition issue occurs when the vehicles transmission or drivetrain system fails to smoothly switch between different drive modes, such as from park to drive, or from drive to reverse. this can cause the vehicle to jerk, hesitate, or make unusual noises when shifting gears',
 'drivetrain noise issue': 'a drivetrain noise issue refers to a problem with the system that transmits power from the engine to the wheels, which includes the transmission, driveshaft, axles, and differential. this system is responsible for propelling the vehicle forward, and any noise coming from it can be unsettling and distracting',
 'egr issue': 'egr exhaust gas recirculation issue refers to a problem with the exhaust gas recirculation system in your vehicles engine. the egr system helps reduce emissions by recirculating a portion of the exhaust gases back into the engines cylinders to reduce the amount of oxygen in the exhaust gases, which in turn reduces the amount of nitrogen oxides nox emitted',
 'electric mode issue': 'an electric mode issue refers to a problem that occurs when a vehicles electric motor or electric powertrain is not functioning correctly, resulting in a loss of electric mode functionality. this can manifest in various ways, such as reduced or no electric only driving range, erratic acceleration, or complete loss of electric mode altogether',
 'electric motor switch over issue': 'an electric motor switch over issue occurs when the vehicles electric motor is not functioning correctly, causing the vehicle to stall or shut down unexpectedly. this can be a frustrating and potentially safety critical issue',
 'electrical burning smell issue': 'an electrical burning smell in a vehicle is a strong, unpleasant odor that is often associated with electrical components or wiring issues. this smell can be caused by a variety of factors, including overheated electrical systems, faulty wiring, or malfunctioning components',
 'electrical component issue': 'an electrical component issue refers to a problem with a vehicles electrical system, which includes the battery, wiring, sensors, and other electronic components that control various vehicle functions such as lights, accessories, and safety features. this issue can affect the overall performance, safety, and reliability of the vehicle',
 'electrical connection issue': 'an electrical connection issue occurs when there is a problem with the way the electrical system in your vehicle is connected. this can cause a range of symptoms, from minor annoyances to major system failures',
 'electrical issue': 'an electrical issue in a vehicle refers to a problem that affects the electrical system of the car, which includes the battery, wiring, and electrical components such as lights, accessories, and sensors. this can cause a range of symptoms, from minor annoyances to major safety hazards',
 'electronic stabilizing system': 'an electronic stabilizing system, also known as an electronic stability control esc system, is a safety feature in modern vehicles that helps the vehicle stay stable and on course, especially when driving on slippery or uneven roads. it works by using sensors and computer controls to detect when the vehicle is losing traction or stability, and then makes adjustments to the engine, transmission, and brakes to help the vehicle stay stable',
 'emission control sensor issues': 'an emission control sensor is a device that helps your vehicles engine run efficiently and meet environmental regulations by monitoring and controlling the amount of pollutants released into the air. it does this by detecting the levels of certain gases, such as oxygen, carbon monoxide, and nitrogen oxides, in the exhaust system',
 'emission fault issue': 'an emission fault, also known as an emission problem or emission issue, refers to a malfunction in a vehicles emissions system that prevents the vehicle from meeting the required emissions standards. this can be caused by a variety of factors, including faulty sensors, clogged air filters, or problems with the catalytic converter',
 'engine breakdown issue': 'an engine breakdown refers to a situation where the engine of a vehicle stops working or fails to run properly, often resulting in a complete loss of power or complete engine failure',
 'engine cold start issue': 'an engine cold start issue occurs when a vehicles engine struggles to start or takes a long time to warm up when the engine is cold. this can be frustrating and may lead to poor fuel efficiency, reduced performance, and even engine damage if left unchecked',
 'engine component issue': 'an engine component issue involves a fault in parts like sensors or spark plugs, leading to poor performance and warning lights. quick diagnosis and repair help prevent further damage',
 'engine damage issue': 'engine damage refers to any harm or deterioration to the internal components of a vehicles engine, which can affect its performance, fuel efficiency, and overall health. this can include issues with the engines mechanical parts, such as the pistons, cylinders, crankshaft, or valves, as well as problems with the engines electrical and fuel systems',
 'engine dies stalls while driving': 'an engine stall occurs when the engine suddenly stops running or dies while the vehicle is in motion. this can be a frustrating and potentially hazardous situation, especially if it happens while driving on a busy road',
 'engine exhaust leak issue': 'an engine exhaust leak occurs when there is a hole or gap in the exhaust system that allows exhaust gases to escape, causing a loss of power, decreased fuel efficiency, and potentially leading to engine damage',
 'engine hesitation issue': 'engine hesitation is a phenomenon where the engine suddenly loses power or feels like its struggling to find its rhythm, causing the vehicle to jerk or stumble. this can be a frustrating and unsettling experience for drivers',
 'engine noise during acceleration issue': 'a loud or unusual noise that occurs when you accelerate your vehicle, often accompanied by a vibration or shaking. this noise can be caused by various factors, and its essential to identify the root cause to ensure your safety on the road',
 'engine noise during deceleration issue': 'engine noise during deceleration refers to a strange sound that occurs when youre slowing down or braking, often accompanied by a rattling, clunking, or grinding noise coming from the engine. this noise can be unsettling and may indicate a problem with the engines mechanical components',
 'engine noise issue': 'an engine noise issue refers to any unusual sounds coming from the engine, such as grinding, clunking, knocking, or rumbling noises. these sounds can be caused by a variety of problems, ranging from minor issues to more serious problems that require immediate attention',
 'engine not restart issue': 'an engine not restarting issue occurs when a vehicles engine fails to start or restart after being turned off, requiring multiple attempts to start or remaining in a no start state',
 'engine oil consumption issue': 'engine oil consumption, also known as oil burn or oil loss, occurs when your vehicles engine is using more oil than it should, causing the oil level to drop faster than usual. this can lead to engine damage and decreased performance if not addressed',
 'engine oil leak issue': 'an engine oil leak occurs when engine oil escapes from the engine and onto the ground or other surfaces, causing damage to the engine and potentially leading to costly repairs',
 'engine overheat issue': 'an engine overheat issue occurs when the engines temperature rises above its normal operating temperature, causing damage to the engine and potentially leading to costly repairs',
 'engine performance issue': 'engine performance refers to how well the engine is running, including its power, efficiency, and overall responsiveness. a performance issue occurs when the engine is not running as smoothly, efficiently, or effectively as it should, resulting in decreased performance, reduced fuel efficiency, or decreased overall driving experience',
 'engine rattle issue': 'an engine rattle is a loud, annoying noise that occurs when there are loose or vibrating parts inside the engine, causing the engine to make a rattling or clattering sound',
 'engine rough idle issue': 'an engine rough idle is a condition where the engines idle speed is not smooth and steady, but instead wobbles, shakes, or vibrates. this can be unsettling and may cause the engine to consume more fuel than usual',
 'engine rough start issue': 'an engine rough start is when the engine struggles to start or stalls when you turn the key, often accompanied by a rough or uneven sound, vibration, or hesitation',
 'engine roughness issue': 'an engine roughness issue refers to a condition where the engine is not running smoothly, producing a rough, uneven, or vibrating sound, and may also be accompanied by decreased performance, decreased fuel efficiency, and potentially even engine damage if left unaddressed',
 'engine sensor issue': 'an engine sensor is a device that measures and reports data to the vehicles onboard computer ecu about the engines performance and condition. these sensors help the ecu make adjustments to optimize engine performance, fuel efficiency, and emissions',
 'engine shake issue': 'an engine shake is a vibration or tremble that you feel while driving your vehicle. its a noticeable movement that can be felt through the steering wheel, seat, and even the entire vehicle',
 'engine shut down issue': 'an engine shut down issue occurs when the engine suddenly stops running or shuts down unexpectedly, often without warning. this can be a frustrating and potentially costly problem for vehicle owners',
 'engine smoke issue': 'engine smoke is a visible sign that something is wrong with your vehicles engine. it can be a sign of a serious problem that needs attention, or it might be a minor issue that can be easily fixed. engine smoke can appear as white, blue, or black smoke, and it can be a sign of oil, coolant, or fuel leaks',
 'engine speed issue': 'an engine speed issue refers to a problem where the engine is not running at the correct speed, either too fast or too slow, which can affect the overall performance and efficiency of the vehicle',
 'engine sputter issue': 'an engine sputter is a sudden, irregular, and often intermittent loss of power or a stumbling of the engine, which can be accompanied by a decrease in engine performance, rough idling, or a complete loss of power',
 'engine stall issue': 'an engine stall is when the engine suddenly stops running or loses power, often accompanied by a loss of control or a stalling feeling',
 'engine start issue': 'an engine start issue refers to a problem that prevents a vehicles engine from starting or stalls when attempting to start. this can be frustrating and may require professional assistance to resolve',
 'engine start vibration issue': 'an engine start vibration issue occurs when the engine of a vehicle vibrates or shakes when its started, often accompanied by a slight or noticeable shaking or wobbling of the vehicle. this vibration can be felt through the steering wheel, seat, or even the entire vehicle',
 'engine stuttering issue': 'engine stuttering, also known as engine hesitation or stumbling, is a phenomenon where the engines power delivery is interrupted or interrupted, causing the vehicle to jerk or stutter when accelerating or under load. this can be a frustrating and unsettling experience for drivers',
 'engine surge issue': 'an engine surge is a sudden, brief increase in engine power or speed that can cause the engine to stumble, hesitate, or jerk. its like a brief hiccup in the engines performance',
 'engine throttling issue': 'engine throttling is a condition where the engines power output is reduced or restricted, causing the vehicle to slow down or hesitate when accelerating. this can be a frustrating and concerning issue for drivers, as it can affect the overall performance and safety of the vehicle',
 'erratic downshifting issue': 'an erratic downshifting issue occurs when the transmission shifts gears in an unpredictable or irregular manner, often causing the vehicle to hesitate, jerk, or hesitate when shifting into lower gears. this can be frustrating and affect the overall driving experience',
 'evap issue': 'an evaporative emission control evap issue occurs when the system that helps reduce emissions from a vehicles engine is not functioning properly. the evap system is designed to capture and recirculate fuel vapors that escape from the engine, preventing them from entering the atmosphere',
 'excessive exhaust smoke issue': 'excessive exhaust smoke is a condition where there is an unusual amount of smoke coming from the exhaust system of a vehicle. this can be a sign of a serious problem that needs to be addressed to prevent damage to the engine, catalytic converter, and the environment',
 'excessive noise issue': 'an excessive noise issue refers to a problem where a vehicle makes an unusual, loud, or persistent sound that can be heard while the engine is running, driving, or even when the vehicle is stationary. this noise can be caused by various factors and can be annoying and distracting for the driver',
 'excessive smoke on startup issue': 'excessive smoke on startup refers to a condition where a vehicle emits a large amount of smoke when the engine is first started, often accompanied by a burning or acrid smell. this can be a concerning issue, as it may indicate a problem with the engine or other critical systems',
 'exhaust': 'the exhaust system is a network of pipes and components that carry exhaust gases away from an engine and out of a vehicle. its main purpose is to remove waste gases, such as carbon monoxide, carbon dioxide, and other pollutants, from the engine and prevent them from entering the vehicles cabin',
 'exhaust damage issue': 'exhaust damage refers to any damage or deterioration of the exhaust system, which is responsible for carrying exhaust gases away from the engine and out of the vehicle. the exhaust system includes pipes, mufflers, catalytic converters, and tailpipes',
 'exhaust leak issue': 'an exhaust leak is a hole or gap in the exhaust system that allows exhaust gases to escape, causing a loss of power, decreased fuel efficiency, and potentially leading to more serious problems',
 'exhaust noise issue': 'an exhaust noise issue refers to a problem with the exhaust system of a vehicle, which can cause unusual sounds, vibrations,coming from the exhaust pipe, muffler, or tailpipe',
 'exhaust rattle issue': 'an exhaust rattle is a loud, annoying noise that occurs when the exhaust system of a vehicle vibrates or rattles, often due to loose or damaged components',
 'exhaust soot issue': 'exhaust soot is a type of residue that forms on the exhaust system of a vehicle, typically due to incomplete combustion of fuel. its a mixture of unburned fuel, carbon particles, and other contaminants that can cause damage to the engine and exhaust system if not addressed',
 'exhaust vibration issue': 'an exhaust vibration issue occurs when the exhaust system of a vehicle vibrates excessively, causing the vehicle to shake, rattle, or wobble, often accompanied by a loud noise',
 'extend fan operation issue': 'an extend fan, also known as a transmission fan or transmission cooler fan, is a crucial component in your vehicles cooling system. its primary function is to help keep your engine at a safe operating temperature by circulating coolant through the engine block and transmission',
 'factory installation issue': 'a factory installation issue refers to a problem that occurs when a vehicles original equipment manufacturer oem installed a component or system incorrectly, or when the installation was not done according to the manufacturers specifications. this can lead to a range of issues, from minor annoyances to major safety concerns',
 'faulty indicator issue': 'a faulty indicator issue refers to a problem with the vehicles turn signal or hazard lights, which are designed to alert other drivers of a vehicles intentions to turn or change lanes',
 'frequent def refill issue': frequent DEF refill issue refers to a situation where a vehicle's diesel exhaust fluid (DEF) level is depleted at an abnormal rate, requiring more frequent refills than expected. This can be caused by various factors, including a faulty DEF system, incorrect fluid usage, or a malfunctioning engine.,
 'fuel consumption issue': 'a fuel consumption issue refers to a problem where a vehicle uses more fuel than normal. this can happen because of engine problems, poor maintenance, or faulty parts, leading to higher fuel costs and more pollution',
 'fuel door issue': 'a fuel door issue refers to problems related to the fuel door, which is the door that opens to access the fuel tank. this door is usually located on the drivers side of the vehicle and is used to fill up the fuel tank',
 'fuel filling issue': 'a fuel filling issue refers to a problem that occurs when the fuel tank is not filling up properly or is not being filled at all. this can be a frustrating and inconvenient issue for drivers, especially when theyre on the go',
 'fuel gauge issue': 'a fuel gauge issue occurs when the fuel level indicator on your vehicles dashboard is not accurately showing the current fuel level. this can be frustrating and may lead to running out of fuel or overfilling your tank',
 'fuel leak issue': 'a fuel leak is a situation where fuel is escaping from a vehicles fuel system, which can cause various problems and potentially lead to engine damage. this can happen due to a crack or hole in the fuel tank, fuel lines, or other components',
 'fuel repair issue': 'a fuel repair issue refers to a problem that affects the fuel system of a vehicle, which is responsible for delivering fuel from the tank to the engine. this system includes the fuel tank, fuel pump, fuel filter, fuel injectors, and fuel lines',
 'fuel tank issue': 'a fuel tank issue refers to problems related to the fuel tank, which is the container that holds the fuel for your vehicle. this can include issues with the tank itself, the fuel pump, or the system that supplies fuel to the tank',
 'gas electric transition issue': 'a gas electric transition issue refers to a problem that occurs when the vehicles engine and electric motor are not working together smoothly, causing the vehicle to stall, hesitate, or struggle to accelerate. this issue can be frustrating and affect the overall performance of the vehicle',
 'gas pedal unresponsiveness issue': 'a gas pedal unresponsiveness issue occurs when the accelerator pedal does not respond properly to your foot pressure, making it difficult to accelerate the vehicle. this can cause the vehicle to stall, jerk, or hesitate when accelerating',
 'gear ratio issue': 'a gear ratio issue occurs when the ratio of the gear teeth on the wheels to the gear teeth on the transmission is not properly matched, causing problems with the vehicles acceleration, shifting, and overall performance',
 'gear shift clunk issue': 'a gear shift clunk is a strange noise or vibration that occurs when shifting gears in a vehicle, often accompanied by a clunking or grinding sound. this issue can be caused by a variety of factors, and its usually a sign that something is amiss with the transmission or drivetrain',
 'gear shift fluctuation issue': 'a gear shift fluctuation issue occurs when the transmission shifts between gears in an unpredictable or irregular manner, often causing the vehicle to jerk, hesitate, or hesitate when shifting gears. this can lead to a rough or uneven driving experience',
 'gear shift hard shifting issue': 'a gear shift hard shifting issue occurs when the transmission shifts gears abruptly or jerkily, making it difficult to smoothly transition between gears. this can be frustrating and affect the overall driving experience',
 'gear shift harshness issue': 'a gear shift harshness issue occurs when the gear shift in your vehicle feels abrupt, jerky, or unresponsive, making it uncomfortable to drive. this can be due to a problem with the transmission or the gear shift mechanism',
 'gear shift hesitation issue': 'gear shift hesitation, also known as slipping or hesitation, is a problem where the transmission shifts gears smoothly, but the car hesitates or lingers in a particular gear before shifting to the next one. this can cause the car to jerk or stutter, making it uncomfortable to drive',
 'gear shift irregularity issue': 'a gear shift irregularity issue occurs when the transmission shifts gears smoothly and consistently, but the gear shifts are not happening at the expected times or are not happening at all. this can cause the vehicle to jerk, hesitate, or stall, making it difficult to drive',
 'gear shift issue': 'a gear shift issue occurs when the transmission or gear selector is not functioning properly, causing problems with shifting gears smoothly and efficiently',
 'gear shift jerk issue': 'a gear shift jerk is a sudden, abrupt movement or jerk when shifting gears, which can be uncomfortable and sometimes even cause the vehicle to jerk or lurch forward. this issue can be frustrating and affect the overall driving experience',
 'gear shift lag issue': 'gear shift lag, also known as hesitation or delay, refers to a phenomenon where the transmission takes a moment to engage or disengage gears, resulting in a noticeable pause or hesitation when shifting gears. this can cause the vehicle to feel sluggish, unresponsive, or even stall',
 'gear shift noise issue': 'a gear shift noise is a strange sound that occurs when you shift gears in your vehicle. it can be a clicking, grinding, or clunking noise that can be annoying and sometimes even alarming',
 'gear shift pull issue': 'a gear shift pull issue occurs when the gearshift in your vehicle pulls or moves to the left or right when youre trying to shift into a specific gear, making it difficult to select the desired gear',
 'gear shift quality issue': 'a gear shift quality issue refers to a problem with the smoothness, feel, or performance of the gear shifting mechanism in a vehicle. this can include issues with the transmission, gear selector, or other related components',
 'gear shift roughness issue': 'a gear shift roughness issue occurs when the gearshift in your vehicle feels stiff, jerky, or uneven when shifting gears, making it uncomfortable to drive',
 'gear shift smoothness issue': 'a gear shift smoothness issue refers to a problem where the gear shifts in a vehicle are not smooth or seamless, often resulting in jerky or abrupt changes between gears. this can be frustrating and affect the overall driving experience',
 'gear shift speed issue': 'a gear shift speed issue occurs when the transmission shifts gears at an unusual or inconsistent speed, often resulting in a jerky or abrupt acceleration, or a delay in shifting between gears',
 'gear shift stall issue': 'a gear shift stall occurs when the transmission fails to engage or disengage gears properly, causing the vehicle to stall or jerk when shifting gears. this can happen when the transmission is not functioning correctly, leading to a loss of power and control',
 'gear shift studder issue': 'a gear shift stutter, also known as a stutter or hunting, is a condition where the transmission shifts between gears in an automatic vehicle, often causing the car to jerk or stutter when shifting. this can be a frustrating and unsettling experience for drivers',
 'gear shift variability issue': 'a gear shift variability issue occurs when the transmission shifts between gears in an unpredictable or irregular manner, often resulting in jerky or abrupt changes in speed, hesitation, or difficulty shifting into specific gears',
 'gear shift vibration issue': 'a gear shift vibration issue occurs when the transmission or gearshift in a vehicle vibrates or shakes when shifting gears, making it uncomfortable and potentially causing damage to the transmission or other components',
 'gear slip issue': 'a gear slip issue occurs when the gears in your vehicles transmission dont engage properly, causing the vehicle to hesitate or jerk when shifting gears. this can lead to a loss of power, decreased fuel efficiency, and potentially even damage to the transmission',
 'grinding noise issue': 'a grinding noise in a vehicle refers to a loud, scraping or abrasive sound that occurs when the engine, transmission, or other mech',
 'heater odor issue': 'a strong, unpleasant smell coming from the vehicles heating system, often described as burning, chemical, or musty',
 'high speed vibration issue': 'a high speed vibration issue occurs when a vehicle shakes or vibrates excessively while driving at high speeds, typically above 40 50 mph. this vibration can be felt through the steering wheel, seat, and even the entire vehicle',
 'home charge issue': 'a home charge issue refers to a problem that occurs when a vehicles electrical system is not functioning properly, preventing the vehicle from charging its battery or holding a charge. this can be a frustrating and potentially costly issue to resolve',
 'hybrid failure alert': 'a hybrid vehicle is a type of vehicle that combines a conventional internal combustion engine with an electric motor and battery pack. when a hybrid vehicles hybrid system fails, it can lead to a range of problems that affect the vehicles performance, fuel efficiency, and overall reliability',
 'hybrid mode drive experience issue': 'a hybrid mode drive experience issue refers to a problem that occurs when a hybrid vehicles powertrain is not functioning as expected, resulting in an abnormal or unpleasant driving experience. this can manifest in various ways, such as reduced fuel efficiency, decreased performance, or unusual noises',
 'hybrid mode issue': 'a hybrid mode issue refers to a problem that occurs when a vehicles hybrid system, which combines a conventional engine with an electric motor, is not functioning properly. this can lead to reduced fuel efficiency, decreased performance, and potentially even safety concerns',
 'hybrid mode warning issue': 'a hybrid mode warning issue occurs when a hybrid vehicles computer system detects a problem with the vehicles powertrain or battery, and alerts the driver to take action to prevent damage or optimize performance. this warning is usually displayed on the dashboard as a message or a light',
 'hybrid to electric transition issue': 'a hybrid to electric transition issue occurs when a hybrid vehicles electric motor and gasoline engine are not working together properly, causing the vehicle to struggle to start, idle, or move',
 'idle hesitation issue': 'idle hesitation is a condition where a vehicles engine experiences a brief pause or stutter when its idling, often accompanied by a slight decrease in power or a hick up or catch in the engines idle',
 'idle speed fluctuation issue': 'idle speed fluctuation refers to a condition where the engines idle speed varies or oscillates, often between 500 1500 rpm, instead of staying steady at a consistent speed. this can be a normal phenomenon, but in some cases, it can be a sign of an underlying problem',
 'idle stop issue': 'an idle stop issue occurs when a vehicles engine continues to run and consume fuel even when the vehicle is stationary, such as when stopped at a red light, in traffic, or parked. this can lead to decreased fuel efficiency, increased emissions, and potentially cause damage to the engine',
 'idle vibration issue': 'an idle vibration is a shaking or wobbling feeling when a vehicle is stationary, typically when the engine is idling. this vibration can be felt through the seat, floor, or steering wheel',
 'incorrect fan activation issue': 'an incorrect fan activation issue occurs when the vehicles cooling system fails to activate the fan when its needed, or the fan activates unnecessarily, causing the engine to overheat or run inefficiently',
 'increase fluid usage issue': 'an increase in fluid usage refers to a situation where a vehicles fluids, such as engine oil, transmission fluid, coolant, brake fluid, or power steering fluid, are being consumed at a faster rate than usual. this can lead to a range of problems if not addressed promptly',
 'injector replacement issue': 'an injector replacement issue occurs when a fuel injector, which is a critical component in a vehicles engine, fails to function properly. fuel injectors are responsible for spraying fuel into the engines cylinders to help the engine run efficiently and effectively',
 'insufficient acceleration power issue': 'insufficient acceleration power refers to a situation where a vehicle is not accelerating quickly enough, making it difficult to merge onto a busy road, climb hills, or accelerate from a standstill',
 'intermittent acceleration issue': 'an intermittent acceleration issue occurs when a vehicle accelerates unexpectedly or hesitates to accelerate, often without warning. this can be frustrating and may be caused by a variety of factors',
 'intermittent engine stall issue': 'an intermittent engine stall is when your cars engine suddenly stops running, but then restarts on its own without warning. this can be frustrating and may be caused by a variety of factors',
 'intermittent light issue': 'an intermittent light issue refers to a situation where a vehicles dashboard warning light comes on and off randomly, often at random intervals, and sometimes stays on for a short period. this can be frustrating and may cause concern for the driver',
 'intermittent noise issue': 'an intermittent noise issue refers to a sound that is heard from a vehicles engine, transmission, or other components, but only occasionally, and not consistently. this means that the noise may be present for a few seconds, then disappear, or come and go at random intervals',
 'intermittent electrical power loss issue': 'an intermittent electrical power loss issue occurs when a vehicles electrical system experiences a temporary loss of power, causing the vehicle to stall, hesitate, or lose acceleration. this can be frustrating and may be caused by a variety of factors',
 'intermittent rough idle issue': 'an intermittent rough idle issue occurs when a vehicles engine idles unevenly, vibrating, or stumbling, and the problem only happens occasionally. this can be frustrating and may be accompanied by other symptoms like decreased fuel efficiency, decreased performance, or decreased engine life',
 'intermittent surging issue': 'an intermittent surging issue occurs when a vehicles engine experiences sudden, brief surges or jerks in power, often accompanied by a decrease in fuel efficiency and a possible decrease in engine performance. this can be unsettling and may cause the vehicle to jerk or hesitate while driving',
 'intermittent vibration issue': 'an intermittent vibration issue refers to a phenomenon where a vehicles engine, transmission, or other components produce a vibration that comes and goes, often at random or in response to specific actions. this vibration can be felt through the steering wheel, seat, or floor of the vehicle',
 'irregular engine start issue': 'an irregular engine start issue refers to a problem where the engine doesnt start or starts intermittently, making it difficult to get the vehicle moving. this can be frustrating and may cause you to be late or stranded',
 'limited fuel capacity issue': 'a limited fuel capacity issue occurs when a vehicles fuel tank is not designed to hold enough fuel to complete a full tank, resulting in the need for frequent fuel stops or running low on fuel',
 'locker disengagement issue': 'a locker disengagement issue occurs when the locking mechanism in your vehicles differential the gearbox that transmits power to the wheels fails to engage properly, causing the wheels to spin freely or not lock up when you shift gears',
 'loud engine noise issue': 'a loud engine noise is a sound that is excessively loud and unpleasant, often accompanied by vibrations, and can be caused by various issues within the engine',
 'loud exhaust issue': 'a loud exhaust issue refers to a problem with the exhaust system of a vehicle that causes an unusual or excessive noise, often accompanied by vibrations or other unusual sounds',
 'loud fan issue': 'a loud fan issue refers to a problem with the vehicles cooling system that causes the fan to operate excessively loudly, often accompanied by unusual noises, vibrations, or reduced airflow',
 'loud interior issue': 'a loud interior issue refers to a problem where the interior of a vehicle is excessively noisy, often due to unusual sounds coming from within the cabin. this can include creaks, rattles, clunks, or other distracting noises that can be uncomfortable and distracting while driving',
 'low speed shift issue': 'a low speed shift issue occurs when the transmission shifts into a higher gear too quickly or hesitates when shifting into a lower gear, often at low speeds, such as when accelerating from a stop or when cruising at low speeds',
 'misfire issue': 'a misfire is a condition where one or more cylinders in a vehicles engine do not fire or run at the correct time, causing the engine to run rough, lose power, or even stall',
 'motor operation issue': 'a motor operation issue refers to a problem that affects the way a vehicles engine or motor runs, causing it to malfunction or not function properly. this can include issues with the engines ability to start, run, or maintain a consistent speed',
 'motor revving issue': 'a motor revving issue occurs when a vehicles engine is revving excessively or making a high pitched whining noise, often accompanied by a decrease in power or performance',
 'motor starting issue': 'a motor starting issue occurs when a vehicles engine fails to start or stalls when attempting to start the engine. this can be frustrating and may be caused by a variety of factors',
 'noisy exhaust issue': 'a noisy exhaust issue refers to a problem with the exhaust system of a vehicle that causes unusual sounds, such as rattling, clunking, hissing, or roaring noises, coming from the exhaust pipe or muffler',
 'oil filter accessibility issue': 'an oil filter accessibility issue occurs when its difficult or impossible to access the oil filter on your vehicle, making it hard to change the oil. this can be due to various reasons, such as a poorly designed oil filter housing, a clogged or obstructed area, or a lack of clearance',
 'oil pressure issue': 'oil pressure refers to the pressure of the engine oil inside the engines oil system. the oil pressure is measured in pounds per square inch psi and is essential for lubricating the engines moving parts, cooling the engine, and preventing wear and tear',
 'parking brake issue': 'a parking brake issue occurs when the parking brake also known as the handbrake or emergency brake on a vehicle is not functioning properly, making it difficult or impossible to hold the vehicle in place when parked',
 'power electronic issue': 'a power electronic issue refers to a problem that affects the electrical system of a vehicle, which is responsible for controlling and regulating the flow of electrical power to various components such as the engine, transmission, and accessories. this issue can manifest in various ways, including problems with the battery, starter motor, alternator, and other electrical systems',
 'power fluctuation issue': 'a power fluctuation issue refers to a problem where the electrical system in your vehicle experiences irregular or unstable power supply, causing the engine, accessories, and other systems to malfunction or behave erratically. this can lead to a range of symptoms, from minor annoyances to safety hazards',
 'power loss during acceleration issue': 'power loss during acceleration refers to a situation where a vehicle experiences a sudden decrease in power or speed when accelerating, often accompanied by a decrease in engine performance',
 'power loss during drive issue': 'a power loss during drive refers to a situation where a vehicles engine or electrical system loses power or slows down while the vehicle is in motion. this can be a concerning issue, as it can affect the vehicles safety and performance',
 'power loss incident issue': 'a power loss incident refers to a situation where a vehicles electrical system suddenly loses power, causing the vehicle to shut down or become unresponsive. this can be a frustrating and potentially hazardous situation, especially if it occurs while driving',
 'power loss issue': 'a power loss issue refers to a situation where a vehicles engine or electrical system is not producing enough power or is not functioning properly, resulting in a decrease in performance, speed, or overall driving experience',
 'power module issue': 'a power module issue refers to a problem with the electrical system of a vehicle that affects the way the vehicles computer ecu, or engine control unit controls the engine, transmission, and other essential systems. the power module is a critical component that manages the flow of electrical power to various systems in the vehicle',
 'power source switch issue': 'a power source switch is a crucial component in your vehicles electrical system that controls the flow of power to various electrical systems, such as the battery, starter motor, and accessories. its essentially a switch that determines whether the vehicles electrical system is receiving power or not',
 'power steering issue': 'power steering is a system that helps make it easier to turn your vehicle by using a pump to assist the effort of turning the steering wheel. when the power steering system is working properly, it reduces the amount of effort needed to turn the wheel, making it easier to maneuver the vehicle',
 'powertrain hesitation issue': 'powertrain hesitation is a phenomenon where a vehicles engine or transmission seems to hesitate or stumble when accelerating, often accompanied by a slight delay or pause in power delivery. this can make the vehicle feel sluggish, unresponsive, or even like its catching up to the drivers input',
 'powertrain performance issue': 'a powertrain performance issue refers to problems with the engine, transmission, and drivetrain of a vehicle that affect its ability to run smoothly, efficiently, and effectively. this can include issues with acceleration, shifting, and overall engine performance',
 'rattle noise issue': 'a rattle noise is a loud, vibrating sound that occurs when there is loose or moving parts inside a vehicle that are not supposed to be making noise. this noise can be caused by a variety of factors, including worn or damaged parts, loose connections, or improper installation',
 'refrigerant loss issue': 'refrigerant is a liquid that helps keep your vehicles air conditioning system cool. its like the coolant in your cars engine, but instead of keeping the engine warm, it keeps the air inside the car cool. when refrigerant leaks out, it can cause the air conditioning system to stop working properly',
 'repair activity issue': 'repair activity issue',
 'repair delay issue': 'a repair delay issue occurs when a vehicles repair process takes longer than expected, causing the customer to wait for a longer period before their vehicle is ready for pickup or return. this can be frustrating for customers and may lead to additional costs or inconvenience',
 'ride discomfort issue': 'ride comfort refers to the overall smoothness and stability of a vehicles ride, making it comfortable for the driver and passengers to travel. a ride discomfort issue occurs when the vehicles suspension, steering, or other components dont work together properly, causing an unpleasant ride',
 'road noise in cabin issue': 'road noise in the cabin refers to a distracting or annoying sound that is heard inside the vehicle, typically when driving on paved roads. this sound is usually caused by the vehicles tires interacting with the road surface, and can be a result of various factors',
 'rough engine run issue': 'a rough engine run is when the engine sounds or feels uneven, unsteady, or jerky when its running, often accompanied by vibrations or stumbling. this can be a concerning issue that may indicate a problem with the engines performance, fuel system, or ignition',
 'sensor issue': 'a sensor issue refers to a problem with a vehicles sensor system, which is a network of devices that provide vital information to the vehicles computer ecu to help it run smoothly and safely. sensors can detect various aspects of the vehicle, such as speed, temperature, pressure, and position, and send this information to the ecu, which uses it to make decisions about engine performance, transmission shifting, and other critical functions',
 'service engine soon light issue': 'the service engine soon light, also known as the check engine light, is a warning indicator on your dashboard that signals to you that theres a problem with your vehicles engine. its usually a yellow or red light that appears on the dashboard and can be accompanied by a message or code on the screen',
 'service shifter warning light': 'a service shifter warning light is a warning indicator on your vehicles dashboard that signals a problem with the transmission or the shift linkage. its usually a yellow or orange light that appears when the transmission is not functioning correctly',
 'service warning issue': 'a service warning light is a message that appears on your dashboard, usually in the shape of a car with a warning symbol, indicating that a maintenance or inspection is required. this light is designed to alert you to potential problems with your vehicles systems or components that need attention to prevent further damage or ensure your safety on the road',
 'software issue': 'a software issue in a vehicle refers to a problem that occurs when the vehicles computer system, also known as the engine control unit ecu, malfunctions or becomes corrupted. this can cause various problems with the vehicles systems, such as the engine, transmission, and other electronic components',
 'speed fluctuation issue': 'a speed fluctuation issue refers to a problem where the vehicles speedometer or tachometer shows inconsistent or erratic readings, often resulting in a variation in the vehicles actual speed. this can be frustrating and may affect the vehicles performance, fuel efficiency, and overall safety',
 'steady speed thud issue': 'a steady speed thud issue refers to a phenomenon where a vehicles engine or transmission makes a consistent, dull thudding or clunking noise while driving at a steady speed, usually between 30 60 mph. this noise is often described as a low pitched rumble or vibration that can be felt through the vehicles seat, floor, or steering whee',
 'sudden lurch issue': 'a sudden lurch is a sudden, unexpected movement or jolt in a vehicle, often accompanied by a loss of traction or a feeling of the vehicle pulling or jerking forward. this can be unsettling and may cause the vehicle to stall or lose control',
 'suspension weight handling issue': 'a suspension weight handling issue occurs when a vehicles suspension system is not able to handle the weight of the vehicle or its cargo properly, leading to uneven tire wear, vibrations, and potentially even damage to the vehicles frame or other components',
 'temperature control issue': 'a temperature control issue refers to a problem with the vehicles heating or cooling system that prevents the vehicle from maintaining a comfortable temperature inside the cabin. this can include issues with the air conditioning, heating, or both',
 'throttle hesitation issue': 'throttle hesitation is a phenomenon where the engines throttle response is delayed or uneven, causing the vehicle to hesitate or stumble when accelerating. this can make the vehicle feel sluggish, unresponsive, or even stall',
 'tire vibration issue': 'a tire vibration issue occurs when a vehicles tires vibrate excessively, often causing a shaking or wobbling feeling while driving. this can be a concerning problem that affects the overall comfort, safety, and handling of the vehicle',
 'transmission breakdown issue': 'a transmission breakdown occurs when the transmission in your vehicle fails to properly transfer power from the engine to the wheels, causing the vehicle to lose power, jerk, or stall. as consequence the transmission is replaced',
 'transmission clunking issue': 'a transmission clunking issue refers to a strange noise that occurs when you shift gears, often accompanied by a clunking, clattering, or grinding sound. this noise can be loud and persistent, and can be unsettling for drivers',
 'transmission code issue': 'a transmission code issue refers to a problem with the transmission system of a vehicle, which is responsible for transmitting power from the engine to the wheels. the transmission system consists of a complex network of gears, shafts, and sensors that work together to ensure smooth and efficient power transfer',
 'transmission dial placement issue': 'a transmission dial placement issue occurs when the gear selector lever or dial on your vehicles transmission is not in the correct position, which can cause the transmission to malfunction or not engage properly',
 'transmission downshifting issue': 'a transmission downshifting issue occurs when the transmission is shifting down to a lower gear than expected, often resulting in a jerky or abrupt change in speed, and sometimes even stalling the vehicle',
 'transmission failure': 'a transmission failure occurs when the transmission in a vehicle is no longer able to properly transfer power from the engine to the wheels, causing the vehicle to stall, jerk, or hesitate when shifting gears',
 'transmission fluid leak issue': 'a transmission fluid leak is when the transmission fluid that lubricates and cools the transmission in your vehicle escapes through a crack or hole in the transmission, causing damage to the transmission and potentially leading to costly repairs',
 'transmission fluid level issue': 'a transmission fluid level issue occurs when the transmission fluid in your vehicles transmission system becomes low or dirty, which can cause problems with the transmissions performance and longevity',
 'transmission gear hunting issue': 'a transmission gear hunting issue occurs when the transmission is slipping or hesitating between gears, making it difficult to smoothly shift between gears. this can cause the vehicle to jerk, hesitate, or feel like its hunting for the correct gear',
 'transmission issue': 'a transmission problem occurs when the transmission in your vehicle is not functioning properly, causing issues with the way the vehicle moves. the transmission is responsible for transferring power from the engine to the wheels, allowing the vehicle to move',
 'transmission mode switch delay issue': 'a transmission mode switch delay issue occurs when the transmissions mode switch, which controls the gear shifting, takes longer than usual to switch between gears. this can cause the transmission to hesitate or jerk when shifting, making it uncomfortable to drive',
 'transmission module issue': 'a transmission module is a computerized system that controls the transmissions operation, including shifting gears and adjusting the transmissions fluid pressure. the transmission module receives input from various sensors and sends signals to the transmission to perform its functions',
 'transmission noise issue': 'a transmission noise issue occurs when the transmission in your vehicle makes unusual sounds, such as grinding, clunking, whining, or whirring noises, while shifting gears or idling. these noises can be annoying and may indicate a problem with the transmissions internal components',
 'transmission odor issue': 'a transmission odor issue refers to a strong, unpleasant smell coming from the transmission of a vehicle. this smell can be caused by various factors, including worn out or damaged transmission components, fluid leaks, or contamination',
 'transmission replacement issue': 'a transmission replacement issue occurs when the transmission in a vehicle fails or becomes damaged, requiring a new transmission to be installed to restore proper vehicle function',
 'transmission rollback issue': 'a transmission rollback issue occurs when the transmission slips or hesitates when shifting gears, causing the vehicle to jerk or jerk back into the previous gear instead of smoothly transitioning into the next one. this can be a frustrating and potentially damaging problem if not addressed promptly',
 'transmission roughness issue': 'a transmission roughness issue refers to a problem where the transmission is making unusual noises, vibrations, or slipping when shifting gears. this can be caused by a variety of factors, leading to an uncomfortable driving experience',
 'transmission shift pattern issue': 'a transmission shift pattern issue occurs when the transmission is not shifting gears smoothly or at the correct times, resulting in an abnormal or jerky shifting pattern. this can cause the vehicle to hesitate, jerk, or grind when shifting gears, making it uncomfortable to drive',
 'transmission slippage issue': 'transmission slippage occurs when the transmission is not engaging gears properly, causing the vehicle to slip or hesitate when shifting gears. this can lead to a decrease in fuel efficiency, decreased performance, and potentially cause damage to the transmission',
 'transmission surge issue': 'a transmission surge occurs when the transmission suddenly surges or jerks, causing the vehicle to jerk or lurch forward, often accompanied by a loud noise. this can be unsettling and may be accompanied by other symptoms such as slipping or hesitation when shifting gears',
 'transmission vibration issue': 'a transmission vibration issue occurs when the transmission in your vehicle vibrates or shakes while driving, which can be uncomfortable and potentially cause damage to the transmission, drivetrain, or other components',
 'transmission warning issue': 'a transmission warning issue occurs when the transmission in your vehicle is not functioning properly, causing the warning lights on your dashboard to illuminate. this can be a frustrating and potentially costly problem if not addressed promptly',
 'unclear iconography issue': 'unclear iconography refers to a problem where the dashboard warning lights or indicators on your vehicles dashboard are not displaying correctly, making it difficult to understand the status of your vehicles systems',
 'uncomfortable handling issue': 'an uncomfortable handling issue refers to a problem with a vehicles ability to respond to steering input, making it difficult to control the vehicle while driving. this can manifest in various ways, such as a loose or wobbly feel when steering, a tendency to pull to one side, or a general lack of responsiveness',
 'unexpected electric mode issue': 'an unexpected electric mode issue occurs when a vehicles electric system malfunctions and displays an incorrect or unexpected mode, such as e for electric mode, ev for electric vehicle mode, or ev range when its not supposed to be in electric mode. this can be confusing and may cause inconvenience to the driver',
 'unexpected engine shutdown issue': 'an unexpected engine shutdown occurs when the engine suddenly stops running, often without warning, while the vehicle is in motion or stationary. this can be a frustrating and potentially hazardous situation, especially if it happens while driving',
 'unexpected vehicle shutdown issue': 'an unexpected vehicle shutdown occurs when your car suddenly stops running or shuts down while its in motion or at a standstill, without warning. this can be a frustrating and potentially hazardous situation',
 'unintended acceleration issue': 'unintended acceleration is a phenomenon where a vehicle accelerates unexpectedly, often without any input from the driver. this can be a frightening and potentially hazardous situation, as it can cause the vehicle to speed up rapidly, making it difficult to control',
 'unintended door movement issue': 'an unintended door movement issue occurs when a car door moves or swings open or closed on its own, without being touched or intentionally opened by the driver or passenger',
 'unintended vehicle movement issue': 'an unintended vehicle movement issue, also known as a rollaway or drift, occurs when a vehicle moves on its own without the drivers input or control. this can happen when the vehicle is parked, stopped, or even when the driver is not paying attention',
 'unpleasant engine sound issue': 'an unpleasant engine sound issue refers to a strange, unusual, or disturbing noise coming from the engine of a vehicle. this noise can be loud, high pitched, low pitched, or even a combination of different sounds, and can be heard when the engine is running, idling, or under load',
 'unpleasant odor issue': 'an unpleasant odor in a vehicle is a common issue that can be caused by various factors. its not just a pleasant smell, but rather a sign of a potential problem that needs attention',
 'unpleasant sound issue': 'an unpleasant sound issue refers to a noise that is unusual, annoying, or disturbing coming from a vehicle. this can include a variety of sounds such as grinding, screeching, clunking, hissing, or rattling noises',
 'unremarkable engine sound issue': 'an unremarkable engine sound issue refers to a situation where the engine makes a normal, non abnormal sound, but its not the usual vroom or rumble youd expect from a running engine. this can be a bit confusing, as its not a clear indication of a specific problem, but rather a lack of a distinct sound',
 'unresponsive a / c control issue': 'an unresponsive air conditioning ac control issue occurs when the air conditioning system in your vehicle fails to function properly, making it difficult or impossible to control the temperature inside the vehicle',
 'unresponsive screen issue': 'an unresponsive screen issue occurs when the dashboard or instrument cluster of a vehicle fails to display information or respond to inputs, such as buttons or gauges. this can be frustrating and may indicate a problem with the vehicles electrical system or computer',
 'unusual acceleration noise issue': 'an unusual acceleration noise is a strange, unexplained sound that occurs when you accelerate your vehicle, often accompanied by a vibration or pulsation. this noise can be loud and unsettling, making it difficult to drive smoothly and confidently',
 'unusual audio noise issue': 'an unusual audio noise issue refers to a strange sound coming from your vehicles speakers, dashboard, or other audio components. this noise can be a distraction while driving and may be caused by a variety of factors',
 'unusual body sound issue': 'an unusual body sound issue refers to a strange noise coming from the exterior of a vehicle, typically from the body or suspension system. this noise can be a vibration, rattling, clunking, or scraping sound that is not typical of normal vehicle operation',
 'unusual drive dynamic issue': 'an unusual drive dynamic issue refers to a problem with the way a vehicle handles or responds to driver input while in motion. this can manifest in various ways, such as vibrations, pulling to one side, or an unbalanced feel while driving. its like the vehicle is feeling off or uncooperative when youre behind the wheel',
 'unusual engine sound issue': 'an unusual engine sound is a noise that is not typical of a normal engine operation. this can include strange vibrations, rattling, clunking, or other unusual sounds that can be unsettling and may indicate a problem with the engine',
 'unusual engine startup noise issue': 'an unusual engine startup noise refers to a strange or abnormal sound that occurs when you start your vehicle. this noise can be a clicking, clunking, grinding, or rattling sound that is different from the typical vroom or hum of the engine',
 'unusual exhaust sound issue': 'an unusual exhaust sound is a noise that is different from the normal, smooth sound of a vehicles exhaust system. this sound can be a series of unusual noises, such as rattling, clunking, hissing, or a loud, piercing whine',
 'unusual loud noise issue': 'an unusual loud noise issue refers to a sound that is louder than usual and unexpected, often coming from the engine, transmission, or other parts of the vehicle. this noise can be unsettling and may indicate a problem that needs attention',
 'unusual noise after drive issue': 'an unusual noise after drive issue refers to a strange sound that occurs when a vehicle is driven, often after a period of inactivity or after the engine has been turned off. this noise can be a vibration, rattling, clunking, or grinding sound that is not typical of the vehicles normal operation',
 'unusual noise at high speed issue': 'an unusual noise that occurs when driving at high speeds, often accompanied by vibrations or pulsations, can be a concerning issue for vehicle owners. this noise can be caused by various factors, and its essential to identify the root cause to ensure the safety and longevity of the vehicle',
 'unusual noise during startup and idle issue': 'an unusual noise that occurs when a vehicle is starting up or idling, which can be a concerning symptom that may indicate a problem with the vehicles engine, transmission, or other components',
 'unusual noise issue': 'an unusual noise issue refers to a sound that is not typical of your vehicles normal operation. this noise can be a clicking, clunking, grinding, or screeching sound that is not usually heard when driving or idling. it can be a bit unsettling, but in most cases, its not a sign of a major problem',
 'unusual noise on cold start issue': 'an unusual noise on a cold start refers to a strange sound that occurs when you turn the key to start the engine, especially when the vehicle is cold. this noise can be a clicking, clunking, grinding, or rattling sound that is not typical of a normal engine start',
 'unusual odor issue': 'an unusual odor in a vehicle refers to a strong, unpleasant smell that is not typical of the vehicles normal scent. this can be a pungent, sweet, sour, or musty smell that can be detected inside the vehicle, often in the interior or from the engine compartment',
 'unusual start noise issue': 'an unusual start noise is a sound that occurs when you turn the key to start your vehicle, but the engine doesnt roar to life as expected. instead, you may hear a strange, unusual, or unfamiliar noise, such as grinding, clicking, whining, or clunking sounds',
 'unwanted app notification issue': 'an unwanted app notification is a problem where a mobile app sends you notifications that you didnt request or didnt ask for. these notifications can be annoying and disrupt your daily activities',
 'vehicle breakdown issue': 'a vehicle breakdown issue means the vehicle cant operate due to a major fault like battery failure, engine trouble, or transmission issues. immediate repair is needed to restore function and avoid further damage',
 'vehicle instability issue': 'vehicle instability refers to a condition where a vehicle feels unbalanced, wobbly, or unpredictable, making it difficult to control or steer. this can be a safety concern, as it can increase the risk of an accident',
 'vehicle oscillation issue': 'a vehicle oscillation issue occurs when a vehicles wheels or suspension system vibrates or oscillates excessively, causing the vehicle to wobble or shake while driving. this can be a concerning problem that affects the vehicles stability, comfort, and overall safety',
 'vehicle electrical power loss issue': 'a vehicle electrical power loss issue occurs when a vehicles engine or electrical system is not producing enough power to run the vehicles accessories, such as the lights, radio, and other electrical systems, or the engine is not running at all',
 'vehicle shut down issue': 'a vehicle shut down issue occurs when a car suddenly stops running or shuts down while its in motion or at a standstill. this can be a frustrating and potentially hazardous situation, especially if it happens while driving',
 'vehicle stall issue': 'a vehicle stall occurs when the engine suddenly stops running, often accompanied by a loss of power and a failure to move forward. this can be a frustrating and potentially hazardous situation, especially on the highway or in heavy traffic',
 'vehicle start issue': 'a vehicle start issue refers to a problem that prevents a car from starting or makes it difficult to start. this can be frustrating and may require professional assistance to resolve',
 'vehicle vibration issue': 'a vehicle vibration issue occurs when the vehicle shakes, wobbles, or vibrates while in motion, making it uncomfortable and potentially affecting the overall driving experience',
 'water pump replacement issue': 'a water pump replacement issue occurs when the engines water pump fails to circulate coolant properly, leading to overheating and potentially causing damage to the engine',
 'wheel vibration issue': 'a wheel vibration issue occurs when a vehicles wheels vibrate or shake while in motion, which can be uncomfortable and potentially cause damage to the vehicles suspension, steering, and other components',
 'whole body vibration issue': 'a whole body vibration issue occurs when a vehicle shakes, rattles, or vibrates excessively, affecting the entire vehicle, including the seats, steering wheel, and even the occupants. this can be a concerning problem that can lead to a range of issues, from minor annoyances to safety concerns',
 'window fog issue': 'a window fog issue occurs when the glass in your vehicles windows becomes hazy or foggy, making it difficult to see out of the window. this can be caused by a variety of factors, including temperature changes, humidity, and poor maintenance',
 'not relevant': 'this topics refers to any coment that can not allow to understand the reason of the complain'}]
 
Strict classification instructions:  

1. Carefully read the entire complaint text (verbatim).
   - Understand the overall meaning ONLY from the explicit content.
   - Do NOT add, interpolate, or use external knowledge.

2. Identify ONLY symptoms and problems explicitly mentioned in the verbatim.

3. Select either:
   - ONE single, precise technical issues from the list of authorized categories best matching the described problem

4. If two problems belong to the same technical system (category), select ONLY ONE topic, the most representative.

5. NEVER make assumptions, deductions, or extrapolations beyond the explicit symptoms.

6. If the complaint is vague, irrelevant, or lacks clear symptoms matching authorized categories, respond ONLY with:
   Feature_AI : not relevant , Trust_score : 10

7. NEVER invent or add categories outside the authorized list.

8. Avoid generic categories like "drivetrain issue" or "transmission issue" when a more specific category is clearly expressed.

Output format (strict):

Feature_AI : <best topic>  
Trust_score : <score>  
Symptoms_mentioned : <list of explicitly mentioned symptoms>
Explanation :  
- Instruction 1 applied: complaint fully read, understood solely from verbatim (<brief symptom summary>)  
- Instruction 2 applied: selected <one or two> topic(s) because <reason based on explicit symptoms>  
- Instruction 3 applied: no assumptions beyond explicit symptoms  
- Instruction 4 applied: complaint judged <relevant or not relevant> based on symptoms  
- Instruction 5 applied: no invented categories, only authorized ones used  
- Link between symptoms and topic(s): <concise explanation of how symptoms relate to chosen topic(s)>


User complaint: "{verbatim}"
    """
    prompt = prompt_template.replace("{verbatim}", verbatim)
    return call_llm(prompt)



# COMMAND ----------

def parse_classification_output(text):
    result = {
        "Feature_AI": None,
        "Trust_score": None,
        "Symptoms_mentioned": None,
        "Explanation": None
    }
    if not isinstance(text, str):
        return result

    feature_match = re.search(r"Feature_AI\s*:\s*([^\n\r]+)", text)
    trust_match = re.search(r"Trust_score\s*:\s*([^\n\r]+)", text)
    symptoms_match = re.search(r"Symptoms_mentioned\s*:\s*([^\n\r]+)", text)
    explanation_match = re.search(r"Explanation\s*:\s*((?:.|\n)+)", text)

    if feature_match:
        result["Feature_AI"] = feature_match.group(1).strip()
    if trust_match:
        score_str = trust_match.group(1).strip()
        try:
            if '.' in score_str:
                result["Trust_score"] = float(score_str)
            else:
                result["Trust_score"] = int(score_str)
        except:
            result["Trust_score"] = score_str
    if symptoms_match:
        result["Symptoms_mentioned"] = symptoms_match.group(1).strip()
    if explanation_match:
        explanation = explanation_match.group(1).strip()
        explanation = re.sub(r"\n+", " ", explanation)
        explanation = re.sub(r"\s{2,}", " ", explanation)
        result["Explanation"] = explanation

    if result["Feature_AI"] is None:
        feature_match_md = re.search(r"\*\*Feature_AI\*\*\s*:\s*(.+)", text)
        trust_match_md = re.search(r"- \*\*Trust_score\*\*\s*:\s*(.+)", text)
        symptoms_match_md = re.search(r"- \*\*Symptoms_mentioned\*\*\s*:\s*(.+)", text)
        explanation_match_md = re.search(r"- \*\*Explanation\*\*\s*:\s*((?:.|\n)+)", text)

        if feature_match_md:
            result["Feature_AI"] = feature_match_md.group(1).strip()
        if trust_match_md:
            score_str = trust_match_md.group(1).strip()
            try:
                if '.' in score_str:
                    result["Trust_score"] = float(score_str)
                else:
                    result["Trust_score"] = int(score_str)
            except:
                result["Trust_score"] = score_str
        if symptoms_match_md:
            result["Symptoms_mentioned"] = symptoms_match_md.group(1).strip()
        if explanation_match_md:
            explanation = explanation_match_md.group(1).strip()
            explanation = re.sub(r"\n+", " ", explanation)
            explanation = re.sub(r"\s{2,}", " ", explanation)
            result["Explanation"] = explanation

    return result


