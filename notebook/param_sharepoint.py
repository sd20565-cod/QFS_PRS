# Databricks notebook source
# MAGIC %run "/Workspace/Users/oumaima.idaline@external.stellantis.com/qfs_classification/qfs_sharepoint/qfs_Functions
"

# COMMAND ----------

#import toolbox_connected_vehicle as tcv
import pyspark.sql.functions as F
import datetime
from pyspark.sql import Window
from pyspark.sql.window import Window
from pyspark.sql import Row
from pyspark.sql.functions import *

# COMMAND ----------

scope_name = ''
token = ''
create_secret_scope(scope_name, token)

# COMMAND ----------

scope_name = ""
parameters = {
     "Application Name": "",
     "tenant_ID": "",
     "Secret_Value": "",
     "client_id": "",
       
}
databricks_url = "" 
store_secrets_in_databricks(scope_name, parameters, databricks_url, token)

# COMMAND ----------

scope_name = ""
tech_user_name = 's'
tech_user_password = ''
store_secret_for_user(scope_name, tech_user_name, tech_user_password, databricks_url, token)


# COMMAND ----------

scope_name = ''  
secret_key = '' 
secret_value = '' 
store_key_value_in_databricks(databricks_url, token, scope_name, secret_key, secret_value)

# COMMAND ----------

databricks_url = '' 
# token = '' 
# scope_name = ''  
# secret_key = '' 
# secret_value = '' 
# store_key_value_in_databricks(databricks_url, token, scope_name, secret_key, secret_value)