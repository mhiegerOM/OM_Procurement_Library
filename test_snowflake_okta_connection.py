import snowflake.connector
import pandas as pd
import yaml

# Load credentials from YAML file
with open("creds.yaml", "r") as yamlfile:
    creds = yaml.safe_load(yamlfile)

# Snowflake connection parameters
account = creds["snowflake"]["account"]
warehouse = creds["snowflake"]["warehouse"]
database = creds["snowflake"]["database"]
schema = creds["snowflake"]["schema"]
user = creds["snowflake"]["user"]
role = creds["snowflake"]["role"]
session_parameters = creds["snowflake"].get("session_parameters", None)

# Establish connection to Snowflake using Okta
conn = snowflake.connector.connect(
    account=account,
    warehouse=warehouse,
    database=database,
    schema=schema,
    user=user,
    role=role,
    session_parameters=session_parameters,
    authenticator="externalbrowser"
)

# Create a cursor object
cur = conn.cursor()

# Path to your CSV file
csv_file_path = 'C:/Users/MatthewHieger/Documents/My Tableau Repository/Datasources/Coupa Datamart/DATAMART COMPILE TEST/Invoice DMtest Source.csv'

# Read CSV into DataFrame
df = pd.read_csv(csv_file_path)

# Create table name for Snowflake
table_name = 'mhieger_test_1'

# Convert DataFrame to Snowflake table
df.to_sql(name=table_name, con=conn, schema=schema, if_exists='replace', index=False)

# Commit changes and close connection
conn.commit()
conn.close()

print("DataFrame uploaded to Snowflake successfully!")
