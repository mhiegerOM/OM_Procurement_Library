# import snowflake.connector
# import pandas as pd
# import yaml

# def read_yaml(file_path):
#     with open(file_path, 'r') as file:
#         return yaml.safe_load(file)

# # Load credentials from YAML file
# credentials = read_yaml('creds.yaml')
# username = credentials['snowflakeusername']
# password = credentials['snowflakepassword']

# # Connect to Snowflake
# conn = snowflake.connector.connect(
#     username=username,
#     password=password,
#     account='ADS_ROLE',
#     warehouse='TABLEAU_WH',
#     database='OM_DB_ADS',
#     schema='ADS_SCHEMA'
# )

# # Fetch data from source (e.g., Coupa API)
# # Transform data if necessary using pandas

# # Write data to Snowflake
# cursor = conn.cursor()
# cursor.execute("""
#     INSERT INTO your_table (column1, column2, ...)
#     VALUES (%s, %s, ...)
# """, your_data)
# conn.commit()

# # Close connection
# conn.close()