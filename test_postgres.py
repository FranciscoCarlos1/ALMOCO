import os
from psycopg import connect

url = os.getenv("DATABASE_URL")

conn = connect(url)
cur = conn.cursor()

cur.execute("SELECT 1;")
print(cur.fetchone())

conn.close()