"""
One-time setup: creates dbo.FIRS_USERS and seeds the Administrator account.
Run once: python create_users_table.py
"""
import pymssql, json
from werkzeug.security import generate_password_hash

with open('config.json') as f:
    cfg = json.load(f)

db = cfg['database']
conn = pymssql.connect(server=db['server'], database=db['name'],
                       user=db['user'], password=db['password'])
cur = conn.cursor()

cur.execute("""
    IF NOT EXISTS (
        SELECT 1 FROM INFORMATION_SCHEMA.TABLES
        WHERE TABLE_SCHEMA='dbo' AND TABLE_NAME='FIRS_USERS'
    )
    CREATE TABLE dbo.FIRS_USERS (
        USER_ID       INT IDENTITY(1,1) PRIMARY KEY,
        EMAIL         NVARCHAR(150) NOT NULL,
        FULL_NAME     NVARCHAR(200) NOT NULL DEFAULT '',
        PASSWORD_HASH NVARCHAR(512) NOT NULL,
        ROLE          NVARCHAR(50)  NOT NULL DEFAULT 'Viewer',
        IS_ACTIVE     BIT           NOT NULL DEFAULT 1,
        CREATED_AT    DATETIME      NOT NULL DEFAULT GETDATE(),
        LAST_LOGIN    DATETIME      NULL,
        CREATED_BY    NVARCHAR(150) NULL,
        CONSTRAINT UQ_FIRS_USERS_EMAIL UNIQUE (EMAIL)
    )
""")
conn.commit()
print("Table dbo.FIRS_USERS: ready.")

cur.execute("SELECT COUNT(*) FROM dbo.FIRS_USERS WHERE EMAIL = 'Administrator'")
if cur.fetchone()[0] == 0:
    ph = generate_password_hash('CRYPTware1')
    cur.execute("""
        INSERT INTO dbo.FIRS_USERS (EMAIL, FULL_NAME, PASSWORD_HASH, ROLE, IS_ACTIVE, CREATED_BY)
        VALUES (%s, %s, %s, %s, 1, 'system')
    """, ('Administrator', 'System Administrator', ph, 'Administrator'))
    conn.commit()
    print("Admin user 'Administrator' created with password 'CRYPTware1'.")
else:
    print("Admin user already exists — not modified.")

conn.close()
print("Done.")
