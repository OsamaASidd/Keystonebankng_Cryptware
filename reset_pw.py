import pymssql, json
from werkzeug.security import generate_password_hash, check_password_hash

with open('e:/EINVOICING_AGENT/config.json') as f:
    db = json.load(f)['database']

conn = pymssql.connect(server=db['server'], database=db['name'],
                       user=db['user'], password=db['password'], timeout=30)
cur = conn.cursor()

new_pw = 'CRYPTware20$'
new_hash = generate_password_hash(new_pw)

# Show current state
cur.execute("SELECT EMAIL, FULL_NAME, IS_ACTIVE, ROLE FROM dbo.FIRS_USERS WHERE EMAIL='Administrator'")
row = cur.fetchone()
if row:
    print(f"Found: EMAIL={row[0]}, NAME={row[1]}, IS_ACTIVE={row[2]}, ROLE={row[3]}")
else:
    print("ERROR: No user with EMAIL='Administrator' found!")
    # List all users
    cur.execute("SELECT EMAIL, FULL_NAME, IS_ACTIVE, ROLE FROM dbo.FIRS_USERS")
    for r in cur.fetchall():
        print(f"  -> {r}")
    conn.close()
    exit(1)

# Reset password and ensure account is active
cur.execute("""
    UPDATE dbo.FIRS_USERS
    SET PASSWORD_HASH=%s, IS_ACTIVE=1
    WHERE EMAIL='Administrator'
""", (new_hash,))
conn.commit()

# Verify
cur.execute("SELECT PASSWORD_HASH FROM dbo.FIRS_USERS WHERE EMAIL='Administrator'")
stored = cur.fetchone()[0]
ok = check_password_hash(stored, new_pw)
print(f"Password set to: {new_pw}")
print(f"Hash verification: {'PASS' if ok else 'FAIL'}")
conn.close()
