import sqlite3
conn = sqlite3.connect('db.db', check_same_thread=False)
cur = conn.cursor()
cur.execute('CREATE TABLE IF NOT EXISTS minha_tabela (pkt TEXT PRIMARY KEY)')
conn.commit()
cur.close()
conn.close()


def set(key, value):
    conn = sqlite3.connect('db.db', check_same_thread=False)
    cur = conn.cursor()
    cur.execute("PRAGMA table_info(minha_tabela)")
    if key not in [column[1] for column in cur.fetchall()]:
        cur.execute(f"ALTER TABLE minha_tabela ADD COLUMN {key} TEXT")
        cur.execute(f"INSERT INTO minha_tabela ({key}) VALUES (?)", (value,))
    else:
        cur.execute(f"UPDATE minha_tabela SET {key} = ?", (value,))
    conn.commit()
    cur.close()
    conn.close()

def get(key):
    conn = sqlite3.connect('db.db', check_same_thread=False)
    cur = conn.cursor()
    cur.execute(f'SELECT {key} FROM minha_tabela LIMIT 1')
    resultado = cur.fetchone()
    valor = resultado[0] if resultado is not None else None
    cur.close()
    conn.close()
    return valor


#set('recontatos', "['14981577499', '14998357508', '14991381717', '14991040872', '14981264043', '14996477058', '14996113326']")
#set('ci', "['Paula', 'Rosana', 'Carol', 'Generozo', 'Joao', 'Luiza']")
#set('last', '3074')
#set('lastduplicate', '3074')
print(get('lastduplicate'))
print(get('last'))


print(get('recontatos'))
print(get('ci'))