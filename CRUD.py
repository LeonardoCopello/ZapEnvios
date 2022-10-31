import mysql.connector

conexao = mysql.connector.connect(
    host='localhost',
    user='root',
    password='123456',
    database='bdyoutube'
)
# CRUD
cursor = conexao.cursor()
nome_produto = "todynho"
comando = f'DELETE FROM vendas WHERE nome_produto = "{nome_produto}")'
cursor.execute(comando)
conexao.commit() # edita o banco de dados
cursor.close()
conexao.close()



# CREATE
cursor = conexao.cursor()
nome_produto = "todynho"
valor = 123
comando = f'INSERT INTO vendas (nome_produto, valor) VALUES ("{nome_produto}", {valor})'
cursor.execute(comando)
conexao.commit() # edita o banco de dados
# resultado = cursor.fetchall() # ler o banco de dados
cursor.close()
conexao.close()

# READ
cursor = conexao.cursor()
comando = f'SELECT * FROM vendas'
cursor.execute(comando)
# conexao.commit() # edita o banco de dados
resultado = cursor.fetchall() # ler o banco de dados
print(resultado)
cursor.close()
conexao.close()

# UPDATE
cursor = conexao.cursor()
nome_produto = "todynho"
valor = 6
comando = f"UPDADE vendas SET valor = {valor} WHERE nome_produto = '{nome_produto}'"
cursor.execute(comando)
conexao.commit() # edita o banco de dados
cursor.close()
conexao.close()


# DELETE
cursor = conexao.cursor()
nome_produto = "todynho"
comando = f'DELETE FROM vendas WHERE nome_produto = "{nome_produto}")'
cursor.execute(comando)
conexao.commit() # edita o banco de dados
cursor.close()
conexao.close()



