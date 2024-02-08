import pandas as pd
import mysql.connector
import psycopg2
import cx_Oracle

def consulta():

    query = input('Informe sua consulta SQL: ')

    df = pd.read_sql(query, conn)

    name = input('Informe o nome do arquivo (não é necessário informar a extensão): ')
    caminho = input('Informe o caminho do arquivo (opcional): ')

    if caminho != '':
        if caminho[-1] != '/':    
            caminho = f"{caminho}/"
            df.to_excel(f'{caminho}{name}.xlsx', index=False)
    else:
        df.to_excel(f'{name}.xlsx', index=False)

print('MySQL ( 1 )')
print('Postgres ( 2 )')
print('Oracle ( 3 )')
banco = input('Informe seu banco de dados: ')

database = input('Informe o nome do seu banco: ')
host = input('Informe o host do seu banco: ')
user = input('Informe o usuario do seu banco: ')
password = input('Informe a senha do seu banco: ')

if banco == '1' or banco.lower() == 'mysql':
    
    conn = mysql.connector.connect(
        host=host,
        user=user,
        password=password,
        database=database
    )

elif banco == '2' or banco.lower() == 'postgres':
    port = input('Informe a porta do seu banco: ')
    db_config = {
        'host': host,
        'port': port,
        'database': database,
        'user': user,
        'password': password
    }
    conn = psycopg2.connect(**db_config)

elif banco == '3' or banco.lower() == 'oracle':
    conn = cx_Oracle.connect(f'{user}/{password}@{database}')

consulta()

while True:
    print('Tabela gerada com sucesso!')
    continuar = input('Deseja gerar outra tabela? (S/N): ')
    if continuar.lower() == 's':
        consulta()
    else:
        break

conn.close()