import mysql.connector

def conexion():
    mydb=mysql.connector.connect(
    host="localhost",
    user="root",
    passwd="",
    database="base_cordi"
    )
    return mydb

