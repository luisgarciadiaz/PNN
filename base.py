import mysql.connector

def conexion():
    mydb=mysql.connector.connect(
    host="172.30.27.40",
    user="root",
    passwd="",
    database="base_cordi"
    )
    return mydb

