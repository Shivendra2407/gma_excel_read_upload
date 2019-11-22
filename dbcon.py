from config import  config as conf
import mysql.connector

connection = None


def get_connection():
    try:
        global connection

        config = conf.config
        connection = mysql.connector.connect(**config)
        return connection

    except Exception as e:
        print("Exception-> "+str(e))
        return None



