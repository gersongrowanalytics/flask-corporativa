
from flask import Flask, jsonify, send_file, request
from config import config
from flask_mysqldb import MySQL
# from flask_cors import CORS
import os
import xlsxwriter
import random

app = Flask(__name__)
# cors = CORS(app)
# app.config['CORS_HEADERS'] = 'Content-Type'

conexion = MySQL(app)

@app.route('/traerDatos')
def traer_datos():
    return jsonify({
        "data" : "datas",
        "otros" : os.getcwd()
    })

# @app.route('/descargar-archivo/<archivo>')
def descargar_archivo(archivo):

    PATH = '../'+archivo
    return send_file(PATH, as_attachment = True)

if __name__ == '__main__':
    app.config.from_object(config['development'])
    app.run()