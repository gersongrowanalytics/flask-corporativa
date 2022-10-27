
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

@app.route('/traerDatos', methods = ['POST'])
def traer_datos():
    try:
        
        path = "./"
        lista_archivos = os.listdir(path)
        for archivo in lista_archivos:
            if archivo.endswith(".xlsx"):
                os.remove(path + archivo)

        cursor = conexion.connection.cursor()

        meses = ['meses','Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Setiembre','Octubre','Noviembre','Diciembre']

        #Parametros enviados por JSON, si se envia por fomulario...=> request.form.get()
        req_fecha   = request.json.get('fecha') + '%'
        req_array_empresas   = request.json.get('empresas')
        req_array_zonas   = request.json.get('zonas')

        suf_nombre  = req_fecha[4:6]
        req_anio    = req_fecha[2:4]

        mes_index   = suf_nombre

        nombre_mes  = meses[int(mes_index)]

        #nombre archivo
        nombre_archivo  = 'Ventas SO ({0} {1})'.format(nombre_mes[:3], req_anio)

        #nombre archivo con extension
        nombre_archivo_ext = 'Ventas SO ({0} {1})-{2}.xlsx'.format(nombre_mes[:3], req_anio, str(random.randint(100,999)))

        #Eliminar el registro si ya existe el mes y año 
        # consulta_eliminar = 'DELETE FROM carcargasarchivos WHERE carurl = "{0}"'.format(nombre_archivo)
        # cursor.execute(consulta_eliminar)
        # conexion.connection.commit()
        # print(cursor.rowcount, " registro eliminado: {0}".format(nombre_archivo))

        if req_array_zonas == []:

            tamanio_empresas = len(req_array_empresas)
            consulta_web = ''

            cont = 1
            for empresa in req_array_empresas:
                if cont == tamanio_empresas:
                    consulta_web = consulta_web + 'vsbempresa = "{0}" '.format(empresa)
                else:
                    consulta_web = consulta_web + 'vsbempresa = "{0}" OR '.format(empresa)
                cont = cont + 1

            consulta_seleccionar = 'SELECT * FROM vsbventassobol WHERE (vsbfecha LIKE "{0}") && ({1}) ORDER BY vsbempresa, vsbtotalreventa DESC'.format(req_fecha, consulta_web)

        else:

            tamanio_zonas = len(req_array_zonas)
            consulta_web = ''

            cont = 1
            for zona in req_array_zonas:
                if cont == tamanio_zonas:
                    consulta_web = consulta_web + 'vsbregion = "{0}"'.format(zona)
                else:
                    consulta_web = consulta_web + 'vsbregion = "{0}" OR '.format(zona)
                cont = cont + 1

            consulta_seleccionar = 'SELECT * FROM vsbventassobol WHERE (vsbfecha LIKE "{0}") && ({1}) ORDER BY vsbfecha, vsbregion, vsbempresa, vsbtotalreventa DESC'.format(req_fecha, consulta_web)

        # print(consulta_seleccionar)

        cursor.execute(consulta_seleccionar)
        datos = cursor.fetchall()
        
        return os.getcwd()

        #Creacion y manipulación del excel
        archivoXls = xlsxwriter.Workbook('./'+nombre_archivo_ext)
        worksheet = archivoXls.add_worksheet('Ventas SO')
        archivoXls.close()
    
        PATH = './' + nombre_archivo_ext

        

    except Exception as e:
        return e

# @app.route('/descargar-archivo/<archivo>')
def descargar_archivo(archivo):

    PATH = '../'+archivo
    return send_file(PATH, as_attachment = True)

if __name__ == '__main__':
    app.config.from_object(config['development'])
    app.run()