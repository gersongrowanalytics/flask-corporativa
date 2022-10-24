
from flask import Flask, jsonify, send_file, request
from config import config
from flask_mysqldb import MySQL
from flask_cors import CORS
import os
import xlsxwriter
import random

app = Flask(__name__)
cors = CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'

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
        consulta_eliminar = 'DELETE FROM carcargasarchivos WHERE carurl = "{0}"'.format(nombre_archivo)
        cursor.execute(consulta_eliminar)
        conexion.connection.commit()
        print(cursor.rowcount, " registro eliminado: {0}".format(nombre_archivo))

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

        print(consulta_seleccionar)

        cursor.execute(consulta_seleccionar)
        datos = cursor.fetchall()

        #Creacion y manipulación del excel
        archivoXls = xlsxwriter.Workbook(nombre_archivo_ext)
        worksheet = archivoXls.add_worksheet('Ventas SO')
        
        celda_empresa   = archivoXls.add_format({
            'font_color':'white',
            'border_color':'#666666',
            'border':2,
            'font_name':'Segoe UI',
            'font_size':9,
            'bg_color': '#F4B084'
        })
        
        celda_cliente   = archivoXls.add_format({
            'font_color':'white',
            'border_color':'#666666',
            'border':2,
            'font_name':'Segoe UI',
            'font_size':9,
            'bg_color': '#FFD966'
        })
        
        celda_ventas    = archivoXls.add_format({
            'font_color':'white',
            'border_color':'#666666',
            'border':2,
            'font_name':'Segoe UI', 
            'font_size':9,
            'bg_color': '#292929'
        })

        celda_empresa.set_align('vcenter')
        celda_cliente.set_align('vcenter')
        celda_ventas.set_align('vcenter')

        celda_fecha = archivoXls.add_format({
            'font_color':'black',
            'border_color':'#666666',
            'border':2,
            'font_name':'Calibri',
            'font_size':9,
            'num_format':'d-mmm'
        })

        celda_fecha.set_align('vcenter')

        celda_venta_numero  = archivoXls.add_format({
            'font_color':'black',
            'border_color':'#666666',
            'border':2,
            'font_name':'Calibri',
            'font_size':11,
            'num_format': '#,##0.00'
        })


        worksheet.set_row(0, 24.75)
        worksheet.set_column(0, 2, 16)
        worksheet.set_column(3, 3, 8)
        worksheet.set_column(4, 4, 16)
        worksheet.set_column(5, 5, 16)
        worksheet.set_column(6, 6, 16)
        worksheet.set_column(10, 10, 8)
        worksheet.set_column(11, 14, 16)

        worksheet.write(0, 0, 'Distribuidor', celda_empresa)
        worksheet.write(0, 1, 'Zona', celda_empresa)
        worksheet.write(0, 2, 'Locality', celda_empresa)
        worksheet.write(0, 3, 'Año', celda_empresa)
        worksheet.write(0, 4, 'Mes', celda_empresa)
        worksheet.write(0, 5, 'Dia', celda_empresa)
        worksheet.write(0, 6, 'Cod. Material', celda_empresa)
        worksheet.write(0, 7, 'Material', celda_empresa)
        worksheet.write(0, 8, 'Categoria', celda_empresa)
        worksheet.write(0, 9, 'Sub Categoría', celda_empresa)
        worksheet.write(0, 10, 'Cod Cliente', celda_cliente)
        worksheet.write(0, 11, 'Vendedor', celda_cliente)
        worksheet.write(0, 12, 'Tipo Negocio', celda_cliente)
        worksheet.write(0, 13, 'Zona Cliente', celda_cliente)
        worksheet.write(0, 14, 'Mercado', celda_cliente)
        worksheet.write(0, 15, 'V-MonedaLocal', celda_ventas)
        worksheet.write(0, 16, 'V-NIV', celda_ventas)
        worksheet.write(0, 17, 'V-Cajas', celda_ventas)

        row = 1
        col = 0

        for a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u in datos:

            if g != '-99':

                anio = f[0:4]
                mes = int(f[4:6])

                dia = f[6:]

                worksheet.write(row, col, b if b != 'SIN ASIGNAR' else '')
                worksheet.write(row, col + 1, c if c != 'SIN ASIGNAR' else '')
                worksheet.write(row, col + 2, d if d != 'SIN ASIGNAR' else '')
                worksheet.write(row, col + 3, int(anio))
                worksheet.write(row, col + 4, meses[mes])
                worksheet.write(row, col + 5, int(dia))
                worksheet.write(row, col + 6, g if g != 'SIN ASIGNAR' else '')
                worksheet.write(row, col + 7, h if h != 'SIN ASIGNAR' else '') 
                worksheet.write(row, col + 8, i if i != 'SIN ASIGNAR' else '') 
                worksheet.write(row, col + 9, j if j != 'SIN ASIGNAR' else '') 
                worksheet.write(row, col + 10, k if k != 'SIN ASIGNAR' else '')
                worksheet.write(row, col + 11, l if l != 'SIN ASIGNAR' else '')
                worksheet.write(row, col + 12, n if n != 'SIN ASIGNAR' else '')
                worksheet.write(row, col + 13, o if o != 'SIN ASIGNAR' else '')
                worksheet.write(row, col + 14, p if p != 'SIN ASIGNAR' else '')
                worksheet.write(row, col + 15,float(q), celda_venta_numero)
                worksheet.write(row, col + 16, u, celda_venta_numero)
                worksheet.write(row, col + 17, float(r), celda_venta_numero)
                row = row + 1

        archivoXls.close()

        consulta = 'INSERT INTO carcargasarchivos (carid, tcaid, fecid, usuid, carnombrearchivo, carubicacion, carexito, created_at, updated_at, carurl) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)'

        valores = (None, 17, None, 1, nombre_archivo_ext,'/', 1, None, None, nombre_archivo)

        cursor.execute(consulta, valores)

        conexion.connection.commit()

        PATH = '../' + nombre_archivo_ext

        return send_file(PATH, as_attachment = True)

    except Exception as e:
        return e

@app.route('/descargar-archivo/<archivo>')
def descargar_archivo(archivo):

    PATH = '../'+archivo
    return send_file(PATH, as_attachment = True)

if __name__ == '__main__':
    app.config.from_object(config['development'])
    app.run()