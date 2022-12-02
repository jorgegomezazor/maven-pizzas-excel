import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xlsxwriter
import openpyxl
from openpyxl.chart import BarChart3D,Reference
def extract(): # Función que extrae los datos
    order_details = pd.read_csv('order_details.csv', encoding='latin1', sep=';') # Leo el archivo order_details.csv
    orders = pd.read_csv('orders.csv', encoding='latin1',sep=';') # Leo el archivo orders.csv
    pizza_types = pd.read_csv('pizza_types.csv',encoding='latin1',sep=',') # Leo el archivo pizza_types.csv
    pizza = pd.read_csv('pizzas.csv',encoding='latin1',sep=',') # Leo el archivo pizza.csv
    return  order_details, orders, pizza_types,pizza
def limpiar_datos(order_details, orders):
    ord = pd.merge(order_details, orders, on='order_id') # Hago un merge de los dataframes order_details y orders
    #quito columna time
    ord = ord.drop(columns=['time'])
    ord = ord.dropna() # Elimino los nulls
    pd.set_option('mode.chained_assignment', None) # Deshabilito el warning SettingWithCopyWarning
    for i in range(1,len(ord['pizza_id'])):
        try:
            m = ord['pizza_id'][i]
            palabra = ''
            for l in range(len(m)):
                if m[l] =='@':
                    palabra += 'a' # Reemplazo las @ por a
                elif m[l] == '0':
                    palabra += 'o' # Reemplazo los 0 por o
                elif m[l] == '-':
                    palabra += '_' # Reemplazo los - por _
                elif m[l] == ' ':
                    palabra += '_' # Reemplazo los espacios por _
                elif m[l] == '3':
                    palabra += 'e' # Reemplazo los 3 por e
                else:
                    palabra += m[l] # Si no es ninguna de las anteriores, la agrego tal cual
            ord['pizza_id'][i] = ord['pizza_id'][i].replace(m,palabra)  # Reemplazo la palabra por la nueva
        except:
            #elimino la fila que no se pudo limpiar
            try:
                ord = ord.drop(i)
            except:
                pass
    order_details_2 = ord[['order_details_id','order_id', 'pizza_id', 'quantity']] # Creo un nuevo dataframe con las columnas que me interesan
    order_details_2.to_csv('order_details_2.csv', index=False) # Guardo el dataframe en un archivo csv
    orders_2 = ord[['order_id','date']] # Creo un nuevo dataframe con las columnas que me interesan
    orders_2.to_csv('orders_2.csv', index=False) # Guardo el dataframe en un archivo csv
    return order_details_2, orders_2
def transform(order_details, orders, pizza_types, pizz):
    pizza = {}
    fechas = []
    for i in range(len(pizza_types)):
        pizza[pizza_types['pizza_type_id'][i]] = pizza_types['ingredients'][i] #guardo los ingredientes en un diccionario
    for fecha in orders['date']:
        try: 
            f = pd.to_datetime(float(fecha)+3600, unit='s') #convierto la fecha a datetime
        except:
            f = pd.to_datetime(fecha) # Convierto la fecha a datetime
        fechas.append(f) #guardo las fechas en una lista
    cant_pedidos = [[] for _ in range(53)] #creo una lista de listas para guardar la cantidad de pedidos por semana
    pedidos = [[] for _ in range(53)] #creo una lista de listas para guardar los pedidos por semana
    for pedido in range(len(fechas)):
        # print(fechas[pedido])
        cant_pedidos[fechas[pedido].week-1].append(pedido+1) #guardo la cantidad de pedidos por semana
    bucle = 0
    for p in range(2,len(order_details['order_details_id'])): 
        try:
            bucle = abs(order_details['quantity'][p])
        except:
            try:
                if order_details['quantity'][p] == 'One' or order_details['quantity'][p] == 'one':
                    bucle = 1
                elif order_details['quantity'][p] == 'Two' or order_details['quantity'][p] == 'two':
                    bucle = 2
            except:
                pass
        try:
            for i in range(bucle):
                pedidos[fechas[abs(order_details['order_id'][p]-1)].week-1].append(order_details['pizza_id'][p]) #guardo los pedidos por semana teniendo en cuenta la cantidad de pizzas
        except:
            pass
    ingredientes_anuales = {}
    diccs = []
    for dic in range(53):
        diccs.append({}) #creo una lista de diccionarios para guardar los ingredientes por semana
    for i in range(len(pizza_types)):
        ingreds = pizza_types['ingredients'][i] #guardo los ingredientes en una variable
        ingreds = ingreds.split(', ') #separo los ingredientes
        for ingrediente in ingreds:
            ingredientes_anuales[ingrediente] = 0
            for i in range(len(diccs)):
                diccs[i][ingrediente] = 0 #guardo los ingredientes en los diccionarios
    for i in range(len(pedidos)):
        for p in pedidos[i]:
            ing = 0
            tamano = 0
            if p[-1] == 's': #guardo el tamaño de la pizza
                ing = 1 #si es s la pizza tiene 1 ingrediente de cada
                tamano = 2 
            elif p[-1] == 'm':
                ing = 2 #si es m la pizza tiene 2 ingredientes de cada
                tamano = 2
            elif p[-1] == 'l':
                if p[-2] == 'x':
                    if p[-3] == 'x':
                        ing = 5 #si es xxl la pizza tiene 5 ingredientes de cada
                        tamano = 4
                    else:
                        ing = 4 #si es xl la pizza tiene 4 ingredientes de cada
                        tamano = 3
                else:
                    ing = 3 #si es l la pizza tiene 3 ingredientes de cada
                    tamano = 2
            ings = pizza[p[:-tamano]].split(', ')
            for ingrediente in ings:
                ingredientes_anuales[ingrediente] += ing #guardo los ingredientes en el diccionario de ingredientes anuales
                diccs[i][ingrediente] += ing #guardo los ingredientes en los diccionarios de ingredientes por semana
    for i in range(len(diccs)):
        for j in diccs[i]:
            diccs[i][j] = int(np.ceil((diccs[i][j] + (ingredientes_anuales[j]/53))/2)) #aplico la predicción
    #pizzas mas pedidas
    pizzas = {}
    for i in range(len(pizz)):
        pizzas[pizz['pizza_id'][i]] = 0 #guardo las pizzas en un diccionario
    for i in range(len(pedidos)):
        for p in pedidos[i]:
            pizzas[p] += 1 #guardo las pizzas en el diccionario
    return diccs, ingredientes_anuales, cant_pedidos, pizzas, pedidos
def load(diccs, ingredientes_anuales, cant_pedidos, pizzas,pedidos,piz):
    ingredientes_anuales =  sorted(ingredientes_anuales.items(), key=lambda x: x[1], reverse=True) #ordeno los ingredientes por cantidad de pedidos
    ing_ans = {}
    for ing in range(len(ingredientes_anuales)):
        ing_ans[ingredientes_anuales[ing][0]] = ingredientes_anuales[ing][1] #guardo los ingredientes anuales en un diccionario
    pizzas = sorted(pizzas.items(), key=lambda x: x[1], reverse=True)
    ingresos = 0
    ingresos_mensuales = {}
    ing_mensual = 0
    for i in range(len(pedidos)):
        for p in pedidos[i]:
            ing_mensual += piz['price'][piz['pizza_id'] == p].values[0] #guardo los ingresos de una semana
        if i == 4 or i == 8 or i == 13 or i == 17 or i == 22 or i == 26 or i == 31 or i == 35 or i == 39 or i == 43 or i == 47 or i == 52: 
            ingresos += ing_mensual #guardo los ingresos de un mes
            if i==4:
                ingresos_mensuales['Enero'] = ing_mensual #guardo los ingresos de enero en un diccionario
            elif i==8:
                ingresos_mensuales['Febrero'] = ing_mensual #guardo los ingresos de febrero en un diccionario
            elif i==13:
                ingresos_mensuales['Marzo'] = ing_mensual #guardo los ingresos de marzo en un diccionario
            elif i==17:
                ingresos_mensuales['Abril'] = ing_mensual #guardo los ingresos de abril en un diccionario
            elif i==22:
                ingresos_mensuales['Mayo'] = ing_mensual #guardo los ingresos de mayo en un diccionario
            elif i==26:
                ingresos_mensuales['Junio'] = ing_mensual #guardo los ingresos de junio en un diccionario
            elif i==31:
                ingresos_mensuales['Julio'] = ing_mensual   #guardo los ingresos de julio en un diccionario
            elif i==35:
                ingresos_mensuales['Agosto'] = ing_mensual #guardo los ingresos de agosto en un diccionario
            elif i==40:
                ingresos_mensuales['Septiembre'] = ing_mensual #guardo los ingresos de septiembre en un diccionario
            elif i==44:
                ingresos_mensuales['Octubre'] = ing_mensual #guardo los ingresos de octubre en un diccionario
            elif i==48:
                ingresos_mensuales['Noviembre'] = ing_mensual #guardo los ingresos de noviembre en un diccionario
            elif i==52:
                ingresos_mensuales['Diciembre'] = ing_mensual #guardo los ingresos de diciembre en un diccionario
            ing_mensual = 0 #reinicio los ingresos de mensuales
    pizzas = pizzas[:10] #guardo las 10 pizzas mas pedidas
    pzs = {}
    for p in range(len(pizzas)):
        pzs[pizzas[p][0]] = pizzas[p][1] #guardo las pizzas en un diccionario
    imagenes = []
    workbook = xlsxwriter.Workbook('reporte_ejecutivo.xlsx') #creo el archivo excel
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 40)
    worksheet.write('A3', 'Reporte ejecutivo') #escribo el titulo
    worksheet.write('A4', 'Maven pizzas')
    worksheet.write('A5', 'Jorge Gómez Azor') #escribo el nombre del autor
    worksheet.write('C2', 'Predicción semanal de ingredientes,')
    worksheet.write('C3', 'ingredientes anuales,') #Escribo el título en esas posiciones del excel
    worksheet.write('C4', 'cantidad de pedidos')
    worksheet.write('C5', '        y           ')
    worksheet.write('C6', 'pizzas más pedidas')
    worksheet.write('C7', 'INGRESOS: ' + str(int(ingresos))+'€') #escribo los ingresos
    plt.bar(ingresos_mensuales.keys(), ingresos_mensuales.values()) #grafico los ingredientes anuales
    plt.xticks(rotation=90, fontsize=5)
    plt.title('Ingresos mensuales')
    plt.savefig('ingresos_mensuales.png', bbox_inches='tight')
    plt.clf()
    worksheet.insert_image('B8', 'ingresos_mensuales.png') #inserto la imagen en el excel
    worksheet.insert_image('A7', 'maven_pizzas.png')
    worksheet = workbook.add_worksheet() #creo una nueva hoja
    worksheet.set_column('A:A', 90)
    posiciones = ['A2','B2','C2','D2','E2','F2','G2','H2','I2','J2','K2','L2','M2','N2','O2','P2','Q2','R2','S2','T2','U2','V2','W2','X2','Y2','Z2','A25','B25','C25','D250','E25','F25','G25','H25','I25','J25','K25','L25','M25','N25','O25','P25','Q25','R25','S25','T25','U25','V25','W25','X25','Y25','Z25','A50','B50','C50','D50'] #guardo las posiciones en una lista
    for i in range(25):
        worksheet.set_column(posiciones[i][0]+':'+posiciones[i][0], 90) #aumento el tamaño de las columnas
    worksheet.write('A1', 'Predicción semanal de ingredientes')
    for i in range(len(diccs)):
        plt.bar(diccs[i].keys(), diccs[i].values()) #grafico los diccionarios
        plt.xticks(rotation=90, fontsize=5)
        plt.title('Semana ' + str(i+1))
        plt.savefig('semana' + str(i+1) + '.png', bbox_inches='tight') #guardo los graficos
        imagenes.append('semana' + str(i+1) + '.png') 
        plt.clf() #limpio el grafico
        worksheet.insert_image(posiciones[i], imagenes[i])
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 20)
    for i in range(2,6):
        worksheet.set_column(posiciones[i][0]+':'+posiciones[i][0], 90) #aumento el tamaño de las columnas
    plt.bar(ing_ans.keys(), ing_ans.values()) #grafico los ingredientes anuales
    plt.xticks(rotation=90, fontsize=5)
    plt.title('Ingredientes anuales')
    plt.savefig('ingredientes_anuales.png', bbox_inches='tight')
    imagenes.append('ingredientes_anuales.png')
    plt.clf()
    worksheet.insert_image('D2', imagenes[-1])
    plt.bar(pzs.keys(), pzs.values()) #grafico las pizzas mas pedidas
    plt.xticks(rotation=90, fontsize=5)
    plt.title('Pizzas mas pedidas')
    plt.savefig('pizzas_mas_pedidas.png', bbox_inches='tight') #guardo el grafico
    imagenes.append('pizzas_mas_pedidas.png')
    plt.clf()
    eje_x = []
    eje_y = []
    worksheet.insert_image('D28', imagenes[-1])
    for i in range(len(cant_pedidos)): #escribo la cantidad de pedidos por semana
        plt.bar(i+1, len(cant_pedidos[i]))
        eje_x.append(str(i+1))
        eje_y.append(len(cant_pedidos[i]))
    plt.title('Cantidad de pedidos por semana')
    plt.savefig('cantidad_pedidos.png', bbox_inches='tight') #guardo el grafico
    imagenes.append('cantidad_pedidos.png')
    plt.clf()
    worksheet.insert_image('C28', imagenes[-1]) #inserto el grafico
    workbook.close()
    rows = []
    for i in range(len(eje_x)):
        rows.append([eje_x[i], eje_y[i]])
    wb = openpyxl.load_workbook('reporte_ejecutivo.xlsx')
    ws = wb['Sheet3']
    ws.append([None,'Cantidad de pedidos por semana'])
    for row in rows:
        ws.append(row)
    #draw a 3D bar chart
    chart = BarChart3D() #creo el grafico
    chart.type = "col" #tipo de grafico
    chart.style = 13 #estilo del grafico
    chart.title = "Cantidad de pedidos por semana (hecho en excel)" #titulo del grafico
    chart.y_axis.title = 'Cantidad de pedidos'
    chart.x_axis.title = 'Semana' #titulos de los ejes
    data = Reference(ws, min_col=2, min_row=1, max_col=2, max_row=52) #datos de la tabla
    cats = Reference(ws, min_col=1, min_row=1, max_col=1, max_row=52) #datos de la tabla
    chart.add_data(data, titles_from_data=True) #agrego los datos al grafico
    chart.set_categories(cats) #agrego las categorias al grafico
    chart.width = 16 #ancho del grafico
    chart.height = 10 #alto del grafico
    ws.add_chart(chart, "C2") #inserto el grafico
    wb.save('reporte_ejecutivo.xlsx') #guardo el excel
    wb.close() #cierro el excel

if __name__ == '__main__':
    order_details, orders, pizza_types,pizza = extract()
    order_details, orders = limpiar_datos(order_details,orders)
    diccs, ingredientes_anuales, cant_pedidos, pizzas, pedidos = transform(order_details, orders, pizza_types,pizza)
    load(diccs, ingredientes_anuales, cant_pedidos, pizzas, pedidos, pizza)