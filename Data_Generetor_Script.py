#Import necessary libraries
import pandas as pd
import random
import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import string

#Generate data for "Clients" sheet.
clients = []
for i in range(1, 1500):
    clientes = {
        "client_id": ''.join(random.choices(string.digits, k=7)),
        "name": random.choice(['Juan', 'María', 'Pedro', 'Laura', 'Gabriela', 'José', 'Carla', 'Antonio', 'Ana', 'Jorge', 'Valeria', 'Miguel', 'Cristina', 'Fernando', 'Julia', 'Ricardo', 'Renata', 'Diego', 'Sofía', 'Daniel']),
        "last_name": random.choice(['González', 'Pérez', 'Martínez', 'Sánchez', 'López', 'Gómez', 'Hernández', 'Fernández', 'Rodríguez', 'García', 'Díaz', 'Torres', 'Ramos', 'Ruiz', 'Moreno', 'Alonso', 'Romero', 'Jiménez', 'Álvarez', 'Vargas']),
        "addresses": f"{random.choice(['Calle','Carrera'])} {random.randint(1, 100)} # {random.randint(1, 20)}-{random.randint(1, 100)}",
        "city": random.choice(["Bogotá", "Medellín", "Cali", "Cartagena", "Ibagué"]),
        "Country": "Colombia"
    }
    clients.append(clientes)

#Generate data for "Products" sheet.
products = []
for i in range(1, 21):
    producto = {
        "product_id": ''.join(random.choices(string.ascii_uppercase, k=4)) + str(''.join(random.choices(string.digits, k=3))),
        "product_name": f"Producto {i}",
        "Description": f"Descripción del producto {i}.",
        "price": round(random.uniform(10, 100), 2),
        "category": random.choice(["Electrónica", "Ropa", "Hogar", "Deportes", "Juguetes"])
    }
    products.append(producto)

#Generate data for "Ordes" sheet.
orders = []
for i in range(1, 5001):
    fecha = datetime.datetime.now() - datetime.timedelta(days=random.randint(0, 30))
    pedido = {
        "Order_id": i,
        "Order_date": fecha.date(),
        "status": random.choice(["pendiente","pagado","rechazado"]),
        "delivery_status": random.choice(["pendiente","enviado","entregado"]),
        "client_id": random.choice(clients)["client_id"],
        "product_id": random.choice(products)["product_id"],
        "items": random.randint(1, 5),
        "discount_amount": round(random.uniform(0, 20), 2)
    }
    orders.append(pedido)

#Create a DataFrame for each sheet.
df_clients = pd.DataFrame(clients)
df_products = pd.DataFrame(products)
df_orders = pd.DataFrame(orders)

#Create a excel workbook
workbook = Workbook()

#Insert each DataFrame in a separate Excel sheets
workbook.create_sheet("Clients")
sheet = workbook["Clients"]
for row in dataframe_to_rows(df_clients, index=False, header=True):
    sheet.append(row)

workbook.create_sheet("Products")
sheet = workbook["Products"]
for row in dataframe_to_rows(df_products, index=False, header=True):
    sheet.append(row)

workbook.create_sheet("Orders")
sheet = workbook["Orders"]
for row in dataframe_to_rows(df_orders, index=False, header=True):
    sheet.append(row)

#Save workbook.
workbook.save(filename="Raw_Data.xlsx")