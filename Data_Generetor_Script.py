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
        "addresses": f"{random.choice(['Calle','Carrera'])} {random.randint(1, 100)} #{random.randint(1, 20)}-{random.randint(1, 100)}",
        "city": random.choice(["Bogotá", "Medellín", "Cali", "Cartagena", "Ibagué"]),
        "Country": "Colombia"
    }
    clients.append(clientes)

#Generate data for "Products" sheet.
products = []
for i in range(1, 21):
    unit_price = round(random.uniform(10, 100), 2)
    discount_rate = random.uniform(1, 50)
    producto = {
        "product_id": ''.join(random.choices(string.ascii_uppercase, k=3)) + '-' + ''.join(random.choices(string.ascii_lowercase + string.digits, k=9)),
        "product_name": f"Producto {i}",
        "Description": f"Product Description {i}.",
        "category": random.choice(["Electrónica", "Ropa", "Hogar", "Deportes", "Juguetes"]),
        "order_weight": round(random.uniform(0, 20), 2),
        "unit_price": unit_price,
        "Unit_Cost": round(unit_price * (1 - (discount_rate/100)),2)
    }
    products.append(producto)

#Generate data for "Ordes" sheet.
orders = []
for i in range(1, 35001):
    start_date = datetime.date(2021, 1, 1)
    end_date = datetime.date.today()
    days_between = (end_date - start_date).days
    delta = end_date - start_date
    date = start_date + datetime.timedelta(days=random.randint(0, delta.days))
    status = random.choice(["pendiente","pagado","rechazado"])
    delivery_status = random.choice(["pendiente","enviado","entregado"]) if status == "pagado" else "pendiente" if status == "pendiente" else "rechazado"
    pedido = {
        "Order_id": i,
        "Order_date": date,
        "status": status,
        "delivery_status": delivery_status,
        "client_id": random.choice(clients)["client_id"],
        "product_id": random.choice(products)["product_id"],
        "items": random.randint(1, 5),
        "discount_amount": round(random.uniform(0, 10), 2)
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
