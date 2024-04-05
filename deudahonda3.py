import tkinter as tk
from tkinter import messagebox, ttk, simpledialog
import pyodbc
import re
import os
from configparser import ConfigParser

# Cargar las credenciales de la base de datos desde un archivo de configuración
config = ConfigParser()
config.read('config.ini')
db_server = config.get('database', 'server')
db_database = config.get('database', 'database')
db_user = config.get('database', 'user')
db_password = config.get('database', 'password')

def ingresar_datos():
    # Validar los datos ingresados
    numero_factura = numero_factura_entry.get().strip()
    if not numero_factura:
        messagebox.showerror("Error", "Por favor, ingrese un número de factura válido.")
        return

    monto_vencimiento = monto_vencimiento_entry.get().replace(',', '.')
    if not monto_vencimiento or not validar_monto_vencimiento(monto_vencimiento):
        messagebox.showerror("Error", "Por favor, ingrese un monto de vencimiento válido.")
        return

    fecha_vencimiento = fecha_vencimiento_entry.get().strip()
    if not fecha_vencimiento:
        messagebox.showerror("Error", "Por favor, ingrese una fecha de vencimiento válida.")
        return

    estado = estado_combobox.get()
    cotizacion = cotizacion_entry.get().strip()
    if not cotizacion:
        messagebox.showerror("Error", "Por favor, ingrese una cotización válida.")
        return

    moneda = moneda_combobox.get()

    # Insertar los datos en la base de datos
    try:
        conexion_str = (
            f"DRIVER={{SQL Server}};"
            f"SERVER={db_server};"
            f"DATABASE={db_database};"
            f"UID={db_user};"
            f"PWD={db_password};"
            "TIMEOUT=180;"
        )
        conn = pyodbc.connect(conexion_str)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO DeudaHonda (NumeroFactura, MontoVencimiento, FechaVencimiento, Estado, Cotizacion, Moneda) VALUES (?, ?, ?, ?, ?, ?)",
                       (numero_factura, monto_vencimiento, fecha_vencimiento, estado, cotizacion, moneda))
        conn.commit()
        messagebox.showinfo("Éxito", "Los datos han sido ingresados correctamente en la base de datos.")
        conn.close()
        
        # Actualizar la lista de datos ingresados
        actualizar_lista()

    except pyodbc.Error as e:
        messagebox.showerror("Error", f"No se pudo ingresar los datos en la base de datos: {e.args[1]}")

def actualizar_lista():
    # Limpiar la lista actual
    lista_datos.delete(0, tk.END)
    
    # Conectar a la base de datos y recuperar los datos ingresados
    try:
        conexion_str = (
            f"DRIVER={{SQL Server}};"
            f"SERVER={db_server};"
            f"DATABASE={db_database};"
            f"UID={db_user};"
            f"PWD={db_password};"
            "TIMEOUT=180;"
        )
        conn = pyodbc.connect(conexion_str)
        cursor = conn.cursor()
        cursor.execute("SELECT NumeroFactura, MontoVencimiento, FechaVencimiento, Estado, Cotizacion, Moneda FROM DeudaHonda")
        
        # Agregar los datos a la lista
        for row in cursor.fetchall():
            lista_datos.insert(tk.END, f"Factura: {row.NumeroFactura}, Monto: {row.MontoVencimiento}, Fecha: {row.FechaVencimiento}, Estado: {row.Estado}, Cotización: {row.Cotizacion}, Moneda: {row.Moneda}")
        
        conn.close()
    except pyodbc.Error as e:
        messagebox.showerror("Error", f"No se pudieron recuperar los datos de la base de datos: {e.args[1]}")

def actualizar_monto_vencimiento():
    seleccionado = lista_datos.curselection()
    if seleccionado:
        indice = seleccionado[0]
        texto_seleccionado = lista_datos.get(indice)
        numero_factura = texto_seleccionado.split(',')[0].split(':')[1].strip()
        nuevo_monto = monto_vencimiento_entry.get()

        try:
            conexion_str = (
                f"DRIVER={{SQL Server}};"
                f"SERVER={db_server};"
                f"DATABASE={db_database};"
                f"UID={db_user};"
                f"PWD={db_password};"
                "TIMEOUT=180;"
            )
            conn = pyodbc.connect(conexion_str)
            cursor = conn.cursor()
            cursor.execute("UPDATE DeudaHonda SET MontoVencimiento=? WHERE NumeroFactura=?", (nuevo_monto, numero_factura))
            conn.commit()
            messagebox.showinfo("Éxito", f"El monto de vencimiento para la factura {numero_factura} ha sido actualizado en la base de datos.")
            conn.close()

            actualizar_lista()
        except pyodbc.Error as e:
            messagebox.showerror("Error", f"No se pudo actualizar el monto de vencimiento en la base de datos: {e.args[1]}")
    else:
        messagebox.showerror("Error", "Por favor, selecciona un registro de la lista antes de actualizar el monto de vencimiento.")
  
def actualizar_estado():
    # Obtener el índice seleccionado en la lista de datos
    seleccionado = lista_datos.curselection()
    if seleccionado:
        indice = seleccionado[0]
        # Obtener el texto completo del elemento seleccionado
        texto_seleccionado = lista_datos.get(indice)
        # Obtener el número de factura del texto seleccionado
        numero_factura = texto_seleccionado.split(',')[0].split(':')[1].strip()
        # Obtener el estado seleccionado
        estado = estado_combobox.get()

        # Actualizar el estado en la base de datos
        try:
            conexion_str = (
                f"DRIVER={{SQL Server}};"
                f"SERVER={db_server};"
                f"DATABASE={db_database};"
                f"UID={db_user};"
                f"PWD={db_password};"
                "TIMEOUT=180;"
            )
            conn = pyodbc.connect(conexion_str)
            cursor = conn.cursor()
            cursor.execute("UPDATE DeudaHonda SET Estado=? WHERE NumeroFactura=?", (estado, numero_factura))
            conn.commit()
            messagebox.showinfo("Éxito", f"El estado para la factura {numero_factura} ha sido actualizado en la base de datos.")
            conn.close()

            # Actualizar la lista de datos
            actualizar_lista()
        except pyodbc.Error as e:
            messagebox.showerror("Error", f"No se pudo actualizar el estado en la base de datos: {e.args[1]}")
    else:
        messagebox.showerror("Error", "Por favor, selecciona un registro de la lista antes de actualizar el estado.")

def eliminar_registro():
    # Obtener el índice seleccionado en la lista de datos
    seleccionado = lista_datos.curselection()
    if seleccionado:
        indice = seleccionado[0]
        # Obtener el texto completo del elemento seleccionado
        texto_seleccionado = lista_datos.get(indice)
        # Obtener el número de factura y la moneda del texto seleccionado
        numero_factura = texto_seleccionado.split(',')[0].split(':')[1].strip()
        moneda = texto_seleccionado.split(',')[-1].split(':')[1].strip()

        # Verificar si hay registros con el mismo número de factura pero diferente moneda
        try:
            conexion_str = (
                f"DRIVER={{SQL Server}};"
                f"SERVER={db_server};"
                f"DATABASE={db_database};"
                f"UID={db_user};"
                f"PWD={db_password};"
                "TIMEOUT=180;"
            )
            conn = pyodbc.connect(conexion_str)
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM DeudaHonda WHERE NumeroFactura=? AND Moneda<>?", (numero_factura, moneda))
            count = cursor.fetchone()[0]
            conn.close()

            if count > 0:
                # Hay registros con el mismo número de factura pero diferente moneda
                respuesta = messagebox.askyesno("Confirmar eliminación", f"Hay registros con el mismo número de factura pero diferente moneda.\n¿Desea eliminar el registro con número de factura {numero_factura} en {moneda}?")
                if respuesta:
                    eliminar_registro_por_moneda(numero_factura, moneda)
            else:
                # No hay registros con el mismo número de factura pero diferente moneda
                eliminar_registro_por_moneda(numero_factura, moneda)
        except pyodbc.Error as e:
            messagebox.showerror("Error", f"No se pudo verificar los registros en la base de datos: {e.args[1]}")
    else:
        messagebox.showerror("Error", "Por favor, selecciona un registro de la lista antes de eliminar.")

def eliminar_registro_por_moneda(numero_factura, moneda):
    try:
        conexion_str = (
            f"DRIVER={{SQL Server}};"
            f"SERVER={db_server};"
            f"DATABASE={db_database};"
            f"UID={db_user};"
            f"PWD={db_password};"
            "TIMEOUT=180;"
        )
        conn = pyodbc.connect(conexion_str)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM DeudaHonda WHERE NumeroFactura=? AND Moneda=?", (numero_factura, moneda))
        conn.commit()
        messagebox.showinfo("Éxito", f"El registro con número de factura {numero_factura} en {moneda} ha sido eliminado de la base de datos.")
        conn.close()

        # Actualizar la lista de datos
        actualizar_lista()
    except pyodbc.Error as e:
        messagebox.showerror("Error", f"No se pudo eliminar el registro de la base de datos: {e.args[1]}")

def actualizar_monto_vencimiento():
    # Obtener el índice seleccionado en la lista de datos
    seleccionado = lista_datos.curselection()
    if seleccionado:
        indice = seleccionado[0]
        # Obtener el texto completo del elemento seleccionado
        texto_seleccionado = lista_datos.get(indice)
        # Obtener el número de factura del texto seleccionado
        numero_factura = texto_seleccionado.split(',')[0].split(':')[1].strip()

        # Solicitar al usuario el nuevo monto de vencimiento
        nuevo_monto_vencimiento = simpledialog.askfloat("Actualizar Monto de Vencimiento", f"Ingrese el nuevo monto de vencimiento para la factura {numero_factura}:")
        if nuevo_monto_vencimiento is not None:
            try:
                conexion_str = (
                    f"DRIVER={{SQL Server}};"
                    f"SERVER={db_server};"
                    f"DATABASE={db_database};"
                    f"UID={db_user};"
                    f"PWD={db_password};"
                    "TIMEOUT=180;"
                )
                conn = pyodbc.connect(conexion_str)
                cursor = conn.cursor()
                cursor.execute("UPDATE DeudaHonda SET MontoVencimiento=? WHERE NumeroFactura=?", (nuevo_monto_vencimiento, numero_factura))
                conn.commit()
                messagebox.showinfo("Éxito", f"El monto de vencimiento para la factura {numero_factura} ha sido actualizado a {nuevo_monto_vencimiento}.")
                conn.close()

                # Actualizar la lista de datos
                actualizar_lista()
            except pyodbc.Error as e:
                messagebox.showerror("Error", f"No se pudo actualizar el monto de vencimiento en la base de datos: {e.args[1]}")
    else:
        messagebox.showerror("Error", "Por favor, selecciona un registro de la lista antes de actualizar el monto de vencimiento.")

def calcular_suma_deuda(moneda):
    # Conectar a la base de datos y recuperar la suma de montos de registros con estado "DEUDA" y la moneda especificada
    try:
        conexion_str = (
            f"DRIVER={{SQL Server}};"
            f"SERVER={db_server};"
            f"DATABASE={db_database};"
            f"UID={db_user};"
            f"PWD={db_password};"
            "TIMEOUT=180;"
        )
        conn = pyodbc.connect(conexion_str)
        cursor = conn.cursor()
        cursor.execute(f"SELECT SUM(MontoVencimiento) FROM DeudaHonda WHERE Estado='DEUDA' AND Moneda=?", (moneda,))
        suma_deuda = cursor.fetchone()[0]
        conn.close()

        if suma_deuda is None:
            suma_deuda = 0

        messagebox.showinfo(f"Suma de Deuda en {moneda}", f"La suma de los montos de los registros con estado 'DEUDA' en {moneda.lower()} es: {suma_deuda}")

    except pyodbc.Error as e:
        messagebox.showerror("Error", f"No se pudo calcular la suma de deuda en {moneda.lower()}: {e.args[1]}")

def calcular_suma_deuda_pesos():
    calcular_suma_deuda("PESOS")

def calcular_suma_deuda_dolares():
    calcular_suma_deuda("DOLARES")

# Función de validación para el campo Monto de Vencimiento
def validar_monto_vencimiento(input):
    # Se permite un máximo de un punto y una coma
    return bool(re.match(r'^\d*[.,]?\d*$', input))

# Crear la ventana principal
root = tk.Tk()
root.title("Gestión de Deuda")
root.geometry("1200x500")  # Tamaño de la ventana
root.configure(bg="#FFFFFF")  # Color de fondo

# Estilo
style = ttk.Style()
style.theme_use("clam")

# Estilo para botones
style.configure("TButton", font=("Arial", 12))

# Estilo para etiquetas
style.configure("TLabel", font=("Arial", 12))

# Estilo para el Entry
style.configure("TEntry", font=("Arial", 12))

# Estilo para el Combobox
style.configure("TCombobox", font=("Arial", 12))

# Estilo para la lista
style.configure("TListbox", font=("Arial", 12))

# Crear y posicionar los elementos de la interfaz
numero_factura_label = ttk.Label(root, text="Número de Factura:", background="#FFFFFF")
numero_factura_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
numero_factura_entry = ttk.Entry(root)
numero_factura_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

monto_vencimiento_label = ttk.Label(root, text="Monto de Vencimiento:", background="#FFFFFF")
monto_vencimiento_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
# Asociar la función de validación al Entry del monto de vencimiento
validacion_monto_vencimiento = root.register(validar_monto_vencimiento)
monto_vencimiento_entry = ttk.Entry(root, validate="key", validatecommand=(validacion_monto_vencimiento, '%P'))
monto_vencimiento_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

fecha_vencimiento_label = ttk.Label(root, text="Fecha de Vencimiento:", background="#FFFFFF")
fecha_vencimiento_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
fecha_vencimiento_entry = ttk.Entry(root)
fecha_vencimiento_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

estado_label = ttk.Label(root, text="Estado:", background="#FFFFFF")
estado_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
estado_combobox = ttk.Combobox(root, values=["DEUDA", "CANCELADO"])
estado_combobox.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

cotizacion_label = ttk.Label(root, text="Cotización:", background="#FFFFFF")
cotizacion_label.grid(row=4, column=0, padx=10, pady=5, sticky="w")
cotizacion_entry = ttk.Entry(root)
cotizacion_entry.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

moneda_label = ttk.Label(root, text="Moneda:", background="#FFFFFF")
moneda_label.grid(row=5, column=0, padx=10, pady=5, sticky="w")
moneda_combobox = ttk.Combobox(root, values=["PESOS", "DOLARES"])
moneda_combobox.grid(row=5, column=1, padx=10, pady=5, sticky="ew")

ingresar_button = ttk.Button(root, text="Ingresar Datos", command=ingresar_datos)
ingresar_button.grid(row=6, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

actualizar_estado_button = ttk.Button(root, text="Actualizar Estado", command=actualizar_estado)
actualizar_estado_button.grid(row=7, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

eliminar_button = ttk.Button(root, text="Eliminar Registro", command=eliminar_registro)
eliminar_button.grid(row=8, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

calcular_suma_pesos_button = ttk.Button(root, text="Calcular Suma de Deuda en Pesos", command=calcular_suma_deuda_pesos)
calcular_suma_pesos_button.grid(row=9, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

calcular_suma_dolares_button = ttk.Button(root, text="Calcular Suma de Deuda en Dolares", command=calcular_suma_deuda_dolares)
calcular_suma_dolares_button.grid(row=10, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

# Crear un marco para la lista de datos
datos_frame = tk.Frame(root, bg="#FFFFFF")
datos_frame.grid(row=0, column=2, rowspan=12, padx=10, pady=10, sticky="nsew")

# Crear una lista para mostrar los datos ingresados
lista_datos = tk.Listbox(datos_frame, width=70, height=20, font=("Arial", 12), background="#FFFFFF")  # Tamaño de la lista
lista_datos.pack(fill=tk.BOTH, expand=True)

# Scrollbar para la lista
scrollbar = ttk.Scrollbar(datos_frame, orient="vertical", command=lista_datos.yview)
scrollbar.pack(side="right", fill="y")
lista_datos.config(yscrollcommand=scrollbar.set)

actualizar_monto_vencimiento_button = ttk.Button(root, text="Actualizar Monto de Vencimiento", command=actualizar_monto_vencimiento)
actualizar_monto_vencimiento_button.grid(row=11, column=0, columnspan=2, padx=10, pady=5, sticky="ew")

# ... (código existente para crear el marco y la lista para los datos)

# Actualizar la lista de datos al iniciar la aplicación
actualizar_lista()

# Ajustar el peso de las columnas y filas
root.grid_columnconfigure(2, weight=1)
root.grid_rowconfigure(13, weight=1)

# Ejecutar la interfaz
root.mainloop()
