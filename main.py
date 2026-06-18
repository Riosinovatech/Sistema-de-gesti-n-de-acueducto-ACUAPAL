import customtkinter as ctk
import sqlite3
import os
import sys
import shutil
from tkinter import messagebox
from tkinter import filedialog
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image as XLImage
from reportlab.pdfgen import canvas
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
import fitz
from cloud.subir_imagen import (
    subir_imagen,
    eliminar_imagen
)
def ruta_recurso(ruta_relativa):

    if getattr(sys, 'frozen', False):

        return os.path.join(
            os.path.dirname(sys.executable),
            ruta_relativa
        )

    return os.path.join(
        os.path.dirname(__file__),
        ruta_relativa
    )

# Configuración visual
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Base de datos
conexion = sqlite3.connect("aguapal.db")
cursor = conexion.cursor()
cursor.execute("PRAGMA journal_mode=WAL")
conexion.commit()


# ==========================
# ==========================
# RESPALDO AUTOMÁTICO
# ==========================
def crear_backup():
    if not os.path.exists("backups"):
        os.makedirs("backups")
    # ==========================
    # LIMPIEZA MENSUAL BACKUPS
    # ==========================

    mes_actual = datetime.now().strftime("%Y-%m")

    archivo_control = "backups/mes_backup.txt"

    if not os.path.exists(archivo_control):

        with open(archivo_control, "w") as f:
            f.write(mes_actual)

    else:

        with open(archivo_control, "r") as f:
            ultimo_mes = f.read().strip()

        if ultimo_mes != mes_actual:

            # BORRAR BACKUPS LOCALES
            for archivo in os.listdir("backups"):

                if archivo.startswith("aguapal_") and archivo.endswith(".db"):

                    os.remove(
                        os.path.join(
                            "backups",
                            archivo
                        )
                    )

            # BORRAR BACKUPS GOOGLE DRIVE
            carpeta_drive = "G:/Mi unidad/acuapal_backups"

            if os.path.exists(carpeta_drive):

                for archivo in os.listdir(carpeta_drive):

                    if archivo.startswith("aguapal_") and archivo.endswith(".db"):

                        os.remove(
                            os.path.join(
                                carpeta_drive,
                                archivo
                            )
                        )

            # GUARDAR NUEVO MES
            with open(archivo_control, "w") as f:
                f.write(mes_actual)

    fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

    origen = "aguapal.db"
    destino = f"backups/aguapal_{fecha}.db"

    if os.path.exists(origen):
        
        # Backup local
        shutil.copy(
            origen,
            destino
        )

        # Backup Google Drive
        carpeta_drive = "G:/Mi unidad/acuapal_backups"

        if os.path.exists(carpeta_drive):

            destino_drive = (
                f"{carpeta_drive}/"
                f"aguapal_{fecha}.db"
            )

            shutil.copy(
                origen,
                destino_drive
            )


def calcular_estado(deuda):
    if deuda <= 12000:
        return "Al día"
    elif deuda <= 50000:
        return "Mora"
    elif deuda <= 100000:
        return "Suspensión"
    else:
        return "Suspensión inmediata"


def cobro_mensual_global():
    from datetime import datetime

    hoy = datetime.now()

    # Obtener última fecha
    cursor.execute("SELECT ultima_fecha FROM control WHERE id=1")
    dato = cursor.fetchone()

    # Primera vez (solo guarda fecha)
    if not dato:
        cursor.execute(
            "INSERT INTO control (id, ultima_fecha) VALUES (1, ?)",
            (hoy.strftime("%Y-%m-%d"),)
        )
        conexion.commit()
        return

    ultima_fecha = datetime.strptime(dato[0], "%Y-%m-%d")
    # ESPERAR HASTA EL DÍA 6
    if hoy.day < 6:
        return

    # Calcular meses transcurridos
    meses = (hoy.year - ultima_fecha.year) * 12 + (hoy.month - ultima_fecha.month)

    if meses <= 0:
        return

    aumento = meses * 12000

    # 🔥 ACTUALIZAR TODOS LOS USUARIOS
    cursor.execute("SELECT id, deuda FROM usuarios")
    usuarios = cursor.fetchall()

    for user_id, deuda in usuarios:
        nueva_deuda = deuda + aumento

        # 🔥 AQUÍ SE ACTUALIZA EL ESTADO AUTOMÁTICO
        estado = calcular_estado(nueva_deuda)

        cursor.execute("""
            UPDATE usuarios
            SET deuda=?, estado=?
            WHERE id=?
        """, (nueva_deuda, estado, user_id))

    # Guardar nueva fecha
    cursor.execute("""
        UPDATE control SET ultima_fecha=?
        WHERE id=1
    """, (hoy.strftime("%Y-%m-%d"),))

    conexion.commit()


# ==========================
# CREAR TABLAS
# ==========================
cursor.execute("""
CREATE TABLE IF NOT EXISTS usuarios (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT,
    direccion TEXT,
    telefono TEXT,
    estado TEXT,
    deuda INTEGER DEFAULT 0,
    link_factura TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS administrador (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    usuario TEXT,
    password TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS pagos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    usuario_id INTEGER,
    monto INTEGER
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS auditoria (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    accion TEXT,
    usuario_afectado TEXT,
    fecha TEXT
)
""")
cursor.execute("""
CREATE TABLE IF NOT EXISTS gastos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    concepto TEXT,
    valor INTEGER,
    fecha TEXT
)
""")
cursor.execute("""
CREATE TABLE IF NOT EXISTS cierres (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    fecha TEXT,
    cantidad_pagos INTEGER,
    total_ingresos INTEGER,
    total_gastos INTEGER,
    efectivo INTEGER
)
""")

conexion.commit()

# 🔥 NUEVA TABLA DE CONTROL
cursor.execute("""
CREATE TABLE IF NOT EXISTS control (
    id INTEGER PRIMARY KEY,
    ultima_fecha TEXT
)
""")

conexion.commit()
try:
    cursor.execute("""
    ALTER TABLE usuarios
    ADD COLUMN mes_factura TEXT
    """)
except:
    pass

try:
    cursor.execute("""
    ALTER TABLE usuarios
    ADD COLUMN cloud_id TEXT
    """)
except:
    pass
try:
    cursor.execute("""
    ALTER TABLE pagos
    ADD COLUMN fecha_pago TEXT
    """)
except:
    pass

conexion.commit()
try:
    cursor.execute("""
    ALTER TABLE pagos
    ADD COLUMN cerrado INTEGER DEFAULT 0
    """)
except:
    pass

try:
    cursor.execute("""
    ALTER TABLE gastos
    ADD COLUMN cerrado INTEGER DEFAULT 0
    """)
except:
    pass

conexion.commit()



# ==========================
# INICIO DEL SISTEMA
# ==========================

# Backup automático
crear_backup()

# 🔥 COBRO AUTOMÁTICO MENSUAL
cobro_mensual_global()


# Crear admin inicial
cursor.execute(
    "SELECT * FROM administrador WHERE usuario=?",
    ("ACUAPAL",)
)

if not cursor.fetchone():
    cursor.execute(
        "INSERT INTO administrador (usuario,password) VALUES (?,?)",
        ("ACUAPAL", "Acuapal123@")
    )
    conexion.commit()


def verificar_admin(usuario, password):
    cursor.execute(
        "SELECT * FROM administrador WHERE usuario=? AND password=?",
        (usuario, password)
    )
    return cursor.fetchone()
def verificar_seguridad(usuario, password):

    return (
        usuario == "ADMIN"
        and password == "Acuapal2026*"
    )


def verificar_reinicio(usuario, password):

    return (
        usuario == "SAMIR"
        and password == "ReiniciarAcuapal2026#"
    )
# ==========================
## ==========================
# VENTANA INICIO
# ==========================

class VentanaInicio(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Inicio")
        self.geometry("350x250")

        # ICONOS
        self.icon_historial = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/historal_factura.png"
            ),
            size=(20, 20)
        )
        self.icon_editar_usuarios = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/editar_usuario.png"
            ),
            size=(20, 20)
        )

        self.icon_reiniciar = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/reiniciar_system.png"
            ),
            size=(20, 20)
        )

        ctk.CTkLabel(
            self,
            text="Panel Inicio",
            font=("Arial", 20, "bold")
        ).pack(pady=20)

        ctk.CTkButton(
            self,
            text="Historial de Facturas",
            image=self.icon_historial,
            compound="left",
            width=250,
            height=40,
            command=self.abrir_historial_facturas
        ).pack(pady=10)
        ctk.CTkButton(
            self,
            text="Editar Usuario",
            image=self.icon_editar_usuarios,
            compound="left",
            width=250,
            height=40,
            command=self.editar_usuario
        ).pack(pady=10)
        ctk.CTkButton(
            self,
            text="Reiniciar Sistema",
            image=self.icon_reiniciar,
            compound="left",
            width=250,
            height=40,
            fg_color="red",
            hover_color="#990000",
            command=self.reiniciar_sistema
        ).pack(pady=10)

    def abrir_historial_facturas(self):

        self.destroy()

        v = VentanaHistorialFacturas()

        v.grab_set()
        v.focus()
        v.lift()
    def editar_usuario(self):

        self.destroy()

        v = VentanaEditarUsuario()

        v.grab_set()
        v.focus()
        v.lift()

    def reiniciar_sistema(self):

        self.destroy()

        v = VentanaReiniciarSistema()

        v.grab_set()
        v.focus()
        v.lift()
# ==========================
# HISTORIAL FACTURAS
# ==========================

class VentanaHistorialFacturas(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Historial de Facturas")
        self.geometry("700x500")

        self.archivos_mostrados = []

        self.buscar = ctk.CTkEntry(
            self,
            placeholder_text="Buscar factura..."
        )
        self.buscar.pack(pady=10)

        self.buscar.bind(
            "<KeyRelease>",
            self.buscar_facturas
        )

        self.lista = ctk.CTkTextbox(
            self,
            width=600,
            height=350
        )
        self.lista.pack(pady=10)

        ctk.CTkButton(
            self,
            text="📂 Abrir Factura",
            width=200,
            command=self.abrir_factura
        ).pack(pady=10)

        self.buscar_facturas()

    def buscar_facturas(self, event=None):

        texto = self.buscar.get().lower()

        self.lista.delete(
            "1.0",
            "end"
        )

        self.archivos_mostrados = []

        if not os.path.exists(
            "historial_facturas"
        ):
            return

        archivos = sorted(
            os.listdir(
                "historial_facturas"
            )
        )

        for archivo in archivos:

            # Mostrar solo PDFs
            if not archivo.lower().endswith(
                ".pdf"
            ):
                continue

            if (
                texto == ""
                or archivo.lower().startswith(
                    texto
                )
            ):

                self.archivos_mostrados.append(
                    archivo
                )

                self.lista.insert(
                    "end",
                    f"{archivo}\n"
                )

    def abrir_factura(self):

        try:

            indice = self.lista.index(
                "insert"
            ).split(".")[0]

            indice = int(
                indice
            ) - 1

            if (
                indice >= 0
                and indice < len(
                    self.archivos_mostrados
                )
            ):

                archivo = (
                    self.archivos_mostrados[
                        indice
                    ]
                )

                ruta = os.path.join(
                    "historial_facturas",
                    archivo
                )

                os.startfile(
                    ruta
                )

        except:

            messagebox.showerror(
                "Error",
                "Selecciona una factura"
            )
# ==========================
# EDITAR USUARIO
# ==========================

class VentanaEditarUsuario(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Editar Usuario")
        self.geometry("500x550")

        self.usuario_actual = None
        self.usuarios_mostrados = []

        self.buscar = ctk.CTkEntry(
            self,
            placeholder_text="Buscar usuario"
        )
        self.buscar.pack(pady=10)

        self.buscar.bind(
            "<KeyRelease>",
            self.buscar_usuarios
        )

        self.lista = ctk.CTkTextbox(
            self,
            width=450,
            height=120
        )
        self.lista.pack(pady=10)

        self.lista.bind(
            "<ButtonRelease-1>",
            self.seleccionar_usuario
        )

        self.nombre = ctk.CTkEntry(
            self,
            placeholder_text="Nombre"
        )
        self.nombre.pack(pady=10)

        self.direccion = ctk.CTkEntry(
            self,
            placeholder_text="Dirección"
        )
        self.direccion.pack(pady=10)

        self.telefono = ctk.CTkEntry(
            self,
            placeholder_text="Teléfono"
        )
        self.telefono.pack(pady=10)

        ctk.CTkButton(
            self,
            text="Guardar Cambios",
            command=self.guardar
        ).pack(pady=20)

    def buscar_usuarios(self, event=None):

        texto = self.buscar.get()

        self.lista.delete(
            "1.0",
            "end"
        )

        self.usuarios_mostrados = []

        if texto.strip() == "":
            return

        cursor.execute(
            "SELECT * FROM usuarios WHERE nombre LIKE ?",
            (texto + "%",)
        )

        usuarios = cursor.fetchall()

        for usuario in usuarios:

            self.usuarios_mostrados.append(
                usuario
            )

            self.lista.insert(
                "end",
                f"{usuario[1]} | {usuario[2]} | {usuario[3]}\n"
            )

    def seleccionar_usuario(self, event=None):

        try:

            indice = self.lista.index(
                "insert"
            ).split(".")[0]

            indice = int(
                indice
            ) - 1

            if indice < len(
                self.usuarios_mostrados
            ):

                usuario = self.usuarios_mostrados[
                    indice
                ]

                self.usuario_actual = usuario

                self.nombre.delete(
                    0,
                    "end"
                )

                self.nombre.insert(
                    0,
                    usuario[1]
                )

                self.direccion.delete(
                    0,
                    "end"
                )

                self.direccion.insert(
                    0,
                    usuario[2]
                )

                self.telefono.delete(
                    0,
                    "end"
                )

                self.telefono.insert(
                    0,
                    usuario[3]
                )

        except:
            pass

    def guardar(self):

        if not self.usuario_actual:
            messagebox.showerror(
                "Error",
                "Seleccione un usuario"
            )
            return
        if (
            self.nombre.get().strip() == ""
            or self.direccion.get().strip() == ""
            or self.telefono.get().strip() == ""
        ):
            messagebox.showerror(
                "Error",
                "Todos los campos son obligatorios"
            )
            return

        cursor.execute(
            """
            UPDATE usuarios
            SET nombre=?,
                direccion=?,
                telefono=?
            WHERE id=?
            """,
            (
                self.nombre.get(),
                self.direccion.get(),
                self.telefono.get(),
                self.usuario_actual[0]
            )
        )

        conexion.commit()

        messagebox.showinfo(
            "Éxito",
            "Datos actualizados correctamente"
        )

        self.destroy()
# ==========================
# REINICIAR SISTEMA
# ==========================

class VentanaReiniciarSistema(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Reiniciar Sistema")
        self.geometry("400x350")

        ctk.CTkLabel(
            self,
            text="Acceso restringido",
            font=("Arial", 20, "bold")
        ).pack(pady=20)

        self.usuario = ctk.CTkEntry(
            self,
            placeholder_text="Usuario administrador"
        )
        self.usuario.pack(pady=10)

        self.password = ctk.CTkEntry(
            self,
            placeholder_text="Contraseña",
            show="*"
        )
        self.password.pack(pady=10)

        ctk.CTkButton(
            self,
            text="🔒 Ingresar",
            fg_color="red",
            hover_color="#990000",
            command=self.verificar
        ).pack(pady=20)

    def verificar(self):

        admin = verificar_reinicio(
            self.usuario.get(),
            self.password.get()
        )

        if not admin:

            messagebox.showerror(
                "Error",
                "Credenciales incorrectas"
            )

            return

        respuesta = messagebox.askyesno(
            "Confirmar",
            "¿Seguro que deseas reiniciar el sistema?\n\n"
            "Esta acción eliminará TODOS los datos."
        )

        if not respuesta:
            return

        self.reiniciar_sistema()

    def reiniciar_sistema(self):

        try:

            # BORRAR TABLAS
            cursor.execute("DELETE FROM usuarios")
            cursor.execute("DELETE FROM pagos")
            cursor.execute("DELETE FROM auditoria")
            cursor.execute("DELETE FROM gastos")
            cursor.execute("DELETE FROM control")
            cursor.execute("DELETE FROM cierres")

            cursor.execute("""
            INSERT INTO control (id, ultima_fecha)
            VALUES (1, ?)
            """, (
                datetime.now().strftime("%Y-%m-%d"),
            ))

            conexion.commit()
            # LIMPIAR FACTURAS
            carpetas = [
                "facturas",
                "historial_facturas",
                "backups"
            ]

            for carpeta in carpetas:

                if os.path.exists(carpeta):

                    for archivo in os.listdir(carpeta):

                        ruta = os.path.join(
                            carpeta,
                            archivo
                        )

                        if os.path.isfile(ruta):

                            os.remove(ruta)

            messagebox.showinfo(
                "Éxito",
                "Sistema reiniciado correctamente"
            )

            self.destroy()

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )
# AGREGAR USUARIO
class VentanaAgregarUsuario(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Agregar Usuario")
        self.geometry("500x500")

        self.nombre = ctk.CTkEntry(self, placeholder_text="Nombre")
        self.nombre.pack(pady=10)

        self.direccion = ctk.CTkEntry(self, placeholder_text="Dirección")
        self.direccion.pack(pady=10)

        self.telefono = ctk.CTkEntry(self, placeholder_text="WhatsApp")
        self.telefono.pack(pady=10)
        # DEUDA DEL CLIENTE
        self.deuda = ctk.CTkEntry(self, placeholder_text="Deuda")
        self.deuda.pack(pady=10)

        ctk.CTkButton(
            self,
            text="Guardar",
            command=self.guardar_usuario
        ).pack(pady=20)

    # FUNCIÓN PARA GUARDAR USUARIO (estado automático usando función global)
    def guardar_usuario(self): 

        # 🔥 VALIDAR CAMPO VACÍO 
        if not self.deuda.get():
            messagebox.showerror("Error", "Ingrese la deuda")
            return

        try:
            deuda_valor = int(self.deuda.get())
        except ValueError:
            messagebox.showerror("Error", "Ingrese un valor numérico válido en deuda")
            return

        estado = calcular_estado(deuda_valor)

        cursor.execute("""
        INSERT INTO usuarios (nombre, direccion, telefono, estado, deuda)
        VALUES (?, ?, ?, ?, ?)
        """, (
            self.nombre.get(),
            self.direccion.get(),
            self.telefono.get(),
            estado,
            deuda_valor
        ))

        conexion.commit()
        messagebox.showinfo("Éxito", f"Usuario guardado con estado: {estado}")
        self.destroy()

# PAGOS
class VentanaRegistrarPago(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Registrar Pago")
        self.geometry("550x550")

        self.usuario_actual = None
        self.usuarios_mostrados = []

        self.nombre = ctk.CTkEntry(
            self,
            placeholder_text="Buscar usuario"
        )
        self.nombre.pack(pady=10)

        self.nombre.bind("<KeyRelease>", self.buscar_sugerencias)

        self.lista = ctk.CTkTextbox(
            self,
            width=450,
            height=150
        )
        self.lista.pack(pady=10)

        self.lista.bind("<ButtonRelease-1>", self.seleccionar_usuario)

        self.label_deuda = ctk.CTkLabel(
            self,
            text="Deuda actual: $0"
        )
        self.label_deuda.pack(pady=10)

        self.monto = ctk.CTkEntry(
            self,
            placeholder_text="Monto a pagar"
        )
        self.monto.pack(pady=10)

        ctk.CTkButton(
            self,
            text="Registrar Pago",
            command=self.registrar_pago
        ).pack(pady=20)

    def buscar_sugerencias(self, event=None):
        texto = self.nombre.get()

        self.lista.delete("1.0", "end")
        self.usuarios_mostrados = []

        if texto.strip() == "":
            return

        cursor.execute(
            "SELECT * FROM usuarios WHERE nombre LIKE ?",
            (texto + "%",)
        )

        usuarios = cursor.fetchall()

        for usuario in usuarios:
            self.usuarios_mostrados.append(usuario)

            self.lista.insert(
                "end",
                f"{usuario[1]} | Deuda: ${usuario[5]}\n"
            )

    def seleccionar_usuario(self, event=None):
        try:
            indice = self.lista.index("insert").split(".")[0]
            indice = int(indice) - 1

            if indice < len(self.usuarios_mostrados):
                usuario = self.usuarios_mostrados[indice]

                self.usuario_actual = usuario

                self.nombre.delete(0, "end")
                self.nombre.insert(0, usuario[1])

                self.label_deuda.configure(
                    text=f"Deuda actual: ${usuario[5]}"
                )
        except:
            pass
    def registrar_pago(self):
        if not self.usuario_actual:
            messagebox.showerror(
                "Error",
                "Selecciona un usuario"
            )
            return

        try:
            monto_pago = int(self.monto.get())
        except ValueError:
            messagebox.showerror("Error", "Ingrese un monto válido")
            return

        deuda_actual = int(self.usuario_actual[5])

        nueva_deuda = deuda_actual - monto_pago

        

        # 🔥 CALCULAR ESTADO AUTOMÁTICO
        estado = calcular_estado(nueva_deuda)

        # 🔥 ACTUALIZAR TODO
        cursor.execute(
            "UPDATE usuarios SET deuda=?, estado=? WHERE id=?",
            (nueva_deuda, estado, self.usuario_actual[0])
        )

       
        # FECHA DEL PAGO
        fecha_pago = datetime.now().strftime(
            "%d/%m/%Y %H:%M"
        )

        cursor.execute(
            """
            INSERT INTO pagos
            (usuario_id,monto,fecha_pago)
            VALUES (?,?,?)
            """,
            (
                self.usuario_actual[0],
                monto_pago,
                fecha_pago
            )
        )

        conexion.commit()
        # ==========================
        # GENERAR RECIBO DE PAGO
        # ==========================

        if not os.path.exists("recibos"):
            os.makedirs("recibos")

        recibo = Image.open(
            "imagenes/plantilla_recibo.png"
        )

        draw = ImageDraw.Draw(recibo)

        try:
            fuente = ImageFont.truetype(
                "arial.ttf",
                28
            )
        except:
            fuente = ImageFont.load_default()

        fecha_actual = datetime.now().strftime(
            "%d/%m/%Y"
        )

        numero_recibo = datetime.now().strftime(
            "%Y%m%d%H%M%S"
        )

        # DATOS
        nombre = self.usuario_actual[1]
        direccion = self.usuario_actual[2]

        saldo_anterior = deuda_actual
        valor_pagado = monto_pago
        saldo_pendiente = nueva_deuda

        # ESCRIBIR DATOS
        draw.text(
            (200, 455),
            numero_recibo,
            fill="black",
            font=fuente
        )

        draw.text(
            (697, 455),
            fecha_actual,
            fill="black",
            font=fuente
        )

        draw.text(
            (323, 571),
            nombre,
            fill="black",
            font=fuente
        )

        draw.text(
            (333, 665),
            direccion,
            fill="black",
            font=fuente
        )

        draw.text(
            (695, 779),
            f"${saldo_anterior:,}",
            fill="black",
            font=fuente
        )

        draw.text(
            (699, 871),
            f"${valor_pagado:,}",
            fill="black",
            font=fuente
        )

        draw.text(
            (699, 957),
            f"${saldo_pendiente:,}",
            fill="black",
            font=fuente
        )

        draw.text(
            (695, 1035),
            estado,
            fill="black",
            font=fuente
        )

        ruta_recibo = (
            f"recibos/Recibo_{nombre}_{numero_recibo}.jpg"
        )

        recibo.save(
            ruta_recibo
        )
       
        
        messagebox.showinfo(
            "Éxito",
            f"Pago registrado. Nueva deuda: ${nueva_deuda} - Estado: {estado}"
        )
         # PREGUNTAR SI QUIERE ENVIAR
        enviar = messagebox.askyesno(
            "WhatsApp",
            "¿Desea enviar este recibo por WhatsApp?"
        )

        if enviar:
            os.startfile("recibos")

        self.destroy()

# USUARIOS
class VentanaUsuarios(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Usuarios")
        self.geometry("800x600")
        ctk.CTkButton(
            self,
            text="📊 Exportar Lista",
            command=self.exportar_excel
        ).place(
            x=620,
            y=15
        )

        self.textbox = ctk.CTkTextbox(
            self,
            width=700, 
            height=500
        )
        self.textbox.pack(
            pady=(60, 20)
        )

        cursor.execute("SELECT * FROM usuarios")
        usuarios = cursor.fetchall()

        for u in usuarios:
            self.textbox.insert(
                "end",
                f"{u[1]} | {u[2]} | {u[3]} | {u[4]} | ${u[5]}\n"
            )
    def exportar_excel(self):

        ruta = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="Cobranza_Acuapal.xlsx"
        )

        if not ruta:
            return

        wb = Workbook()
        ws = wb.active

        ws.title = "Cobranza"

        # LOGO
        logo = XLImage("logo.jpg")
        logo.width = 120
        logo.height = 120
        ws.add_image(logo, "A1")

        # TITULO
        ws.merge_cells("C1:G1")
        ws["C1"] = "ACUAPAL"
        ws["C1"].font = Font(
            size=20,
            bold=True
        )
        ws["C1"].alignment = Alignment(
            horizontal="center"
        )

        ws.merge_cells("C2:G2")
        ws["C2"] = "LISTADO OFICIAL DE COBRANZA"
        ws["C2"].font = Font(
            size=14,
            bold=True
        )
        ws["C2"].alignment = Alignment(
            horizontal="center"
        )

        ws.merge_cells("C3:G3")
        ws["C3"] = f"Generado: {datetime.now().strftime('%d/%m/%Y')}"
        ws["C3"].alignment = Alignment(
            horizontal="center"
        )

        # ENCABEZADOS
        encabezados = [
            "Visitado",
            "Nombre",
            "Dirección",
            "Teléfono",
            "Estado",
            "Deuda",
            "Observación"
        ]

        for columna, titulo in enumerate(
            encabezados,
            start=1
        ):

            celda = ws.cell(
                row=7,
                column=columna
            )

            celda.value = titulo

            celda.font = Font(
                bold=True,
                color="FFFFFF"
            )

            celda.fill = PatternFill(
                "solid",
                fgColor="1F4E78"
            )

        cursor.execute(
            """
            SELECT * FROM usuarios
            ORDER BY direccion, nombre
            """
        )

        usuarios = cursor.fetchall()
        fila = 8

        for u in usuarios:

            ws.cell(fila, 1).value = ""
            ws.cell(fila, 2).value = u[1]
            ws.cell(fila, 3).value = u[2]
            ws.cell(fila, 4).value = u[3]
            ws.cell(fila, 5).value = u[4]
            ws.cell(fila, 6).value = u[5]
            ws.cell(fila, 7).value = ""

            fila += 1

            # ANCHO DE COLUMNAS
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 30
        ws.column_dimensions["C"].width = 35
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 18
        ws.column_dimensions["F"].width = 15
        ws.column_dimensions["G"].width = 40

        # FILTROS
        ws.auto_filter.ref = f"A7:G{fila-1}"
        #congelar encabezado
        ws.freeze_panes = "A8"

        wb.save(ruta)

        messagebox.showinfo(
            "Éxito",
            "Archivo Excel generado correctamente"
        )

# HISTORIAL
class VentanaHistorial(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Historial de Pagos")
        self.geometry("800x600")
        self.buscar = ctk.CTkEntry(
            self,
            placeholder_text="Buscar usuario"
        )
        self.buscar.pack(pady=10)
        self.lista = ctk.CTkTextbox(
            self,
            width=700,
            height=120
        )
        self.lista.pack(pady=5)

        self.lista.bind(
            "<ButtonRelease-1>",
            self.seleccionar_usuario
        )

        self.buscar.bind(
            "<KeyRelease>",
            self.buscar_usuario
        )

        self.usuarios_mostrados = []

        self.textbox = ctk.CTkTextbox(
            self,
            width=700,
            height=500
        )
        self.textbox.pack(pady=20)
        self.buscar_usuario()

    def buscar_usuario(self, event=None):

        texto = self.buscar.get()

        self.textbox.delete(
            "1.0",
            "end"
        )
        self.lista.delete(
            "1.0",
            "end"
        )

        self.usuarios_mostrados = []

        if texto.strip() == "":

            cursor.execute("""
            SELECT
                usuarios.nombre,
                pagos.monto,
                pagos.fecha_pago
            FROM pagos
            JOIN usuarios
            ON pagos.usuario_id = usuarios.id
            ORDER BY pagos.id DESC
            """)

            datos = cursor.fetchall()

            for d in datos:

                fecha = d[2] if len(d) > 2 and d[2] else "Sin fecha"

                self.textbox.insert(
                    "end",
                    f"👤 Usuario: {d[0]}\n"
                    f"💰 Pago: ${d[1]}\n"
                    f"📅 Fecha: {fecha}\n\n"
                    f"-----------------------------------------\n"
                )

            return
        cursor.execute(
            """
            SELECT DISTINCT nombre
            FROM usuarios
            WHERE nombre LIKE ?
            ORDER BY nombre
            """,
            (texto + "%",)
        )

        usuarios = cursor.fetchall()

        for usuario in usuarios:

            self.usuarios_mostrados.append(
                usuario[0]
            )

            self.lista.insert(
                "end",
                f"{usuario[0]}\n"
            )

    def seleccionar_usuario(self, event=None):

        try:

            indice = self.lista.index(
                "insert"
            ).split(".")[0]

            indice = int(
                indice
            ) - 1

            if indice < len(
                self.usuarios_mostrados
            ):

                nombre = self.usuarios_mostrados[
                    indice
                ]

                self.buscar.delete(
                    0,
                    "end"
                )

                self.buscar.insert(
                    0,
                    nombre
                )

                self.textbox.delete(
                    "1.0",
                    "end"
                )

                cursor.execute(
                    """
                    SELECT
                        usuarios.nombre,
                        pagos.monto,
                        pagos.fecha_pago
                    FROM pagos
                    JOIN usuarios
                    ON pagos.usuario_id = usuarios.id
                    WHERE usuarios.nombre = ?
                    ORDER BY pagos.id DESC
                    """,
                    (nombre,)
                )

                datos = cursor.fetchall()

                for d in datos:

                    fecha = d[2] if d[2] else "Sin fecha"

                    self.textbox.insert(
                        "end",
                        f"👤 Usuario: {d[0]}\n"
                        f"💰 Pago: ${d[1]}\n"
                        f"📅 Fecha: {fecha}\n\n"
                        f"-----------------------------------------\n"
                    )

        except:
            pass


from tkinter import filedialog, messagebox
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from datetime import datetime
import random
import os

# FACTURA PDF PRO
class VentanaFactura(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Factura")
        self.geometry("500x450")

        self.usuarios_mostrados = []

        self.nombre = ctk.CTkEntry(
            self,
            placeholder_text="Nombre usuario"
        )
        self.nombre.pack(pady=10)

        self.nombre.bind(
            "<KeyRelease>",
            self.buscar_sugerencias
        )

        self.lista = ctk.CTkTextbox(
            self,
            width=450,
            height=150
        )
        self.lista.pack(pady=10)

        self.lista.bind(
            "<ButtonRelease-1>",
            self.seleccionar_usuario
        )

        ctk.CTkButton(
            self,
            text="Generar PDF",
            command=self.generar
        ).pack(pady=20)

    def buscar_sugerencias(self, event=None):

        texto = self.nombre.get()

        self.lista.delete(
            "1.0",
            "end"
        )

        self.usuarios_mostrados = []

        if texto.strip() == "":
            return

        cursor.execute(
            "SELECT * FROM usuarios WHERE nombre LIKE ?",
            (texto + "%",)
        )

        usuarios = cursor.fetchall()

        for usuario in usuarios:

            self.usuarios_mostrados.append(
                usuario
            )

            self.lista.insert(
                "end",
                f"{usuario[1]} | Deuda: ${usuario[5]}\n"
            )

    def seleccionar_usuario(self, event=None):

        try:

            indice = self.lista.index(
                "insert"
            ).split(".")[0]

            indice = int(
                indice
            ) - 1

            if indice < len(
                self.usuarios_mostrados
            ):

                usuario = self.usuarios_mostrados[
                    indice
                ]

                self.nombre.delete(
                    0,
                    "end"
                )

                self.nombre.insert(
                    0,
                    usuario[1]
                )

        except:
            pass

    def generar(self):

        cursor.execute(
            "SELECT * FROM usuarios WHERE nombre=?",
            (self.nombre.get(),)
        )

        usuario = cursor.fetchone()

        if usuario:

            ruta = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[
                    (
                        "Archivos PDF",
                        "*.pdf"
                    )
                ],
                title="Guardar factura como",
                initialfile=f"Factura_{usuario[1]}.pdf"
            )

            if not ruta:
                return

            pdf = canvas.Canvas(ruta, pagesize=letter)

            # DATOS GENERALES
            fecha = datetime.now().strftime("%d/%m/%Y")
            factura_num = random.randint(1000, 9999)

                     # ===== ENCABEZADO =====
            pdf.setFont("Helvetica-Bold", 18)
            pdf.drawString(50, 750, "ACUAPAL")

            # ===== LOGO =====
            ruta_logo = ruta_recurso(
                "imagenes/logo.jpg"
            )

            if os.path.exists(ruta_logo):
                try:
                    logo = ImageReader(ruta_logo)
                    pdf.drawImage(logo, 450, 720, width=80, height=50)
                except Exception as e:
                    print("Error cargando logo:", e)
            else:
                print("Logo no encontrado en:", ruta_logo)

            # ===== CONTINÚA NORMAL =====
            pdf.setFont("Helvetica", 10)
            pdf.drawString(50, 735, "Servicio de Agua Potable")

          # Línea superior
            pdf.line(50, 725, 550, 725)
            # ===== INFO FACTURA =====
            pdf.setFont("Helvetica", 10)
            pdf.drawCentredString(306, 750, f"Factura N°: {factura_num}")
            
            pdf.drawCentredString(306, 735, f"Fecha: {fecha}")

            fecha_actual = datetime.now()

            # DEL 1 AL 5 MOSTRAR MES ANTERIOR
            if fecha_actual.day < 6:

                if fecha_actual.month == 1:

                    fecha_factura = fecha_actual.replace(
                        year=fecha_actual.year - 1,
                        month=12
                    )

                else:

                    fecha_factura = fecha_actual.replace(
                        month=fecha_actual.month - 1
                    )

            # DESDE EL 6 MOSTRAR MES ACTUAL
            else:

                fecha_factura = fecha_actual

            periodo = fecha_factura.strftime("%B %Y")
            pdf.setFont("Helvetica-Bold", 10)
            pdf.drawCentredString(306, 700, "Mes facturado:")

            pdf.setFont("Helvetica", 10)
            pdf.drawCentredString(306, 685, periodo)
            pdf.drawCentredString(306, 670, "pago del mes:$12000")
            # ===== DATOS CLIENTE =====
            pdf.setFont("Helvetica-Bold", 12)
            pdf.drawString(50, 700, "Datos del Cliente")
            
            pdf.setFont("Helvetica", 11)
            pdf.drawString(50, 680, f"Nombre: {usuario[1]}")
            pdf.drawString(50, 660, f"Dirección: {usuario[2]}")

            # Línea separadora
            pdf.line(50, 650, 550, 650)

            # ===== ESTADO DE CUENTA =====
            if usuario[5] < 0:

                pdf.setFont("Helvetica-Bold", 12)
                pdf.drawString(50, 620, "SALDO A FAVOR")

                pdf.setFont("Helvetica", 11)
                pdf.drawString(
                    50,
                    600,
                    f"Saldo a favor: ${abs(usuario[5])}"
                )

            else:

                pdf.setFont("Helvetica-Bold", 12)
                pdf.drawString(50, 620, "TOTAL A PAGAR")

                pdf.setFont("Helvetica", 11)
                pdf.drawString(
                    50,
                    600,
                    f"Valor a pagar: ${usuario[5]}"
                )

            pdf.setFont("Helvetica-Bold", 11)
            pdf.drawString(300, 615, "Estado de cuenta:")

            pdf.setFont("Helvetica", 11)
            pdf.drawString(300, 600, usuario[4])

            # Línea separadora
            pdf.line(50, 585, 550, 585)

            # ===== INFORMACIÓN DE PAGO =====
            pdf.setFont("Helvetica-Bold", 12)
            pdf.drawString(50, 555, "Información de Pago")

            pdf.setFont("Helvetica", 11)
            pdf.drawString(50, 535, "Nequi:")
            pdf.drawString(120, 535, "3215777800")  # CAMBIA ESTE NÚMERO
            pdf.drawString(50, 515, "Titular:")
            pdf.drawString(120, 515, "ACUAPAL")
            pdf.setFont("Helvetica-Bold", 11)
            pdf.drawString(300, 535, "WhatsApp:")

            pdf.setFont("Helvetica", 11)
            pdf.drawString(300, 520, "+57 3215777800")

            # ===== QR NEQUI =====
            ruta_qr = ruta_recurso(
                "imagenes/QR_nequi.png"
            )

            if os.path.exists(ruta_qr):

                pdf.setFont("Helvetica-Bold", 10)
                pdf.drawCentredString(
                    460,
                    430,
                    "Pague por Nequi"
                )

                pdf.drawImage(
                    ruta_qr,
                    360,
                    220,
                    width=180,
                    height=180,
                    preserveAspectRatio=True
                )
                pdf.setFont("Helvetica", 9)

                pdf.drawCentredString(
                    450,
                    205,
                    "3215777800"
                )

            # ===== MENSAJE FINAL =====
            pdf.setFont("Helvetica-Oblique", 10)
            pdf.drawString(50, 470, "Gracias por confiar en ACUAPAL...llevando alegria a tu hogar")
            pdf.setFont("Helvetica-Bold", 9)
            pdf.drawString(
                50,
                450,
                "OJO: Esta factura contiene la deuda total acumulada incluyendo el mes vigente"
            )

            # ===== PIE DE PÁGINA =====
            pdf.setFont("Helvetica", 8)
            pdf.drawString(50, 50, "Sistema generado automáticamente cualquier duda o inquietud comunicarse al WhatsApp de la empresa")

            # finalizar página PDF
            pdf.showPage()

            # guardar PDF
            pdf.save()

            messagebox.showinfo(
                "Éxito",
                "Factura profesional generada"
            )

        else:
            messagebox.showerror(
                "Error",
                "Usuario no encontrado"
            )


# SEGURIDAD
class VentanaSeguridad(ctk.CTkToplevel):
    def __init__(self, accion):
        super().__init__()

        self.accion = accion
        self.usuario_actual = None
        self.usuarios_mostrados = []

        self.title("Seguridad")
        self.geometry("450x500")

        self.nombre = ctk.CTkEntry(
            self,
            placeholder_text="Buscar usuario"
        )
        self.nombre.pack(pady=10)

        self.nombre.bind("<KeyRelease>", self.buscar_sugerencias)

        self.lista = ctk.CTkTextbox(
            self,
            width=350,
            height=120
        )
        self.lista.pack(pady=10)

        self.lista.bind(
            "<ButtonRelease-1>",
            self.seleccionar_usuario
        )

        if self.accion == "editar":
            self.valor = ctk.CTkEntry(
                self,
                placeholder_text="Nueva deuda"
            )
            self.valor.pack(pady=10)

        self.usuario_admin = ctk.CTkEntry(
            self,
            placeholder_text="Usuario admin"
        )
        self.usuario_admin.pack(pady=10)

        self.password_admin = ctk.CTkEntry(
            self,
            placeholder_text="Contraseña admin",
            show="*"
        )
        self.password_admin.pack(pady=10)

        ctk.CTkButton(
            self,
            text="Confirmar",
            command=self.confirmar
        ).pack(pady=20)

    def buscar_sugerencias(self, event=None):
        texto = self.nombre.get()

        self.lista.delete("1.0", "end")
        self.usuarios_mostrados = []

        if texto.strip() == "":
            return

        cursor.execute(
            "SELECT * FROM usuarios WHERE nombre LIKE ?",
            (texto + "%",)
        )

        usuarios = cursor.fetchall()

        for usuario in usuarios:
            self.usuarios_mostrados.append(usuario)

            self.lista.insert(
                "end",
                f"{usuario[1]}\n"
            )

    def seleccionar_usuario(self, event=None):
        try:
            indice = self.lista.index("insert").split(".")[0]
            indice = int(indice) - 1

            if indice < len(self.usuarios_mostrados):
                usuario = self.usuarios_mostrados[indice]

                self.usuario_actual = usuario

                self.nombre.delete(0, "end")
                self.nombre.insert(0, usuario[1])

        except:
            pass

    def confirmar(self):
        admin = verificar_seguridad(
            self.usuario_admin.get(),
            self.password_admin.get()
        )

        if not admin:
            messagebox.showerror(
                "Error",
                "Acceso denegado"
            )
            return

        if not self.usuario_actual:
            messagebox.showerror(
                "Error",
                "Selecciona un usuario"
            )
            return
        # Backup antes de cambios críticos
        crear_backup()

        if self.accion == "editar":
            try:
                nueva_deuda = int(self.valor.get())
            except ValueError:
                messagebox.showerror(
                    "Error",
                    "Ingrese un valor numérico válido"
                )
                return

            nuevo_estado = calcular_estado(nueva_deuda)

            cursor.execute(
                """
                UPDATE usuarios
                SET deuda=?, estado=?
                WHERE id=?
                """,
                (
                    nueva_deuda,
                    nuevo_estado,
                    self.usuario_actual[0]
                )
            )

        elif self.accion == "eliminar":
            cursor.execute(
                "DELETE FROM usuarios WHERE id=?",
                (self.usuario_actual[0],)
            )

        cursor.execute("""
        INSERT INTO auditoria
        (accion, usuario_afectado, fecha)
        VALUES (?, ?, ?)
        """, (
            self.accion,
            self.usuario_actual[1],
            str(datetime.now())
        ))

        conexion.commit()

        if self.accion == "editar":
            messagebox.showinfo(
                "Éxito",
                f"Deuda actualizada. Nuevo estado: {nuevo_estado}"
            )

        elif self.accion == "eliminar":
            messagebox.showinfo(
                "Éxito",
                "Usuario eliminado correctamente"
            )

        self.destroy()
# ==========================
# REGISTRAR GASTOS
# ==========================

class VentanaGastos(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Registrar Gasto")
        self.geometry("450x350")

        ctk.CTkLabel(
            self,
            text="Registro de Gastos",
            font=("Arial", 20, "bold")
        ).pack(pady=20)

        self.concepto = ctk.CTkEntry(
            self,
            placeholder_text="Concepto del gasto"
        )
        self.concepto.pack(pady=10)

        self.valor = ctk.CTkEntry(
            self,
            placeholder_text="Valor"
        )
        self.valor.pack(pady=10)

        ctk.CTkButton(
            self,
            text="Guardar Gasto",
            command=self.guardar_gasto
        ).pack(pady=20)

    def guardar_gasto(self):

        try:
            valor = int(self.valor.get())
        except:
            messagebox.showerror(
                "Error",
                "Ingrese un valor válido"
            )
            return

        fecha = datetime.now().strftime(
            "%d/%m/%Y %H:%M"
        )

        cursor.execute("""
        INSERT INTO gastos
        (concepto, valor, fecha)
        VALUES (?, ?, ?)
        """,
        (
            self.concepto.get(),
            valor,
            fecha
        ))

        conexion.commit()

        messagebox.showinfo(
            "Éxito",
            "Gasto registrado correctamente"
        )

        self.destroy()
# ==========================
# INVENTARIO FINANCIERO
# ==========================

class VentanaInventario(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Inventario Financiero")
        self.geometry("500x400")

        # TOTAL USUARIOS
        cursor.execute("""
        SELECT COUNT(*)
        FROM usuarios
        """)
        total_usuarios = cursor.fetchone()[0]
        # TOTAL RECAUDADO
        cursor.execute("""
        SELECT SUM(monto)
        FROM pagos 
        """)

        total_recaudado = cursor.fetchone()[0]

        if total_recaudado is None:
            total_recaudado = 0

        # TOTAL GASTOS
        cursor.execute("""
        SELECT SUM(valor)
        FROM gastos
        """)

        total_gastos = cursor.fetchone()[0]

        if total_gastos is None:
            total_gastos = 0

        # EFECTIVO
        efectivo = total_recaudado - total_gastos

        ctk.CTkLabel(
            self,
            text="📊 Inventario Financiero",
            font=("Arial", 22, "bold")
        ).pack(pady=20)
        ctk.CTkLabel(
            self,
            text=f"👥 Usuarios Registrados: {total_usuarios}",
            font=("Arial", 18)
        ).pack(pady=15)

        ctk.CTkLabel(
            self,
            text=f"💰 Total Recaudado: ${total_recaudado}",
            font=("Arial", 18)
        ).pack(pady=15)

        ctk.CTkLabel(
            self,
            text=f"📉 Total Gastos: ${total_gastos}",
            font=("Arial", 18)
        ).pack(pady=15)

        ctk.CTkLabel(
            self,
            text=f"🏦 Efectivo Disponible: ${efectivo}",
            font=("Arial", 20, "bold")
        ).pack(pady=20)
# ==========================
# HISTORIAL DE GASTOS
# ==========================

class VentanaHistorialGastos(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Historial de Gastos")
        self.geometry("800x600")

        self.textbox = ctk.CTkTextbox(
            self,
            width=700,
            height=500
        )
        self.textbox.pack(pady=20)

        cursor.execute("""
        SELECT concepto, valor, fecha
        FROM gastos
        ORDER BY id DESC
        """)

        gastos = cursor.fetchall()

        if not gastos:

            self.textbox.insert(
                "end",
                "No hay gastos registrados"
            )

        else:

            for gasto in gastos:

                self.textbox.insert(
                    "end",
                    f"📝 Concepto: {gasto[0]}\n"
                    f"💸 Valor: ${gasto[1]}\n"
                    f"📅 Fecha: {gasto[2]}\n\n"
                    f"--------------------------------------\n"
                )
class VentanaCierreDiario(ctk.CTkToplevel):
     def __init__(self):
        super().__init__()

        self.title("Cierre Diario")
        self.geometry("500x400")

        hoy = datetime.now().strftime("%d/%m/%Y")

        # TOTAL PAGOS DEL DÍA

        cursor.execute("""
        SELECT COUNT(*),
               COALESCE(SUM(monto),0)
        FROM pagos
        WHERE cerrado = 0
        AND fecha_pago LIKE ?
        """, (hoy + "%",))

        datos = cursor.fetchone()

        cantidad_pagos = datos[0]
        total_pagos = datos[1]

        # TOTAL GASTOS DEL DÍA
        cursor.execute("""
        SELECT COALESCE(SUM(valor),0)
        FROM gastos
        WHERE cerrado = 0
        AND fecha LIKE ?
        """, (hoy + "%",))
        total_gastos = cursor.fetchone()[0]

        efectivo = total_pagos - total_gastos

        ctk.CTkLabel(
            self,
            text="📊 Cierre Diario",
            font=("Arial", 22, "bold")
        ).pack(pady=20)

        ctk.CTkLabel(
            self,
            text=f"Fecha: {hoy}",
            font=("Arial", 18)
        ).pack(pady=10)

        ctk.CTkLabel(
            self,
            text=f"Pagos realizados: {cantidad_pagos}",
            font=("Arial", 18)
        ).pack(pady=10)

        ctk.CTkLabel(
            self,
            text=f"Total recaudado: ${total_pagos:,}",
            font=("Arial", 18)
        ).pack(pady=10)

        ctk.CTkLabel(
            self,
            text=f"Total gastos: ${total_gastos:,}",
            font=("Arial", 18)
        ).pack(pady=10)

        ctk.CTkLabel(
            self,
            text=f"Efectivo disponible: ${efectivo:,}",
            font=("Arial", 20, "bold")
        ).pack(pady=20)
        self.cantidad_pagos = cantidad_pagos
        self.total_pagos = total_pagos
        self.total_gastos = total_gastos
        self.efectivo = efectivo
        self.fecha_cierre = hoy

        ctk.CTkButton(
            self,
            text="💾 Registrar Cierre",
            fg_color="#2e7d32",
            hover_color="#388e3c",
            command=self.registrar_cierre
        ).pack(pady=15)
     def registrar_cierre(self):

        cursor.execute("""
        INSERT INTO cierres
        (
            fecha,
            cantidad_pagos,
            total_ingresos,
            total_gastos,
            efectivo
        )
        VALUES (?,?,?,?,?)
        """,
        (
            self.fecha_cierre,
            self.cantidad_pagos,
            self.total_pagos,
            self.total_gastos,
            self.efectivo
        ))

        conexion.commit()
        # MARCAR PAGOS COMO CERRADOS
        cursor.execute("""
        UPDATE pagos
        SET cerrado = 1
        WHERE cerrado = 0
        """)

        # MARCAR GASTOS COMO CERRADOS
        cursor.execute("""
        UPDATE gastos
        SET cerrado = 1
        WHERE cerrado = 0
        """)

        conexion.commit()
        crear_backup()

        messagebox.showinfo(
            "Éxito",
            "Cierre diario guardado correctamente"
        )
        self.destroy()
class VentanaHistorialCierres(ctk.CTkToplevel):
     def __init__(self):
        super().__init__()

        self.title("Historial de Cierres")
        self.geometry("800x500")

        self.texto = ctk.CTkTextbox(
            self,
            width=760,
            height=450
        )
        self.texto.pack(padx=10, pady=10)

        cursor.execute("""
        SELECT *
        FROM cierres
        ORDER BY id DESC
        """)

        cierres = cursor.fetchall()

        for cierre in cierres:

            self.texto.insert(
                "end",
                f"""
Fecha: {cierre[1]}
Pagos: {cierre[2]}
Ingresos: ${cierre[3]:,}
Gastos: ${cierre[4]:,}
Efectivo: ${cierre[5]:,}
----------------------------------------
"""
            )
class VentanaRestaurarBackup(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()

        self.title("Restaurar Backup")
        self.geometry("500x200")

        ctk.CTkLabel(
            self,
            text="Seleccionar Backup para Restaurar",
            font=("Arial", 18, "bold")
        ).pack(pady=20)

        ctk.CTkButton(
            self,
            text="📂 Seleccionar Backup",
            command=self.restaurar
        ).pack(pady=20)

    def restaurar(self):

        ruta = filedialog.askopenfilename(
            title="Seleccionar Backup",
            filetypes=[
                ("Base de datos SQLite", "*.db")
            ]
        )

        if not ruta:
            return

        respuesta = messagebox.askyesno(
            "Confirmar",
            "¿Desea restaurar este backup?\n\nSe reemplazará toda la información actual."
        )

        if not respuesta:
            return

        try:

            conexion.close()

            shutil.copy(
                ruta,
                "aguapal.db"
            )

            messagebox.showinfo(
                "Éxito",
                "Backup restaurado correctamente.\n\nReinicie el sistema."
            )

            self.destroy()

        except Exception as e:

            messagebox.showerror(
                "Error",
                str(e)
            )
# DASHBOARD
# DASHBOARD
class Dashboard(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("ACUAPAL")
        self.iconbitmap(
            "imagenes/logo.ico"
        )
        ancho = self.winfo_screenwidth()
        alto = self.winfo_screenheight()
        self.geometry(f"{ancho}x{alto}+0+0")

        # GRID
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # PANEL IZQUIERDO
        self.sidebar = ctk.CTkFrame(
            self,
            width=260,
            corner_radius=0,
            fg_color="#071a2f"
        )
        self.sidebar.grid(
            row=0,
            column=0,
            sticky="ns"
        )

        # LOGO SUPERIOR
        self.logo_img = ctk.CTkImage(
            light_image=Image.open("imagenes/logo.png"),

            size=(130, 130)
        )

        self.logo_label = ctk.CTkLabel(
            self.sidebar,
            image=self.logo_img,
            text=""
        )
        self.logo_label.pack(pady=20)
        self.icon_agregar_usuario = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/agregausuari.png"
            ),
            size=(22, 22)
        )
        self.icon_inicio = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/inicio.png"
            ),
            size=(22, 22)
        )
       
        self.icon_pago = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/pagos.png"
            ),
            size=(22, 22)
        )
        self.icon_gastos = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/gastos.png"
            ),
            size=(22, 22)
        )
        self.icon_inventario = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/inventario.png"
            ),
            size=(22, 22)
        )
        self.icon_factura = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/facturas.png"
            ),
            size=(22, 22)
        )
        self.icon_facturas_masivas = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/facturas masivas.png"
            ),
            size=(22, 22)
        )
        self.icon_facturas_WhatsApp = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/enviar_factura.png"
            ),
            size=(22, 22)
        )
        self.icon_historial_pago = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/historial_pago.png"
            ),
            size=(22, 22)
        )
        self.icon_configuracion = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/configuracion.png"
            ),
            size=(22, 22)
        )
        self.icon_cerrar = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/cerrarsecion.png"
            ),
            size=(20, 20)
        )
        self.icon_editar = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/editardeuda.png"
            ),
            size=(20, 20)
        )
        self.icon_eliminar = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/eliminarusuario.png"
            ),
            size=(20, 20)
        )
        self.icon_historial_gasto = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/historial_gasto.png"
            ),
            size=(20, 20)
        )

        self.icon_cierre_diario = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/cierre_diario.png"
            ),
            size=(20, 20)

        )
        self.icon_historial_cierres = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/historial_cierre.png"
            ),
            size=(20, 20)
        )
        self.icon_restaurar_backup = ctk.CTkImage(
            light_image=Image.open(
                "imagenes/restaurar_backup.png"
            ),
            size=(20, 20)
        )
        # BOTONES DEL MENÚ
        botones = [
            (" Inicio",VentanaInicio),
            (" Agregar Usuario", VentanaAgregarUsuario),
            (" Registrar Pago", VentanaRegistrarPago),
            (" Registrar Gasto", VentanaGastos),
            (" Inventario Financiero", VentanaInventario),
            (" Generar Factura", VentanaFactura),
            (" Generar Facturas Masivas", "facturas_masivas"),
            (" Enviar Facturas WhatsApp", "enviar_facturas"),
            (" Historial de Pagos", VentanaHistorial),
            (" Configuración", VentanaUsuarios),
           
        ]

        for texto, accion in botones:

            if accion == "editar":
                comando = lambda: self.abrir_seguridad("editar")

            elif accion == "eliminar":
                comando = lambda: self.abrir_seguridad("eliminar")

            elif accion == "facturas_masivas":
                comando = self.generar_facturas_masivas

            elif accion == "enviar_facturas":
                comando = self.enviar_facturas_whatsapp

            elif accion:
                comando = lambda v=accion: self.abrir(v)

            else:
                comando = None

            icono = None

            if texto == " Inicio":
                icono = self.icon_inicio

            elif texto == " Agregar Usuario":
                icono = self.icon_agregar_usuario

            elif texto == " Registrar Pago":
                icono = self.icon_pago

            elif texto == " Registrar Gasto":
                icono = self.icon_gastos
            elif texto == " Inventario Financiero":
                icono = self.icon_inventario

            elif texto == " Generar Factura":
                icono = self.icon_factura
            elif texto == " Generar Facturas Masivas":
                icono = self.icon_facturas_masivas 
            elif texto == " Enviar Facturas WhatsApp":
                icono = self.icon_facturas_WhatsApp
            elif texto == " Historial de Pagos":
                icono = self.icon_historial_pago
            elif texto == " Configuración":
                icono = self.icon_configuracion
            ctk.CTkButton(
                self.sidebar,
                text=texto,
                image=icono,
                compound="left",
                width=220,
                height=42,
                fg_color="#0d47a1",
                hover_color="#1565c0",
                corner_radius=12,
                command=comando
            ).pack(pady=5)

       # PANEL PRINCIPAL
        self.main = ctk.CTkFrame(
            self,
            fg_color="#0b0f19"
        )

        self.main.grid(
            row=0,
            column=1,
            sticky="nsew"
        )

        # BIENVENIDA
        self.bienvenida = ctk.CTkLabel(
            self.main,
            text="👋 ¡Bienvenido, Administrador!",
            font=("Arial", 30, "bold")
        )
        self.bienvenida.pack(pady=30)

        # CERRAR SESIÓN ARRIBA DERECHA
        self.logout = ctk.CTkButton(
            self.main,
            text="Cerrar Sesión",
            image=self.icon_cerrar,
            compound="left",
            width=140,
            height=35,
            fg_color="red",
            hover_color="#990000",
            command=self.cerrar_sistema
        )

        self.logout.place(
            relx=0.95,
            y=20,
            anchor="ne"
        )
        self.btn_editar = ctk.CTkButton(
            self.main,
            text=" Editar Deuda",
            image=self.icon_editar,
            compound="left",
            width=140,
            height=35,
            fg_color="#ff9800",
            command=lambda: self.abrir_seguridad("editar")
        )

        self.btn_editar.place(
            relx=0.95,
            y=70,
            anchor="ne"
        )

        self.btn_eliminar = ctk.CTkButton(
            self.main,
            text=" Eliminar Usuario",
            image=self.icon_eliminar,
            compound="left",
            width=140,
            height=35,
            fg_color="#d32f2f",
            command=lambda: self.abrir_seguridad("eliminar")
        )

        self.btn_eliminar.place(
            relx=0.95,
            y=120,
            anchor="ne"
        )
        self.btn_historial_gastos = ctk.CTkButton(
            self.main,
            text=" Historial Gastos",
            image=self.icon_historial_gasto,
            compound="left",
            width=140,
            height=35,
            fg_color="#0d47a1",
            hover_color="#1565c0",
            command=lambda: self.abrir(VentanaHistorialGastos)
        )

        self.btn_historial_gastos.place(
            relx=0.95,
            y=170,
            anchor="ne"
        )

        self.btn_cierre_diario = ctk.CTkButton(
            self.main,
            text=" Cierre Diario",
            image=self.icon_cierre_diario,
            compound="left",
            width=140,
            height=35,
            fg_color="#2e7d32",
            hover_color="#388e3c",
            command=lambda: self.abrir(
                VentanaCierreDiario
            )
        )

        self.btn_cierre_diario.place(
            relx=0.95,
            y=220,
            anchor="ne"
        )
        self.btn_historial_cierres = ctk.CTkButton(
            self.main,
            text=" Historial Cierres",
            image=self.icon_historial_cierres,
            compound="left",
            width=140,
            height=35,
            fg_color="#6a1b9a",
            hover_color="#8e24aa",
            command=lambda: self.abrir(
                VentanaHistorialCierres
            )
        )

        self.btn_historial_cierres.place(
            relx=0.95,
            y=270,
            anchor="ne"
        )
        self.btn_restaurar_backup = ctk.CTkButton(
            self.main,
            text=" Restaurar Backup",
            image=self.icon_restaurar_backup,
            compound="left",
            width=140,
            height=35,
            fg_color="#455a64",
            hover_color="#546e7a",
            command=lambda: self.abrir(
                VentanaRestaurarBackup
            )
        )

        self.btn_restaurar_backup.place(
            relx=0.95,
            y=320,
            anchor="ne"
        )

        # FECHA Y HORA
        fecha_actual = datetime.now().strftime(
            "%d/%m/%Y - %H:%M"
        )

        self.fecha = ctk.CTkLabel(
            self.main,
            text=fecha_actual,
            font=("Arial", 18)
        )
        self.fecha.pack()

        # LOGO CENTRAL GRANDE
        self.logo_center_img = ctk.CTkImage(
            light_image=Image.open("imagenes/logo.png"),
            size=(480, 480)
        )

        self.logo_center = ctk.CTkLabel(
            self.main,
            image=self.logo_center_img,
            text=""
        )
        self.logo_center.pack(expand=True)

    def abrir(self, ventana):
        v = ventana()
        v.grab_set()
        v.focus()
        v.lift()

    def abrir_seguridad(self, accion):
        v = VentanaSeguridad(accion)
        v.grab_set()
        v.focus()
        v.lift()
    def cerrar_sistema(self):

        crear_backup()

        conexion.commit()

        self.destroy()


     # GENERADOR DE FACTURAS MASIVAS
    def generar_facturas_masivas(self):

        if not os.path.exists("facturas"):
            os.makedirs("facturas")

        if not os.path.exists("historial_facturas"):
            os.makedirs("historial_facturas")

        cursor.execute("SELECT * FROM usuarios")
        usuarios = cursor.fetchall()

        if not usuarios:
            messagebox.showerror(
                "Error",
                "No hay usuarios registrados"
            )
            return

        facturas_subidas = 0

        fecha_control = datetime.now()

        if fecha_control.day < 6:

            if fecha_control.month == 1:

                fecha_mes = fecha_control.replace(
                    year=fecha_control.year - 1,
                    month=12
                )

            else:

                fecha_mes = fecha_control.replace(
                    month=fecha_control.month - 1
                )

        else:

            fecha_mes = fecha_control

        mes_actual = fecha_mes.strftime("%Y-%m")
        periodo = mes_actual
        # ==========================
        # CONTROL DE PERIODO FACTURAS
        # ==========================

        archivo_periodo = "facturas/periodo_actual.txt"

        if not os.path.exists(archivo_periodo):

            with open(archivo_periodo, "w", encoding="utf-8") as f:
                f.write(periodo)

        else:

            with open(archivo_periodo, "r", encoding="utf-8") as f:
                periodo_guardado = f.read().strip()

            if periodo_guardado != periodo:

                # MOVER FACTURAS AL HISTORIAL
                for archivo in os.listdir("facturas"):
                    if archivo == "periodo_actual.txt":
                        continue

                    if archivo.endswith(".pdf"):

                        nombre_sin_pdf = archivo[:-4]

                        nuevo_nombre = (
                            f"{nombre_sin_pdf}_"
                            f"{periodo_guardado}.pdf"
                        )

                        shutil.move(
                            os.path.join(
                                "facturas",
                                archivo
                            ),
                            os.path.join(
                                "historial_facturas",
                                nuevo_nombre
                            )
                        )
 
                # GUARDAR NUEVO PERIODO
                with open(
                    archivo_periodo,
                    "w",
                    encoding="utf-8"
                ) as f:

                    f.write(periodo)

        for usuario in usuarios:

            nombre_archivo = usuario[1].replace(" ", "_")
            nombre_cloud = nombre_archivo
            ruta = f"facturas/Factura_{nombre_archivo}.pdf"

            # =========================================
            # BUSCAR INFORMACIÓN ANTERIOR
            # =========================================

            cursor.execute("""
            SELECT mes_factura, cloud_id
            FROM usuarios
            WHERE id=?
            """, (usuario[0],))

            datos = cursor.fetchone()

            mes_guardado = None
            cloud_id_viejo = None

            if datos:
                mes_guardado = datos[0]
                cloud_id_viejo = datos[1]
            # =========================================
            # SI CAMBIÓ EL MES
            # =========================================

            if mes_guardado and mes_guardado != mes_actual:

                historial = (
                    f"historial_facturas/"
                    f"{nombre_cloud}_{mes_guardado}.pdf"
                )

                # mover factura vieja al historial
                ruta_vieja = f"facturas/Factura_{nombre_cloud}.pdf"

                if os.path.exists(ruta_vieja):

                    if not os.path.exists(historial):

                        shutil.move(
                            ruta_vieja,
                            historial
                        )

                # eliminar factura vieja de Cloudinary
                if cloud_id_viejo:

                    try:
                        eliminar_imagen(
                            cloud_id_viejo
                        )

                    except Exception as e:
                        print(
                            f"Error Cloudinary: {e}"
                        )

            # =========================================
            # CREAR NUEVO PDF
            # =========================================

            pdf = canvas.Canvas(
                ruta,
                pagesize=letter
            )

            fecha = datetime.now().strftime("%d/%m/%Y")

            factura_num = random.randint(
                1000,
                9999
            )

            fecha_actual = datetime.now()

            # DEL 1 AL 5 MOSTRAR MES ANTERIOR
            if fecha_actual.day < 6:

                if fecha_actual.month == 1:

                    fecha_factura = fecha_actual.replace(
                        year=fecha_actual.year - 1,
                        month=12
                    )

                else:

                    fecha_factura = fecha_actual.replace(
                        month=fecha_actual.month - 1
                    )

            # DESDE EL 6 MOSTRAR MES ACTUAL
            else:

                fecha_factura = fecha_actual

            periodo = fecha_factura.strftime(
                "%B %Y"
            )

            # ===== ENCABEZADO =====

            pdf.setFont("Helvetica-Bold", 18)
            pdf.drawString(
                50,
                750,
                "ACUAPAL"
            )

            # ===== LOGO =====

            ruta_logo = ruta_recurso(
                "imagenes/logo.jpg"
            )
            if os.path.exists(ruta_logo):

                try:

                    logo = ImageReader(
                        ruta_logo
                    )

                    pdf.drawImage(
                        logo,
                        450,
                        720,
                        width=80,
                        height=50
                    )

                except Exception as e:

                    print(
                        "Error cargando logo:",
                        e
                    )

            # ===== CONTINÚA NORMAL =====

            pdf.setFont(
                "Helvetica",
                10
            )

            pdf.drawString(
                50,
                735,
                "Servicio de Agua Potable"
            )

            pdf.line(
                50,
                725,
                550,
                725
            )

            # ===== INFO FACTURA =====

            pdf.setFont(
                "Helvetica",
                10
            )

            pdf.drawCentredString(
                306,
                750,
                f"Factura N°: {factura_num}"
            )

            pdf.drawCentredString(
                306,
                735,
                f"Fecha: {fecha}"
            )

            pdf.setFont(
                "Helvetica-Bold",
                10
            )

            pdf.drawCentredString(
                306,
                700,
                "Mes facturado:"
            )

            pdf.setFont(
                "Helvetica",
                10
            )

            pdf.drawCentredString(
                306,
                685,
                periodo
            )

            pdf.drawCentredString(
                306,
                670,
                "pago del mes:$12000"
            )

            # ===== DATOS CLIENTE =====

            pdf.setFont(
                "Helvetica-Bold",
                12
            )

            pdf.drawString(
                50,
                700,
                "Datos del Cliente"
            )

            pdf.setFont(
                "Helvetica",
                11
            )

            pdf.drawString(
                50,
                680,
                f"Nombre: {usuario[1]}"
            )

            pdf.drawString(
                50,
                660,
                f"Dirección: {usuario[2]}"
            )

            pdf.line(
                50,
                650,
                550,
                650
            )

            # ===== ESTADO DE CUENTA =====

            pdf.setFont(
                "Helvetica-Bold",
                12
            )

            if usuario[5] < 0:

                pdf.drawString(
                    50,
                    620,
                    "SALDO A FAVOR"
                )

                pdf.setFont(
                    "Helvetica",
                    11
                )

                pdf.drawString(
                    50,
                    600,
                    f"Saldo a favor: ${abs(usuario[5])}"
                )

            else:

                pdf.drawString(
                    50,
                    620,
                    "TOTAL A PAGAR"
                )

                pdf.setFont(
                    "Helvetica",
                    11
                )

                pdf.drawString(
                    50,
                    600,
                    f"Valor a pagar: ${usuario[5]}"
                )

            pdf.setFont(
                "Helvetica-Bold",
                11
            )

            pdf.drawString(
                300,
                615,
                "Estado de cuenta:"
            )

            pdf.setFont(
                "Helvetica",
                11
            )

            pdf.drawString(
                300,
                600,
                usuario[4]
            )

            # ===== INFORMACIÓN DE PAGO =====

            pdf.line(
                50,
                585,
                550,
                585
            )

            pdf.setFont(
                "Helvetica-Bold",
                12
            )

            pdf.drawString(
                50,
                555,
                "Información de Pago"
            )

            pdf.setFont(
                "Helvetica",
                11
            )

            pdf.drawString(
                50,
                535,
                "Nequi:"
            )

            pdf.drawString(
                120,
                535,
                "3215777800"
            )

            pdf.drawString(
                50,
                515,
                "Titular:"
            )

            pdf.drawString(
                120,
                515,
                "ACUAPAL"
            )

            pdf.setFont(
                "Helvetica-Bold",
                11
            )

            pdf.drawString(
                300,
                535,
                "WhatsApp:"
            )

            pdf.setFont(
                "Helvetica",
                11
            )

            pdf.drawString(
                300,
                520,
                "+57 3215777800"
            )
            # ===== QR NEQUI =====

            ruta_qr = ruta_recurso(
                "imagenes/QR_nequi.png"
            )

            if os.path.exists(ruta_qr):

                pdf.setFont(
                    "Helvetica-Bold",
                    10
                )

                pdf.drawCentredString(
                    450,
                    430,
                    "Pague por Nequi"
                )

                pdf.drawImage(
                    ruta_qr,
                    360,
                    220,
                    width=180,
                    height=180,
                    preserveAspectRatio=True
                )

                pdf.setFont(
                    "Helvetica",
                    9
                )

                pdf.drawCentredString(
                    450,
                    205,
                    "3215777800"
                )

            # ===== MENSAJE FINAL =====

            pdf.setFont(
                "Helvetica-Oblique",
                10
            )

            pdf.drawString(
                50,
                470,
                "Gracias por confiar en ACUAPAL...llevando alegria a tu hogar"
            )

            pdf.setFont(
                "Helvetica-Bold",
                9
            )

            pdf.drawString(
                50,
                450,
                "OJO: Esta factura contiene la deuda total acumulada incluyendo el mes vigente"
            )

            # ===== PIE DE PÁGINA =====

            pdf.setFont(
                "Helvetica",
                8
            )

            pdf.drawString(
                50,
                50,
                "Sistema generado automáticamente cualquier duda o inquietud comunicarse al WhatsApp de la empresa"
            )

            pdf.showPage()
            pdf.save()
            # CONVERTIR PDF A IMAGEN
            ruta_imagen = (
                f"facturas/Factura_{nombre_archivo}.jpg"
            )

            doc = fitz.open(ruta)

            pagina = doc.load_page(0)

            pix = pagina.get_pixmap(
                matrix=fitz.Matrix(2, 2)
            )

            pix.save(ruta_imagen)

            doc.close()

            # =========================================
            # SUBIR O REEMPLAZAR FACTURA
            # =========================================

            resultado = subir_imagen(
                ruta_imagen,
                nombre_cloud
            )

            # =========================================
            # ACTUALIZAR BASE DE DATOS
            # =========================================

            cursor.execute("""
            UPDATE usuarios
            SET
            link_factura=?,
            mes_factura=?,
            cloud_id=?
            WHERE id=?
            """,
            (
                resultado["url"],
                mes_actual,
                resultado["cloud_id"],
                usuario[0]
            ))

            

            facturas_subidas += 1

            print(
                f"Factura subida: {resultado['url']}"
            )

        conexion.commit()
        # =========================================
        # MENSAJE FINAL
        # =========================================

        messagebox.showinfo(
            "Proceso terminado",
            f"Se generaron y subieron "
            f"{facturas_subidas} facturas correctamente"
        )
    ## ENVIAR FACTURAS POR WHATSAPP DESKTOP

    def enviar_facturas_whatsapp(self):

        cursor.execute("SELECT * FROM usuarios")
        usuarios = cursor.fetchall()

        if not usuarios:
            messagebox.showerror(
                "Error",
                "No hay usuarios registrados"
            )
            return

        import pyautogui
        import pyperclip
        import time

        messagebox.showinfo(
            "WhatsApp",
            "Abre WhatsApp Desktop y presiona OK"
        )

        time.sleep(4)
        numeros_invalidos = []

        for usuario in usuarios:

            try:

                nombre = usuario[1]
                telefono = str(usuario[3]).strip()
                link = usuario[6] if len(usuario) > 6 else None
                # VALIDAR NÚMERO
                if not telefono.isdigit() or len(telefono) != 10:

                    numeros_invalidos.append(
                        f"{nombre} -> {telefono}"
                    )

                    print(f"Número inválido: {nombre}")

                    continue

                if not link:
                    print(f"{nombre} no tiene factura subida")
                    continue

                mensaje = f"""
    Hola {nombre},

    Aquí está tu factura del servicio de agua:

    {link}

    ACUAPAL
    """

                # BUSCAR CONTACTO
                pyautogui.hotkey("ctrl", "f")
                time.sleep(1)

                pyperclip.copy(telefono)
                pyautogui.hotkey("ctrl", "v")

                time.sleep(2)

                pyautogui.press("down")
                pyautogui.press("enter")

                time.sleep(2)

                # PEGAR MENSAJE
                pyperclip.copy(mensaje)
                pyautogui.hotkey("ctrl", "v")

                time.sleep(1)

                # ENVIAR
                pyautogui.press("enter")

                print(f"Factura enviada a {nombre}")

                time.sleep(4)

            except Exception as e:

                print(f"Error con {nombre}: {e}")

        # MOSTRAR RESUMEN FINAL

        if numeros_invalidos:

            mensaje_error = "\n".join(numeros_invalidos)

            messagebox.showwarning(
                "Números inválidos",
                f"Estos contactos tienen números inválidos:\n\n{mensaje_error}"
            )

        messagebox.showinfo(
            "Éxito",
            "Todas las facturas fueron enviadas correctamente"
        )
# LOGIN
class LoginApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Login AGUAPAL")
        self.geometry("400x300")
        self.iconbitmap(
            "imagenes/logo.ico"
        )
        self.usuario = ctk.CTkEntry(
            self,
            placeholder_text="Usuario"
        )
        self.usuario.pack(pady=10)

        self.password = ctk.CTkEntry(
            self,
            placeholder_text="Contraseña",
            show="*"
        )
        self.password.pack(pady=10)

        ctk.CTkButton(
            self,
            text="Ingresar",
            command=self.login
        ).pack(pady=20)

    def login(self):
        if verificar_admin(
            self.usuario.get(),
            self.password.get()
        ):
            self.destroy()
            Dashboard().mainloop()
        else:
            messagebox.showerror(
                "Error",
                "Credenciales incorrectas"
            )


app = LoginApp()
app.mainloop()