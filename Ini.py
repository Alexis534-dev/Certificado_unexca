import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import sqlite3
import os
import uuid
from datetime import datetime
import openpyxl
from PIL import Image, ImageDraw, ImageFont
import re

class UnexcaCertificateSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Certificados UNEXCA")
        self.root.geometry("1200x800")
        self.root.configure(bg='#F0F4F8')

        # Crear directorios necesarios
        self.crear_directorios()

        # Configuración de estilos
        self.configurar_estilos()

        # Inicializar base de datos
        self.inicializar_base_datos()

        # Crear interfaz principal
        self.crear_interfaz_principal()

    def crear_directorios(self):
        directorios = [
            "bases_datos", 
            "certificados", 
            "templates", 
            "fonts", 
            "logs",
            "exports"
        ]
        for directorio in directorios:
            os.makedirs(directorio, exist_ok=True)

    def configurar_estilos(self):
        self.colores = {
            'fondo_principal': '#F0F4F8',
            'texto_principal': '#2C3E50',
            'boton_primario': '#3498DB',
            'boton_secundario': '#2ECC71'
        }

        self.fuentes = {
            'titulo_grande': ('Helvetica', 24, 'bold'),
            'texto_normal': ('Arial', 12),
            'texto_pequeño': ('Arial', 10)
        }

        estilo = ttk.Style()
        estilo.theme_use('clam')
        
        estilo.configure('primary.TButton', 
                         background=self.colores['boton_primario'], 
                         foreground='white', 
                         font=self.fuentes['texto_normal'])
        
        estilo.configure('secondary.TButton', 
                         background=self.colores['boton_secundario'], 
                         foreground='white', 
                         font=self.fuentes['texto_normal'])

    def inicializar_base_datos(self):
        db_path = os.path.join('bases_datos', 'unexca_certificados.db')
        
        try:
            self.conn = sqlite3.connect(db_path)
            self.cursor = self.conn.cursor()

            self.cursor.executescript('''
                CREATE TABLE IF NOT EXISTS estudiantes (
                    id TEXT PRIMARY KEY,
                    nombre TEXT NOT NULL,
                    apellido TEXT NOT NULL,
                    cedula TEXT UNIQUE,
                    email TEXT,
                    fecha_registro DATETIME DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS cursos (
                    id TEXT PRIMARY KEY,
                    nombre TEXT NOT NULL,
                    codigo TEXT UNIQUE,
                    area TEXT,
                    duracion TEXT,
                    descripcion TEXT,
                    instructor TEXT,
                    fecha_creacion DATETIME DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS certificados (
                    id TEXT PRIMARY KEY,
                    estudiante_id TEXT,
                    curso_id TEXT,
                    fecha_emision DATETIME,
                    archivo_certificado TEXT,
                    FOREIGN KEY(estudiante_id) REFERENCES estudiantes(id),
                    FOREIGN KEY(curso_id) REFERENCES cursos(id)
                );
            ''')
            
            self.conn.commit()
            
        except sqlite3.Error as e:
            messagebox.showerror("Error de Base de Datos", f"No se pudo inicializar la base de datos: {e}")
            raise

    def crear_interfaz_principal(self):
        frame_principal = ttk.Frame(self.root)
        frame_principal.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        titulo = ttk.Label(
            frame_principal, 
            text="Sistema de Certificados UNEXCA",
            font=self.fuentes['titulo_grande'],
            foreground=self.colores['texto_principal']
        )
        titulo.pack(pady=20)

        modulos = [
            ("Gestión de Estudiantes", self.abrir_gestion_estudiantes),
            ("Gestión de Cursos", self.abrir_gestion_cursos),
            ("Generación de Certificados", self.abrir_generacion_certificados)
        ]

        for texto, comando in modulos:
            boton = ttk.Button(
                frame_principal, 
                text=texto, 
                command=comando,
                style='primary.TButton',
                width=40
            )
            boton.pack(pady=10)

    def abrir_gestion_estudiantes(self):
        ventana_estudiantes = tk.Toplevel(self.root)
        ventana_estudiantes.title("Gestión de Estudiantes")
        ventana_estudiantes.geometry("1000x600")

        # Frame para entrada de datos
        frame_entrada = ttk.Frame(ventana_estudiantes)
        frame_entrada.pack(pady=10, padx=20, fill=tk.X)

        # Campos de entrada
        campos = [
            ("Nombre", "nombre"),
            ("Apellido", "apellido"),
            ("Cédula", "cedula"),
            ("Email", "email")
        ]

        entradas = {}
        for i, (label, campo) in enumerate(campos):
            ttk.Label(frame_entrada, text=label).grid(row=i//2, column=(i%2)*2, padx=5, pady=5)
            entrada = ttk.Entry(frame_entrada, width=30)
            entrada.grid(row=i//2, column=(i%2)*2+1, padx=5, pady=5)
            entradas[campo] = entrada

        # Botones de acción
        frame_botones = ttk.Frame(ventana_estudiantes)
        frame_botones.pack(pady=10)

        botones = [
            ("Registrar", lambda: self.registrar_estudiante(entradas)),
            ("Cargar Estudiantes", lambda: self.cargar_estudiantes()),
            ("Importar Excel", self.importar_estudiantes_excel)
        ]

        for texto, comando in botones:
            ttk.Button(frame_botones, text=texto, command=comando).pack(side=tk.LEFT, padx=5)

        # Tabla de estudiantes
        columns = ("ID", "Nombre", "Apellido", "Cédula", "Email")
        tabla = ttk.Treeview(ventana_estudiantes, columns=columns, show="headings")
        
        for col in columns:
            tabla.heading(col, text=col)
            tabla.column(col, width=150, anchor=tk.CENTER)
        
        tabla.pack(expand=True, fill=tk.BOTH, padx=20, pady=10)

        # Cargar estudiantes iniciales
        self.cargar_estudiantes(tabla)

    def registrar_estudiante(self, entradas):
        # Validaciones básicas
        datos = {}
        for campo, entrada in entradas.items():
            valor = entrada.get().strip()
            if not valor:
                messagebox.showerror("Error", f"El campo {campo} no puede estar vacío")
                return
            datos[campo] = valor

        # Validar email
        if not re.match(r"[^@]+@[^@]+\.[^@]+", datos['email']):
            messagebox.showerror("Error", "Email inválido")
            return

        # Generar ID único
        estudiante_id = str(uuid.uuid4())
        
        try:
            self.cursor.execute('''
                INSERT INTO estudiantes (id, nombre, apellido, cedula, email) 
                VALUES (?, ?, ?, ?, ?)
            ''', (estudiante_id, datos['nombre'], datos['apellido'], 
                  datos['cedula'], datos['email']))
            
            self.conn.commit()
            messagebox.showinfo("Éxito", "Estudiante registrado correctamente")
            
            # Limpiar entradas
            for entrada in entradas.values():
                entrada.delete(0, tk.END)
        
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", "Ya existe un estudiante con esta cédula")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo registrar el estudiante: {e}")

    def cargar_estudiantes(self, tabla=None):
        try:
            self.cursor.execute("SELECT * FROM estudiantes")
            estudiantes = self.cursor.fetchall()

            if tabla:
                # Limpiar tabla
                for i in tabla.get_children():
                    tabla.delete(i)
                
                # Insertar estudiantes
                for estudiante in estudiantes:
                    tabla.insert("", "end", values=estudiante)
            
            return estudiantes
        
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los estudiantes: {e}")

    def importar_estudiantes_excel(self):
        archivo = filedialog.askopenfilename(
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        
        if not archivo:
            return

        try:
            wb = openpyxl.load_workbook(archivo)
            hoja = wb.active

            for fila in hoja.iter_rows(min_row=2, values_only=True):
                if len(fila) >= 4:
                    nombre, apellido, cedula, email = fila[:4]
                    
                    estudiante_id = str(uuid.uuid4())
                    
                    try:
                        self.cursor.execute('''
                            INSERT INTO estudiantes (id, nombre, apellido, cedula, email) 
                            VALUES (?, ?, ?, ?, ?)
                        ''', (estudiante_id, nombre, apellido, cedula, email))
                    except sqlite3.IntegrityError:
                        # Omitir registros duplicados
                        continue

            self.conn.commit()
            messagebox.showinfo("Éxito", "Estudiantes importados correctamente")
        
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron importar los estudiantes: {e}")

    def abrir_gestion_cursos(self):
        # Implementación similar a gestión de estudiantes
        pass

    def abrir_generacion_certificados(self):
        # Implementación del método de generación de certificados
        pass

    def __del__(self):
        if hasattr(self, 'conn'):
            self.conn.close()

def main():
    root = tk.Tk()
    app = UnexcaCertificateSystem(root)
    root.mainloop()

if __name__ == "__main__":
    main()
