"""
Convertidor de Documentos - EasyDoc Suite V 1.1
Estilo Retro/Cyberpunk basado en Figma (Soporte Multiformato)
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import threading
import os
import sys
from pathlib import Path
from PIL import Image, ImageTk, ImageSequence
import time


# --- NUEVA LIBRERÍA PARA PDF A WORD ---
try:
    from pdf2docx import Converter as PDFConverter
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pdf2docx"])
    from pdf2docx import Converter as PDFConverter

def obtener_ruta_recurso(ruta_relativa):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, ruta_relativa)

try:
    import win32com.client
    import pythoncom
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
    import win32com.client
    import pythoncom

# --- LÓGICA DE CONVERSIÓN MULTIFORMATO ---
def convertir_lote(archivos, carpeta_destino, modo, callback_estado, callback_error):
    app_com = None
    try:
        # Inicializar el motor COM si es necesario
        if modo == "Word -> PDF":
            app_com = win32com.client.Dispatch("Word.Application")
            app_com.Visible = False
            app_com.DisplayAlerts = False
        elif modo == "Excel -> PDF":
            app_com = win32com.client.Dispatch("Excel.Application")
            app_com.Visible = False
            app_com.DisplayAlerts = False
        elif modo == "PowerPoint -> PDF":
            app_com = win32com.client.Dispatch("PowerPoint.Application")

        for i, archivo in enumerate(archivos):
            nombre = Path(archivo).name
            callback_estado(f"Procesando: {nombre}")
            
            time.sleep(0.1)

            try:
                archivo_abs = str(Path(archivo).resolve())
                
                # Determinar extensión de salida
                ext_dest = ".docx" if modo == "PDF -> Word" else ".pdf"

                # Determinar ruta de salida
                if carpeta_destino:
                    dest_path = os.path.join(carpeta_destino, Path(archivo).stem + ext_dest)
                else:
                    dest_path = str(Path(archivo).with_suffix(ext_dest))
                
                dest_abs = str(Path(dest_path).resolve())

                # Ejecutar la conversión según el modo
                if modo == "Word -> PDF":
                    doc = app_com.Documents.Open(archivo_abs)
                    doc.SaveAs(dest_abs, FileFormat=17) # 17 = wdFormatPDF
                    doc.Close(False)
                
                elif modo == "Excel -> PDF":
                    wb = app_com.Workbooks.Open(archivo_abs)
                    wb.ExportAsFixedFormat(0, dest_abs) # 0 = xlTypePDF
                    wb.Close(False)
                
                elif modo == "PowerPoint -> PDF":
                    prs = app_com.Presentations.Open(archivo_abs, WithWindow=False)
                    prs.SaveAs(dest_abs, 32) # 32 = ppSaveAsPDF
                    prs.Close()
                
                elif modo == "PDF -> Word":
                    # Usando pdf2docx
                    cv = PDFConverter(archivo_abs)
                    cv.convert(dest_abs, start=0, end=None)
                    cv.close()

            except Exception as e:
                callback_error(nombre, str(e))

                time.sleep(0.1)
                
    finally:
        # Limpieza de procesos en segundo plano
        if app_com:
            try:
                app_com.Quit()
            except:
                pass

class TkinterDnD_CTk(ctk.CTk, TkinterDnD.DnDWrapper):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)

# --- PALETA DE COLORES ---
BG_COLOR = "#150A21"         
PANEL_COLOR = "#2A1B4E"      
ACCENT_PURPLE = "#8A2BE2"    
NEON_GREEN = "#9DFF00"       
RED_BTN = "#E53935"          
GRAY_BTN = "#78909C"         
TEXT_COLOR = "#E0E0E0"
FONT_RETRO = ("Consolas", 12, "bold") 

ctk.set_appearance_mode("dark") 

class ConversorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("EasyDoc Suite")
        self.root.geometry("850x650")
        self.root.configure(fg_color=BG_COLOR)
        self.root.resizable(False, False)
        
        self.root.iconbitmap(obtener_ruta_recurso("icono.ico"))
        
        self.archivos = []
        self.errores = []
        self.carpeta_destino = None
        self.build_ui()
        self.cargar_gif()

    def cargar_gif(self):
        self.gif_frames = []
        self.animating = False
        self.frame_index = 0
        try:
            ruta_gif = obtener_ruta_recurso("loading.gif")
            img = Image.open(ruta_gif)
            for img_frame in ImageSequence.Iterator(img):
                img_frame = img_frame.convert("RGBA")
                img_frame = img_frame.resize((350, 250), Image.Resampling.LANCZOS)
                self.gif_frames.append(ImageTk.PhotoImage(img_frame))
        except Exception as e:
            print("Aviso: No se encontró loading.gif", e)
        self.gif_label = tk.Label(self.frame_lista, bg=PANEL_COLOR, bd=0)

    def animar_gif(self):
        if not self.animating or not self.gif_frames: return
        self.gif_label.config(image=self.gif_frames[self.frame_index])
        self.frame_index = (self.frame_index + 1) % len(self.gif_frames)
        self.root.after(50, self.animar_gif)

    def iniciar_animacion(self):
        if self.gif_frames:
            self.animating = True
            self.gif_label.place(relx=0.5, rely=0.5, anchor="center")
            self.animar_gif()

    def detener_animacion(self):
        self.animating = False
        self.gif_label.place_forget()

    def build_ui(self):
        header_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        header_frame.pack(fill="x", padx=20, pady=(15, 5))
        ctk.CTkLabel(header_frame, text="EASYDOC SUITE", font=("Consolas", 18, "bold"), text_color=TEXT_COLOR).pack(anchor="w")
        ctk.CTkLabel(header_frame, text="Conversor de Documentos", font=("Consolas", 10), text_color="#A991D4").pack(anchor="w")
        ctk.CTkLabel(header_frame, text="Desarrollado por Peter Güette", font=("Consolas", 10), text_color="#A991D4").pack(anchor="w")
        ctk.CTkFrame(self.root, height=2, fg_color=ACCENT_PURPLE).pack(fill="x", pady=5)
        ctk.CTkLabel(self.root, text="> Arrastra tus documentos aqui o usa los controles <", text_color=TEXT_COLOR, font=FONT_RETRO).pack(pady=10)

        self.frame_lista = ctk.CTkFrame(self.root, fg_color=PANEL_COLOR, border_color=ACCENT_PURPLE, border_width=2, corner_radius=0)
        self.frame_lista.pack(fill="both", expand=True, padx=30, pady=5)
        scrollbar = ctk.CTkScrollbar(self.frame_lista, button_color=ACCENT_PURPLE, button_hover_color="#6A1CB2")
        scrollbar.pack(side="right", fill="y", pady=2, padx=2)

        self.lista = tk.Listbox(self.frame_lista, yscrollcommand=scrollbar.set, font=("Consolas", 11), selectmode="extended", bg=PANEL_COLOR, fg=TEXT_COLOR, bd=0, highlightthickness=0, activestyle="none", selectbackground=ACCENT_PURPLE)
        self.lista.pack(fill="both", expand=True, padx=10, pady=10)
        scrollbar.configure(command=self.lista.yview)

        self.lista.insert("end", "      [ Arrastrar archivos aquí ]")
        self.lista.insert("end", "      o usa los botones abajo")
        self.lista.drop_target_register(DND_FILES)
        self.lista.dnd_bind('<<Drop>>', self.al_soltar)

        btn_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        btn_frame.pack(pady=15)
        ctk.CTkButton(btn_frame, text="[+] AGREGAR", width=120, fg_color=ACCENT_PURPLE, hover_color="#6A1CB2", corner_radius=2, font=FONT_RETRO, text_color="white", command=self.agregar).grid(row=0, column=0, padx=10)
        ctk.CTkButton(btn_frame, text="[-] QUITAR", width=120, fg_color=RED_BTN, hover_color="#B71C1C", corner_radius=2, font=FONT_RETRO, text_color="white", command=self.eliminar).grid(row=0, column=1, padx=10)
        ctk.CTkButton(btn_frame, text="[x] LIMPIAR", width=120, fg_color=GRAY_BTN, hover_color="#546E7A", corner_radius=2, font=FONT_RETRO, text_color="white", command=self.limpiar).grid(row=0, column=2, padx=10)

        controls_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        controls_frame.pack(fill="x", padx=30)
        self.destino_var = ctk.StringVar(value="Carpeta original")
        ctk.CTkEntry(controls_frame, textvariable=self.destino_var, state="readonly", width=500, fg_color=BG_COLOR, border_color=ACCENT_PURPLE, corner_radius=0, font=("Consolas", 10), text_color=TEXT_COLOR).pack(side="left", padx=(0, 10))
        ctk.CTkButton(controls_frame, text="📁 Destino", width=100, fg_color=ACCENT_PURPLE, hover_color="#6A1CB2", corner_radius=2, font=FONT_RETRO, command=self.elegir_destino).pack(side="left")

        motor_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        motor_frame.pack(pady=(15, 5))
        ctk.CTkLabel(motor_frame, text="Tipo de conversión:", text_color=TEXT_COLOR, font=FONT_RETRO).pack()
        
        # --- OPCIONES DE CONVERSIÓN ---
        self.opciones_motor = ["Word -> PDF", "PDF -> Word", "Excel -> PDF", "PowerPoint -> PDF"]
        self.tipo_conversion_var = ctk.StringVar(value=self.opciones_motor[0])
        ctk.CTkOptionMenu(
            motor_frame, 
            values=self.opciones_motor, 
            variable=self.tipo_conversion_var, 
            fg_color=BG_COLOR, 
            button_color=ACCENT_PURPLE, 
            button_hover_color="#6A1CB2", 
            dropdown_fg_color=PANEL_COLOR, 
            corner_radius=0, 
            font=FONT_RETRO,
            command=self.cambio_motor # Limpia la lista si cambias de modo
        ).pack(pady=5)

        self.estado_lbl = ctk.CTkLabel(self.root, text="Listo para procesar", text_color="#A991D4", font=("Consolas", 10))
        self.estado_lbl.pack(pady=(5, 10))

        self.btn_conv = ctk.CTkButton(self.root, text="⚡ INICIAR CONVERSION ⚡", font=("Consolas", 14, "bold"), height=45, width=300, fg_color=NEON_GREEN, hover_color="#7BCC00", text_color="black", corner_radius=2, border_color="white", border_width=1, command=self.iniciar_conversion)
        self.btn_conv.pack(pady=(0, 20))

    # --- LÓGICA DE FILTROS SEGÚN EL MOTOR ---
    def obtener_filtros_archivos(self):
        modo = self.tipo_conversion_var.get()
        if modo == "Word -> PDF":
            return [("Archivos de Word", "*.docx *.doc")], ('.docx', '.doc')
        elif modo == "Excel -> PDF":
            return [("Archivos de Excel", "*.xlsx *.xls")], ('.xlsx', '.xls')
        elif modo == "PowerPoint -> PDF":
            return [("Archivos de PowerPoint", "*.pptx *.ppt")], ('.pptx', '.ppt')
        elif modo == "PDF -> Word":
            return [("Archivos PDF", "*.pdf")], ('.pdf',)
        return [("Todos los archivos", "*.*")], ()

    def cambio_motor(self, seleccion):
        """Si el usuario cambia el tipo de conversión, se vacía la lista por seguridad"""
        if self.archivos:
            self.limpiar()
            self.estado_lbl.configure(text=f"Modo cambiado a {seleccion}. Lista limpiada.")

    def chequear_placeholder(self):
        if self.lista.size() > 0 and "[ Arrastrar archivos aquí ]" in self.lista.get(0):
            self.lista.delete(0, "end")

    def al_soltar(self, event):
        self.chequear_placeholder()
        archivos = self.root.tk.splitlist(event.data)
        _, extensiones_validas = self.obtener_filtros_archivos()
        
        for f in archivos:
            if f.lower().endswith(extensiones_validas):
                if f not in self.archivos:
                    self.archivos.append(f)
                    self.lista.insert("end", f"> {Path(f).name}")

    def agregar(self):
        self.chequear_placeholder()
        filtros_dialogo, extensiones_validas = self.obtener_filtros_archivos()
        files = filedialog.askopenfilenames(filetypes=filtros_dialogo)
        for f in files:
            # Re-verificar por si acaso en algunos SO el filtro falla
            if f.lower().endswith(extensiones_validas) and f not in self.archivos:
                self.archivos.append(f)
                self.lista.insert("end", f"> {Path(f).name}")

    def eliminar(self):
        sel = list(self.lista.curselection())[::-1]
        for i in sel:
            self.lista.delete(i)
            self.archivos.pop(i)
        if not self.archivos:
            self.limpiar()

    def limpiar(self):
        self.lista.delete(0, "end")
        self.archivos.clear()
        self.lista.insert("end", "      [ Arrastrar archivos aquí ]")
        self.lista.insert("end", "      o usa los botones abajo")

    def elegir_destino(self):
        folder = filedialog.askdirectory()
        if folder:
            self.carpeta_destino = folder
            self.destino_var.set(folder)

    def iniciar_conversion(self):
        if not self.archivos: return
        self.btn_conv.configure(state="disabled")
        self.iniciar_animacion()
        
        modo_actual = self.tipo_conversion_var.get()
        threading.Thread(target=self._hilo, args=(modo_actual,), daemon=True).start()

    def _hilo(self, modo):
        # 1. Inicializar el motor COM para este hilo secundario (Vital para que no se congele)
        pythoncom.CoInitialize()
        
        total = len(self.archivos)
        
        # 2. Truco de magia: Enviar la actualización de texto AL HILO PRINCIPAL de forma segura
        def actualizar_estado(texto):
            self.root.after(0, lambda: self.estado_lbl.configure(text=texto))
        
        # Enviar los errores también de forma segura
        def registrar_error(n, m):
            self.errores.append(f"{n}: {m}")
        
        # 3. Ejecutar la conversión
        convertir_lote(self.archivos, self.carpeta_destino, modo, actualizar_estado, registrar_error)
        
        # 4. Finalizar la UI de forma segura en el hilo principal
        def finalizar_ui():
            self.detener_animacion()
            self.btn_conv.configure(state="normal")
            self.estado_lbl.configure(text="> CONVERSION COMPLETADA <")
            messagebox.showinfo("Éxito", f"Se procesaron {total} archivos correctamente.")

        self.root.after(0, finalizar_ui)
        
        # 5. Cerrar el motor COM de este hilo
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    root = TkinterDnD_CTk()
    app = ConversorApp(root)
    root.mainloop()