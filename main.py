import csv
import os


import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename, askdirectory
from tkinter.colorchooser import askcolor
from tkinter import messagebox
import tkinter.font as tkFont


class VentanaPrincipal(tk.Frame):
    def __init__(self, parent=None):
        tk.Frame.__init__(self, parent)
        self.parent = parent
        self.yaGuardo = True
        self.init_ui()

    def init_ui(self):
        # Configuración de ventana
        # Medidas de la pantalla
        
        self.a = self.parent.winfo_screenwidth()
        self.h = self.parent.winfo_screenheight()
        print(f"a: {self.a}")
        print(f"h: {self.h}")

        # Para manejar el evento cuando se cierre la ventana
        self.parent.protocol("WM_DELETE_WINDOW", self.cerrandoVentana)

        # Muestra la ventana maximizada
        try:
            self.parent.state('zoomed')
        except:
            self.parent.state('normal')

        # Título de la ventana
        self.parent.title("Lee y mueve")

        # Medidas minimas a las que se puede ajustar la ventana
        self.parent.minsize(400, 400)

        # Icono de la ventana
        try:
            self.parent.iconbitmap('Logo_Bienvenida_IMP_PREDICT_V2.ico')
            self.anchoLienzo = self.a * 0.65
        except:
            self.anchoLienzo = self.a * 0.65

        self.altoLienzo = self.h
        ancho = 400
        alto = 400
        posx = int((self.a - ancho) / 2)
        posy = int((self.h - alto) / 2)
        self.parent.geometry(str(ancho) + "x" + str(alto) + "+" + str(posx) + "+" + str(posy))
        self.parent.attributes('-alpha', 0.0)
        self.parent.resizable(0, 0)
        # Se crea un menu
        self.menubar = tk.Menu(self.parent)
        # Se añade a la ventana principal
        self.parent.config(menu=self.menubar)

        # Tearoff=0 es para definir la posición del primer menú que se creara
        # y posteriormente los demas menús tendran un consecutivo de 0
        # Se añade un nuevo submenu a la barra de menus
        self.fileMenu = tk.Menu(self.menubar, tearoff=0)


        # Para que se añadan como submenus de file los comandos antes mencionados
        self.menubar.add_cascade(label="Archivo", menu=self.fileMenu)

        # añadida de submenu
        self.submenu = tk.Menu(self.fileMenu, tearoff=0)
        self.fileMenu.add_command(label="Nuevo", 
                            command=self.abrirNuevo,
                            )

        # Separador de menus
        self.fileMenu.add_separator()
        self.fileMenu.add_command(label="Salir", 
                                command=self.cerrandoVentana, 
                                accelerator="Ctrl+s")

        self.menubar.add_cascade(label="Acerca",  accelerator="Ctrl+d")
        self.menubar.add_cascade(label="Ayuda", accelerator="Ctrl+y")
        
        # Bineo a atajos de teclado
        self.parent.bind("<Control-s>", self.cerrandoVentanaEvent)

    def cerrandoVentanaEvent(self, event):
        self.cerrandoVentana()


   
    """Función para evitar cerrar abruptamente el programa"""

    def cerrandoVentana(self):

        respuestaDeCierre = messagebox.askyesno("Salir de Lee y Mueve", 
                                                "¿Desea realmente salir de Lee y mueve")
        if respuestaDeCierre == True:
            self.parent.destroy()

    def abrirNuevo(self):
        ventanaRutas(self.parent)

    def get_curr_screen_geometry(self):
        """
        Workaround to get the size of the current screen in a multi-screen setup.

        Returns:
            geometry (str): The standard Tk geometry string.
                [width]x[height]+[left]+[top]
        """
        root = tk.Tk()
        root.update_idletasks()
        root.attributes('-fullscreen', True)
        root.state('iconic')
        geometry = root.winfo_geometry()
        root.destroy()
        return geometry

class ventanaRutas:
    def __init__(self, parent):
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        a = self.parent.winfo_screenwidth()
        h = self.parent.winfo_screenheight()
        print(self.parent.winfo_reqheight())
        print(self.parent.winfo_reqwidth())
        print(self.parent.winfo_height())
        print(self.parent.winfo_width())
        


        ancho = 670
        alto = 310
        posx = int((a - ancho) / 2)
        posy = int((h - alto) / 2)
        self.ventaDatosRuta = tk.Toplevel(self.parent)
        self.ventaDatosRuta.geometry(str(ancho) + "x" + str(alto) + "+" + str(posx) + "+" + str(posy))
        self.ventaDatosRuta.attributes('-alpha', 0.0)
        self.ventaDatosRuta.resizable(0, 0)
        try:
            self.ventaDatosRuta.iconbitmap('Logo_Bienvenida_IMP_PREDICT_V2.ico')
            self.ventaDatosRuta.title("Datos del cubo sísmico")
        except:
            self.ventaDatosRuta.title("Datos cubo sísmico")
        # Configuración widgets
        self.miFuente = tkFont.Font(size=10)
        self.etiquetaArchivo = tk.Label(self.ventaDatosRuta, 
                                        text="Seleccionar el archivo de lectura de datos")
        self.campoArchivo = tk.Entry(self.ventaDatosRuta, width=50)
        self.btnExaminar = tk.Button(self.ventaDatosRuta, 
                                    text="Examinar", 
                                    command=self.abrirArchivo, 
                                    width=16,
                                    height=1,
                                    anchor="center", 
                                    justify="center", 
                                    relief="groove")
        self.etiquetaBusqueda = tk.Label(self.ventaDatosRuta, 
                                        text="Seleccionar la carpeta de busqueda de archivos")
        self.campoRutaBusqueda = tk.Entry(self.ventaDatosRuta, width=50)
        self.btnExaminarUbicacion= tk.Button(self.ventaDatosRuta, 
                                            text="Examinar", 
                                            command=self.abrirRutaBusqueda,
                                            width=16,
                                            height=1,
                                            anchor="center", 
                                            justify="center",
                                            relief="groove")
                                            
        self.etiquetaDestino = tk.Label(self.ventaDatosRuta, 
                                        text="Seleccionar la carpeta de destino de archivos")
        self.campoRutaDestino = tk.Entry(self.ventaDatosRuta, width=50)
        self.btnExaminarDestino= tk.Button(self.ventaDatosRuta, 
                                            text="Examinar", 
                                            command=self.abrirRutaDestino,
                                            width=16,
                                            height=1,
                                            anchor="center", 
                                            justify="center", 
                                            relief="groove")

        self.btnContinuar = tk.Button(self.ventaDatosRuta, 
                                    text="Continuar", 
                                    state="disabled", 
                                    command=self.continuar,
                                    width=16, 
                                    height=1,             
                                    anchor="center", 
                                    justify="center", 
                                    relief="groove")


        # Posicionamiento
        self.etiquetaArchivo.grid(column=0, row=0, padx=5, pady=5)
        self.campoArchivo.grid(column=0, row=1, padx=10, pady=7)
        self.btnExaminar.grid(column=1, row=1, padx=5, pady=7)

        self.etiquetaBusqueda.grid(column=0, row=2, padx=5, pady=5)
        self.campoRutaBusqueda.grid(column=0, row=3, padx=10, pady=7)
        self.btnExaminarUbicacion.grid(column=1, row=3, padx=5, pady=7)

        self.etiquetaDestino.grid(column=0, row=4, padx=5, pady=5)
        self.campoRutaDestino.grid(column=0, row=5, padx=10, pady=7)
        self.btnExaminarDestino.grid(column=1, row=5, padx=5, pady=7)

        self.btnContinuar.grid(column=0, row=6, padx=5, pady=7)
        self.ventaDatosRuta.after(0, self.ventaDatosRuta.attributes, '-alpha', 1.0)
        self.ventaDatosRuta.focus_force()
        self.ventaDatosRuta.transient(master=self.parent)
        self.ventaDatosRuta.grab_set()
        self.parent.wait_window(self.ventaDatosRuta)

    """Funcion para obtener ruta de archivo de cubo"""

    def abrirArchivo(self):
        filename = askopenfilename(initialdir="", title="Seleccionar el archivo de cubo sísmico",
                                   parent=self.ventaDatosRuta,
                                   filetypes=(("xlm files", "*.xlm"), ("xls files", "*.xls*")))
        if filename == '':
            messagebox.showwarning(title="Advertencia", message="Favor de seleccionar un achivo")
            return 
        
        self.rutaArchivo = filename
        self.campoArchivo.delete(0, "end")
        self.campoArchivo.insert(0, self.rutaArchivo)
        # Validación para habilitar o deshabilitar botón de continuar
        if (self.campoArchivo.get() == '' 
            or self.campoRutaBusqueda.get() == ''
            or self.campoRutaDestino.get() == ''):
            self.btnContinuar.config(state="disabled")
        else:
            self.btnContinuar.config(state="normal")

    """Funcion para obtener ruta donde se alojaran los resultadas de preiddcion del cubo"""

    def abrirRutaBusqueda(self):

        filename = askdirectory(initialdir="",
                                title="Seleccionar la ruta para guardar los resultados de predicción de cubo",
                                parent=self.ventaDatosRuta)
        if filename == '':
            messagebox.showwarning(title="Advertencia", message="Favor de seleccionar una ruta")
        else:
            self.rutaBusqueda = filename
            self.campoRutaBusqueda.delete(0, "end")
            self.campoRutaBusqueda.insert(0, self.rutaBusqueda)
            # Validación para habilitar o deshabilitar botón de continuar
            if (self.campoArchivo.get() == '' 
                or self.campoRutaBusqueda.get() == ''
                or self.campoRutaDestino.get() == ''):
                self.btnContinuar.config(state="disabled")
            else:
                self.btnContinuar.config(state="normal")

    def abrirRutaDestino(self):

        filename = askdirectory(initialdir="",
                                title="Seleccionar la ruta para guardar los resultados de predicción de cubo",
                                parent=self.ventaDatosRuta)
        if filename == '':
            messagebox.showwarning(title="Advertencia", message="Favor de seleccionar una ruta")
        else:
            self.rutaDestino = filename
            self.campoRutaDestino.delete(0, "end")
            self.campoRutaDestino.insert(0, self.rutaDestino)
            # Validación para habilitar o deshabilitar botón de continuar
            if (self.campoArchivo.get() == '' 
                or self.campoRutaBusqueda.get() == ''
                or self.campoRutaDestino.get() == ''):
                self.btnContinuar.config(state="disabled")
            else:
                self.btnContinuar.config(state="normal")

    """Funcion para llamar a la clase que realizara la prediccion de los cubos
    Asi como una validacion para saber si el modelo actual cumple con las variables
    indicadas para la prediccion del cubo"""

    def continuar(self):
        self.ventaDatosRuta.destroy()
        # Validacion de variables predictoras

    def analizarMover(self):
        pass
    
"""Función principal"""
def main():
    root = tk.Tk()
    APP = VentanaPrincipal(root)
    try:
        root.iconbitmap('Logo_Bienvenida_IMP_PREDICT_V2.ico')
        root.mainloop()
    except:
        root.mainloop()


if __name__ == '__main__':
    main()
