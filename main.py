import csv
import os
import queue
import threading
from shutil import copyfile

from os import listdir
from os.path import isfile, join

import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename, asksaveasfilename, askdirectory
from tkinter.colorchooser import askcolor
from tkinter import messagebox
import tkinter.font as tkFont

import pandas as pd


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
        self.parent.bind("<Control-n>", self.abrirProcesoEvent)

    def cerrandoVentanaEvent(self, event):
        self.cerrandoVentana()

    def abrirProcesoEvent(self, event):
        self.abrirNuevo()

   
    """Función para evitar cerrar abruptamente el programa"""

    def cerrandoVentana(self):

        respuestaDeCierre = messagebox.askyesno("Salir de Lee y Mueve", 
                                                "¿Desea realmente salir de Lee y mueve")
        if respuestaDeCierre == True:
            self.parent.destroy()

    def abrirNuevo(self):
        VentanaLecturaArchivo(self.parent,"")


class VentanaLecturaArchivo:
    def __init__(self, parent,rutaArchivo):
        self.parent = parent
        self.rutaArchivo  = rutaArchivo
        self.init_ui()


    def init_ui(self):
        a = self.parent.winfo_screenwidth()
        h = self.parent.winfo_screenheight()
        self.columnasDisponibles = []
        self.variablesBotones = []
        
        # Definicion de tamaño de venatana
        ancho = 670
        alto = 160
        posx = int((a - ancho) / 2)
        posy = int((h - alto) / 2)
        self.ventaDatosRuta = tk.Toplevel(self.parent)
        self.ventaDatosRuta.geometry(str(ancho) + "x" + str(alto) + "+" + str(posx) + "+" + str(posy))
        self.ventaDatosRuta.attributes('-alpha', 0.0)
        self.ventaDatosRuta.resizable(0, 0)
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
        self.btnContinuar = tk.Button(self.ventaDatosRuta, 
                                    text="Continuar", 
                                    state="disabled", 
                                    command=self.lecturaArchivos,
                                    width=16, 
                                    height=1,             
                                    anchor="center", 
                                    justify="center", 
                                    relief="groove")

        # Posicionamiento de widgets
        self.etiquetaArchivo.grid(column=0, row=0, padx=5, pady=5)
        self.campoArchivo.grid(column=0, row=1, padx=10, pady=7)
        self.btnExaminar.grid(column=1, row=1, padx=5, pady=7)
        self.btnContinuar.grid(column=0, row=2, padx=5, pady=7)    
        # Validar si se recibe datos de la ruta con la que se trabaja
        if self.rutaArchivo:
            self.campoArchivo.delete(0, "end")
            self.campoArchivo.insert(0, self.rutaArchivo)
            self.btnContinuar.config(state="normal")


    def abrirArchivo(self):
        """
        Funcion para obtener ruta de archivo Excel
        """
        filename = askopenfilename(initialdir="", title="Seleccionar el archivo a Excel",
                                   parent=self.ventaDatosRuta,
                                   filetypes=(("xls* files", "*.xls*"),("xl* files", "*.xl*")))
        if filename == '':
            messagebox.showwarning(title="Advertencia", message="Favor de seleccionar un achivo")
            return 

        self.rutaArchivo = filename
        self.campoArchivo.delete(0, "end")
        self.campoArchivo.insert(0, self.rutaArchivo)
        # Validación para habilitar o deshabilitar botón de continuar
        if self.campoArchivo.get() == '':
            self.btnContinuar.config(state="disabled")
        else:
            self.btnContinuar.config(state="normal")

    def lecturaArchivos(self):
        try:
            
            xls = pd.read_excel(self.rutaArchivo ) 
            print(xls)
        except FileNotFoundError as e:
            messagebox.showwarning(title="Error",
                            message="No se Encuentra el archivo")
            return
        except Exception as e:
            messagebox.showwarning(title="Error",
                            message=e)
            return
        print("columnas a e leer")
        print(xls.columns)
        self.columnasDisponibles = list(xls.columns)
        self.ventaDatosRuta.destroy()
        VentanaRutas(self.parent,self.columnasDisponibles,self.rutaArchivo)

    
class VentanaRutas:
    def __init__(self, parent, columnasDisponibles,rutaArchivo ):
        self.parent = parent
        self.columnasDisponibles = columnasDisponibles 
        self.rutaArchivo = rutaArchivo
        self.init_ui()
        

    def init_ui(self):
        a = self.parent.winfo_screenwidth()
        h = self.parent.winfo_screenheight()
        self.variablesBotones = []
        self.columnaAnaliza = self.columnasDisponibles[0]
        # Definicion de tamaño de venatana
        ancho = 670
        alto = 470
        posx = int((a - ancho) / 2)
        posy = int((h - alto) / 2)
        self.ventaDatosRuta = tk.Toplevel(self.parent)
        self.ventaDatosRuta.geometry(str(ancho) + "x" + str(alto) + "+" + str(posx) + "+" + str(posy))
        self.ventaDatosRuta.attributes('-alpha', 0.0)
        self.ventaDatosRuta.resizable(0, 0)

      
        # Etiquetas
        etiquetaIntervalosE = tk.Label(self.ventaDatosRuta, text="Seleccionar la columna a analizar ")
        etiquetaIntervalosE.grid(column=0, row=2, padx=10, pady=2, sticky="s")
        

        # Posicinamiento de radiobuttom para seleccion de
        # column a buscar
        self.frameScrollIntervaloBlind = tk.Frame(self.ventaDatosRuta, relief="groove", width=150, height=150, bd=1,
                                                  bg="#FFFFFF")
        self.frameScrollIntervaloBlind.grid(column=0, row=3, padx=20, pady=5)
        self.canvasScrollIntervaloBlind = tk.Canvas(self.frameScrollIntervaloBlind, bg="#FFFFFF")
        self.frameVisualScrollIntervaloBlind = tk.Frame(self.canvasScrollIntervaloBlind, bg="#FFFFFF")
        self.scrollBarIntervaloBlind = tk.Scrollbar(self.frameScrollIntervaloBlind, orient="vertical",
                                                    command=self.canvasScrollIntervaloBlind.yview)
        self.canvasScrollIntervaloBlind.configure(yscrollcommand=self.scrollBarIntervaloBlind.set)
        self.scrollBarIntervaloBlind.pack(side="right", fill="y")
        self.canvasScrollIntervaloBlind.pack(side="left")
        self.canvasScrollIntervaloBlind.create_window((0, 0), window=self.frameVisualScrollIntervaloBlind, anchor='nw')
        self.frameVisualScrollIntervaloBlind.bind("<Configure>", self.habilitaScrollIntervaloBlind)

        for l in range(0, len(self.columnasDisponibles)):
            self.variablesBotones.append(l)
        # Variable para guardar dato de radioButton
        self.nombreVariablesIntervaloBlind = tk.IntVar()
        
        for i in range(0, len(self.columnasDisponibles)):
            self.variablesBotones[i] = tk.Radiobutton(self.frameVisualScrollIntervaloBlind,
                                                                    text=self.columnasDisponibles[i],
                                                                    variable=self.nombreVariablesIntervaloBlind,
                                                                    value=i, bg="#FFFFFF",
                                                                    command=lambda i=i: self.printColumna(i))
            self.variablesBotones[i].grid(row=i + 1, padx=20, pady=1)




        # Creacion de widgets de ventana
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

        self.btnAtras= tk.Button(self.ventaDatosRuta, 
                                    text="Anterior", 
                                    state="active", 
                                    command=self.irAtras,
                                    width=16, 
                                    height=1,             
                                    anchor="center", 
                                    justify="center", 
                                    relief="groove")

        # Posicionamiento de widgets
        self.etiquetaBusqueda.grid(column=0, row=5, padx=5, pady=5)
        self.campoRutaBusqueda.grid(column=0, row=6, padx=10, pady=7)
        self.btnExaminarUbicacion.grid(column=1, row=6, padx=5, pady=7)

        self.etiquetaDestino.grid(column=0, row=7, padx=5, pady=5)
        self.campoRutaDestino.grid(column=0, row=8, padx=10, pady=7)
        self.btnExaminarDestino.grid(column=1, row=8, padx=5, pady=7)

        self.btnContinuar.grid(column=1, row=9, padx=5, pady=7)
        self.btnAtras.grid(column=0, row=9, padx=1, pady=7)
        self.ventaDatosRuta.after(0, self.ventaDatosRuta.attributes, '-alpha', 1.0)
        self.ventaDatosRuta.focus_force()
        self.ventaDatosRuta.transient(master=self.parent)
        self.ventaDatosRuta.grab_set()
        self.parent.wait_window(self.ventaDatosRuta)


    def printColumna(self,index):
        self.columnaAnaliza = self.columnasDisponibles[index]
        print(f"La columna selecionada es : {self.columnaAnaliza}")
    

    def habilitaScrollIntervaloBlind(self, event):
        """
        Función para habilitar el scroll cuando los elementos no caben en la
        región del frame
        Args:
        event: Para habilitar la que sea activa por evento esta funcion
        """
        self.canvasScrollIntervaloBlind.configure(
            scrollregion=self.canvasScrollIntervaloBlind.bbox("all"), 
            width=150,
            height=150)

    def abrirRutaBusqueda(self):
        """
        Obtener Ruta de busqueda de archivos
        """

        filename = askdirectory(initialdir="",
                                title="Seleccionar la ruta para buscar los archivos",
                                parent=self.ventaDatosRuta)
        if filename == '':
            messagebox.showwarning(title="Advertencia", message="Favor de seleccionar una ruta")
        else:
            self.rutaBusqueda = filename
            self.campoRutaBusqueda.delete(0, "end")
            self.campoRutaBusqueda.insert(0, self.rutaBusqueda)
            # Validación para habilitar o deshabilitar botón de continuar
            if (self.campoRutaBusqueda.get() == ''
                or self.campoRutaDestino.get() == ''):
                self.btnContinuar.config(state="disabled")
            else:
                self.btnContinuar.config(state="normal")

    def abrirRutaDestino(self):

        filename = askdirectory(initialdir="",
                                title="Seleccionar la ruta para copiar los archivos",
                                parent=self.ventaDatosRuta)
        if filename == '':
            messagebox.showwarning(title="Advertencia", message="Favor de seleccionar una ruta")
        else:
            self.rutaDestino = filename
            self.campoRutaDestino.delete(0, "end")
            self.campoRutaDestino.insert(0, self.rutaDestino)
            # Validación para habilitar o deshabilitar botón de continuar
            if (self.campoRutaBusqueda.get() == ''
                or self.campoRutaDestino.get() == ''):
                self.btnContinuar.config(state="disabled")
            else:
                self.btnContinuar.config(state="normal")


    def continuar(self):
        parametros = {
                    "rutaArchivo":self.rutaArchivo,
                    "rutaDestino":self.rutaDestino,
                    "rutaBusqeuda":self.rutaBusqueda,
                    "columnaAnaliza":self.columnaAnaliza}
        self.ventaDatosRuta.destroy()
        self.queueMia = queue.Queue()
        ThreadedTask(self.queueMia, self.parent, parametros).start()
        self.parent.after(100, self.process_queue)
    
    def process_queue(self):
        try:
            msg = self.queueMia.get(0)
        except queue.Empty:
            self.parent.after(100, self.process_queue)

    def irAtras(self):
        self.ventaDatosRuta.destroy()
        VentanaLecturaArchivo(self.parent, self.rutaArchivo)
        


"""Clase que crea un hilo para manejar ambos eventos en multitarea"""


class ThreadedTask(threading.Thread):
    def __init__(self, queueMia, parent, parametros):
        threading.Thread.__init__(self)
        self.queueMia = queueMia
        self.parent = parent
        self.rutaArchivo =  parametros['rutaArchivo']        
        self.rutaDestino =  parametros['rutaDestino']
        self.rutaBusqueda =  parametros['rutaBusqeuda']
        self.columnaAnaliza =  parametros['columnaAnaliza']

    def run(self):
        self.tareaPrincipal()
        
    def ventanaInicio(self):
        a = self.parent.winfo_screenwidth()
        h = self.parent.winfo_screenheight()
        ancho = 445
        alto = 85
        posx = int((a - ancho) / 2)
        posy = int((h - alto) / 2)
        self.ventanaIni = tk.Toplevel(self.parent)
        self.ventanaIni.geometry(str(ancho) + "x" + str(alto) + "+" + str(posx) + "+" + str(posy))
        self.ventanaIni.attributes('-alpha', 0.0)
        self.ventanaIni.resizable(0, 0)
        try:
            self.ventanaIni.iconbitmap('Logo_Bienvenida_IMP_PREDICT_V2.ico')
            self.ventanaIni.title("Procesando....")
        except:
            self.ventanaIni.title("Procesando....")
        self.ventanaIni.protocol("WM_DELETE_WINDOW", self.cerrandoVentanaInicio)

        # Configuración widgets
        self.miFuente = tkFont.Font(size=10)
        self.etiquetaC = tk.Label(self.ventanaIni,
                                  text="El proceso ha iniciado porfavor esper ")
        self.etiquetaGamma = tk.Label(self.ventanaIni, text="Nota: Esta etapa podría demorar varios minutos...")

        # Posicionamiento
        self.etiquetaC.pack(pady=10)
        self.etiquetaGamma.pack(pady=1)
        self.ventanaIni.after(0, self.ventanaIni.attributes, '-alpha', 1.0)
        self.ventanaIni.focus_force()
        self.ventanaIni.grab_set()
        self.ventanaIni.transient(master=self.parent)


    def cerrandoVentanaInicio(self):
        messagebox.showwarning(title="Advertencia", message="Favor de esperar a que termine el procesamiento...")

    def tareaPrincipal(self):
        self.ventanaInicio()
        self.rutaArchivo 
        self.rutaDestino 
        self.rutaBusqueda 
        self.columnaAnaliza
        # Lectura de archivo excel
        try:
            xls = pd.read_excel(self.rutaArchivo ) 
            print(xls)
        except FileNotFoundError as e:
            messagebox.showwarning(title="Error",
                            message="No se Encuentra el archivo")
            return
        except Exception as e:
            messagebox.showwarning(title="Error",
                            message=e)
            return
        # Busqeuda de archivos en la carpeta de busqueda
        mypath = self.rutaBusqueda
        onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]
        print(f"Ruta busqeudda: {onlyfiles}")
        columnaTotal = xls[self.columnaAnaliza]
        for i in columnaTotal:
            if i in onlyfiles:
                origen = os.path.join(mypath,i)
                destino =os.path.join(self.rutaDestino,i)
                print(f"ruta total: {origen}")
                try:
                    copyfile(origen,destino)
                except Exception as e:
                    print(f"No se logro copiar el archivo")
                    print(e)
                    messagebox.showinfo('Upss!', 
                                        f'No se copio el archivo: {i}')
        
            
        self.ventanaIni.destroy()
        messagebox.showinfo('Proceso finalizado', 'Se copiaron los archivos')



    
"""Función principal"""
def main():
    root = tk.Tk()
    APP = VentanaPrincipal(root)

    root.mainloop()


if __name__ == '__main__':
    main()
