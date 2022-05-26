from cgitb import text
from doctest import master
from fileinput import filename
from multiprocessing.sharedctypes import Value
import pandas as pd
import openpyxl as op
from tkinter import filedialog, messagebox
import tkinter as tk
import customtkinter
from PIL import Image, ImageTk
from Controlyseguimiento.control_seguimiento import con_y_seg
from Cambios.cambios import cambios, path_leaf
import tkinter
import tkinter.messagebox
import customtkinter
import sys


customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
#customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):

    WIDTH = 600
    HEIGHT = 350

    def __init__(self):
        super().__init__()

        self.title("Contro y seguimiento de censos")
        self.geometry(f"{App.WIDTH}x{App.HEIGHT}")
        # self.minsize(App.WIDTH, App.HEIGHT)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        if sys.platform == "darwin":
            self.bind("<Command-q>", self.on_closing)
            self.bind("<Command-w>", self.on_closing)
            self.createcommand('tk::mac::Quit', self.on_closing)

        # ============ create frame ============

        # configure grid layout (1x2)
        self.grid_columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        self.frame_right = customtkinter.CTkFrame(master=self)
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)

        # ============ frame_right ============

        # configure grid layout (3x7)
        #for i in [0, 1, 2, 3]:
            #self.frame_right.rowconfigure(i, weight=1)
        self.frame_right.rowconfigure(7, weight=10)
        self.frame_right.columnconfigure(0, weight=1)
        self.frame_right.columnconfigure(1, weight=1)
        self.frame_right.columnconfigure(2, weight=0)
        self.frame_info = customtkinter.CTkFrame(master=self.frame_right)
        self.frame_info.grid(row=1, column=0, columnspan=4, rowspan=4, pady=20, padx=20, sticky="nsew")


        self.switch_2 = customtkinter.CTkSwitch(master=self.frame_info,
                                                text="Modo oscuro",
                                                command=self.change_mode)
        self.switch_2.grid(row=8, column=0, pady=10, padx=20, sticky="w")
        # ============ frame_right -> frame_info ============

        self.frame_info.rowconfigure(1, weight=1)
        self.frame_info.columnconfigure(0, weight=2)

        self.label_info_1 = customtkinter.CTkLabel(master=self.frame_info,
                                                    text="                Contro y seguimiento de censos.    \n" +
                                                        "1. Seleccione los archivos a comparar.\n" +
                                                        "2. Al terminar el proceso aparecerá un mensaje exitoso." ,
                                                    height=100,
                                                    fg_color=("white", "gray38"),  # <- custom tuple-color
                                                    justify=tkinter.LEFT)
        self.label_info_1.grid(column=0, row=0, sticky="nwe", padx=15, pady=15)

        self.progressbar = customtkinter.CTkProgressBar(master=self.frame_info)
        self.progressbar.grid(row=1, column=0, sticky="ew", padx=15, pady=15)

        # ============ frame_right <- ============

        self.boton_importar = customtkinter.CTkButton(master=self.frame_right,
                                                        text= "Formato de observaciones",
                                                        command=self.new_window)
        self.boton_importar.grid(row=6, column=1, pady=20, padx=20, sticky="w")
        
        self.boton_importar_2 = customtkinter.CTkButton(master=self.frame_right,
                                                        text="Revisión de cambios",
                                                        command=self.revisar_cambios)
        self.boton_importar_2.grid(row=6, column=0, pady=20, padx=20, sticky="w")

        self.button_5 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Salir",
                                                command=self.destroy)
        self.button_5.grid(row=6, column=2, columnspan=1, pady=20, padx=20, sticky="we")

    def window_2(self):
        """


        Parameters
        ----------
        función Self que se invoca para regresar a la ventana de inicio

        """        
        self.progressbar.set(0.5)
        self.title("Contro y seguimiento de censos")
        self.geometry(f"{App.WIDTH}x{App.HEIGHT}")
        # self.minsize(App.WIDTH, App.HEIGHT)

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        if sys.platform == "darwin":
            self.bind("<Command-q>", self.on_closing)
            self.bind("<Command-w>", self.on_closing)
            self.createcommand('tk::mac::Quit', self.on_closing)

        # ============ create frame ============

        # configure grid layout (1x2)
        self.grid_columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)

        self.frame_right = customtkinter.CTkFrame(master=self)
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)

        # ============ frame_right ============

        # configure grid layout (3x7)
        #for i in [0, 1, 2, 3]:
            #self.frame_right.rowconfigure(i, weight=1)
        self.frame_right.rowconfigure(7, weight=10)
        self.frame_right.columnconfigure(0, weight=1)
        self.frame_right.columnconfigure(1, weight=1)
        self.frame_right.columnconfigure(2, weight=0)
        self.frame_info = customtkinter.CTkFrame(master=self.frame_right)
        self.frame_info.grid(row=1, column=0, columnspan=4, rowspan=4, pady=20, padx=20, sticky="nsew")


        self.switch_2 = customtkinter.CTkSwitch(master=self.frame_info,
                                                text="Modo oscuro",
                                                command=self.change_mode)
        self.switch_2.grid(row=8, column=0, pady=10, padx=20, sticky="w")
        # ============ frame_right -> frame_info ============

        self.frame_info.rowconfigure(1, weight=1)
        self.frame_info.columnconfigure(0, weight=2)

        self.label_info_1 = customtkinter.CTkLabel(master=self.frame_info,
                                                    text="                Contro y seguimiento de censos.    \n" +
                                                        "1. Seleccione los archivos a comparar.\n" +
                                                        "2. Al terminar el proceso aparecerá un mensaje exitoso." ,
                                                    height=100,
                                                    fg_color=("white", "gray38"),  # <- custom tuple-color
                                                    justify=tkinter.LEFT)
        self.label_info_1.grid(column=0, row=0, sticky="nwe", padx=15, pady=15)

        self.progressbar = customtkinter.CTkProgressBar(master=self.frame_info)
        self.progressbar.grid(row=1, column=0, sticky="ew", padx=15, pady=15)

        # ============ frame_right <- ============

        self.boton_importar = customtkinter.CTkButton(master=self.frame_right,
                                                        text= "Formato de observaciones",
                                                        command=self.new_window)
        self.boton_importar.grid(row=6, column=1, pady=20, padx=20, sticky="w")
        
        self.boton_importar_2 = customtkinter.CTkButton(master=self.frame_right,
                                                        text="Revisión de cambios",
                                                        command=self.revisar_cambios)
        self.boton_importar_2.grid(row=6, column=0, pady=20, padx=20, sticky="w")

        self.button_5 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Salir",
                                                command=self.destroy)
        self.button_5.grid(row=6, column=2, columnspan=1, pady=20, padx=20, sticky="we")


    def change_mode(self):
        """


        Parameters
        ----------
        función self para cambiar apariencia de color de la app

        Returns
        -------
        cambia de modo dark a modo light de acuerdo al valor del switch
        1 = dark
        0 = light
        por default la app comienza en modo dark

        """        
        if self.switch_2.get() == 1:
            customtkinter.set_appearance_mode("dark")
        else:
            customtkinter.set_appearance_mode("light")

    def on_closing(self, event=0):
        """


        Parameters
        ----------
        función self para destruir la ventana y el proceso.

        """
        self.destroy()

    def start(self):
        """


        Parameters
        ----------
        función self para comenzar en forma de loop la app

        """
        self.mainloop()

    def revisar_cambios(self):
        """


        Parameters
        ----------
        la función recibe dos parametros. Dos libros de excel para compararlos y llamar a la función cambios 

        Returns
        -------
        regresa un mensaje avisando que se ha terminado el proceso

        """

        libro = filedialog.askopenfilename()
        libro2 = filedialog.askopenfilename()        
        cambios(libro, libro2)
        messagebox.showinfo('Aviso', 'Se ha terminado el proceso de revisión')

    
    def control(self, entry1, entry2, entry3, entry4, entry5):
        """


        Parameters
        ----------
        Se requieren dos libros de excel para comenzar proceso de llenado de FO llamando
        a la función con_y_seg del modulo control_seguimiento, ademas requiere de variables que se obtienen
        de los entrys.

        Returns
        -------
        regresa un excel con los datos que se piden para el llenado
        al finalizar muestra un mensje de que se ha terminado el proceso.
        

        """
        book = filedialog.askopenfilename()
        book2  = op.load_workbook(book)
        observaciones = filedialog.askopenfilename()
        foco = con_y_seg(book2, observaciones, entry1, entry2, entry3, entry4, entry5)
        nombre_archivo_salvado = path_leaf(book)
        nombre_archivo_salvado = 'observaciones_' + nombre_archivo_salvado
        data = [('xlsx', '*.xlsx')] 
        filename = filedialog.asksaveasfile(filetypes=data, defaultextension=data,initialfile = nombre_archivo_salvado)
        foco.save(filename.name)
        messagebox.showinfo('Aviso', 'Se ha completado el proceso')

    def get_value(self):
        """


        Parameters
        ----------
        Función self para obtener los valores de los entry para pasarselos a la función control

        Returns
        -------
        devuelve los valores de los entrys

        """
        text1=self.entry1.get()
        text2=self.entry2.get()
        text3=self.entry3.get()
        text4=self.entry4.get()
        text5=self.entry5.get()
        self.control(text1, text2, text3, text4, text5)

    def new_window(self):
        self.title("Formato de Observaciones (FO)")
        self.geometry("750x400")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.grid_columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self.frame_right = customtkinter.CTkFrame(master=self)
        self.frame_right.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)
                # configure grid layout (3x7)
        #for i in [0, 1, 2, 3]:
            #self.frame_right.rowconfigure(i, weight=1)
        self.frame_right.rowconfigure(7, weight=10)
        self.frame_right.columnconfigure(0, weight=1)
        self.frame_right.columnconfigure(1, weight=1)
        self.frame_right.columnconfigure(2, weight=0)
        self.frame_info = customtkinter.CTkFrame(master=self.frame_right)
        #self.frame_info.grid(row=1, column=0, columnspan=4, rowspan=4, pady=20, padx=20, sticky="nsew")

        # ============ frame_right -> frame_info ============

        self.frame_info.rowconfigure(5, weight=1)
        self.frame_info.columnconfigure(0, weight=2)


        # ============ frame_right <- ===========
        self.boton_importar = customtkinter.CTkButton(master=self.frame_right,
                                                        text= "Generar Formato",
                                                        command=self.get_value)
        self.boton_importar.grid(row=3, column=2, pady=20, padx=20, sticky="w")                            
        self.button_5 = customtkinter.CTkButton(master=self.frame_right,
                                                text="Salir",
                                                command=self.destroy)
        self.button_5.grid(row=4, column=2, columnspan=1, pady=20, padx=20, sticky="we")
        self.gobackButton = customtkinter.CTkButton(master=self.frame_right,
                                                    text="Regresar",
                                                    command=self.window_2) #here should be the command for the button
        self.gobackButton.grid(row=2, column=2, columnspan=1, pady=20, padx=20, sticky="we")   
        self.entry1 = customtkinter.CTkEntry(master=self.frame_right,
                                            width=120,
                                            placeholder_text='Qué número de revisión es? ')
        self.entry1.grid(row=0, column=0, columnspan=2,  rowspan=1, pady=20, padx=20, sticky="we")

        self.entry2 = customtkinter.CTkEntry(master=self.frame_right,
                                            width=120,
                                            placeholder_text='Cuál es la fecha programada de entrega? ')
        self.entry2.grid(row=1, column=0, columnspan=2, rowspan=1, pady=20, padx=20, sticky="we")

        self.entry3 = customtkinter.CTkEntry(master=self.frame_right,
                                            width=120,
                                            placeholder_text='Cuál fue la fecha real de entrega? ')
        self.entry3.grid(row=2, column=0, columnspan=2, rowspan=1, pady=20, padx=20, sticky="we")

        self.entry4 = customtkinter.CTkEntry(master=self.frame_right,
                                            width=120,
                                            placeholder_text='Cuál es el año del levantamiento?')
        self.entry4.grid(row=3, column=0, columnspan=2, rowspan=1, pady=20, padx=20, sticky="we")

        self.entry5 = customtkinter.CTkEntry(master=self.frame_right,
                                            width=120,
                                            placeholder_text='Cuántas preguntas tiene el módulo? ')
        self.entry5.grid(row=4, column=0, columnspan=2, rowspan=1, pady=20, padx=20, sticky="we")



if __name__ == "__main__":
    app = App()
    app.start()