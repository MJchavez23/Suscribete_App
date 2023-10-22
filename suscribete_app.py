import csv
import pywhatkit
import tkinter as tk
from tkinter import messagebox
from tkinter.filedialog import askopenfile, asksaveasfilename
import os.path
from os.path import abspath
import win32clipboard
from PIL import Image, ImageTk

class Menu(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("App de Mensajería")
        self.resizable(False, False)
        self.lista_numeros = []
        self.label_verificar = tk.Label(self ,text="")
        self.direccion_imagen = ""
        self.label_verificar.grid(row=8, column = 1)
        self.componentes()
        self.mainloop()

    def componentes(self):
        label1 = tk.Label(self, text="1. Suba el archivo csv.", font=("Helvetica",12)).grid(row=0, column= 0)
        button_archivo_csv = tk.Button(self, text="Subir", command=self.guardar_numeros)
        button_archivo_csv.grid(row= 1, column=0)
        label2 = tk.Label(self, text="2. Suba la imagen que se va enviar.", font=("Helvetica",12)).grid(row = 3)
        button_subir_archivo = tk.Button(self, text="Subir imagen.", command=self.guardar_imagen).grid(row = 4,)
        label3 = tk.Label(self, text="3. Ingrese el texto que se va a enviar.", font=("Helvetica", 12)).grid(row=5)
        self.texto_mensaje = tk.Text(self, width=50, height=15, font=("Helvetica",12))
        self.texto_mensaje.grid(row=6 , padx = 5, pady=5)
        label4 = tk.Label(self, text="4. Si todo esta correcto prosiga a enviar los mensajes.", font=("Helvetica",12)).grid(row=7)
        self.label_archivo_correctamente = tk.Label(self, text="")
        self.label_archivo_correctamente.grid(row=1, column = 1)
        self.label_imagen_subida = tk.Label(self, text="")
        self.label_imagen_subida.grid(row=4, column = 1)
        button_enviar_mensajes = tk.Button(self, text="Enviar", command=self.enviar_mensaje)
        button_enviar_mensajes.grid(row=8, pady = 5)

    def guardar_imagen(self):
        self.direccion_imagen = ""
        try:
            self.archivo_abierto = askopenfile(mode="r")
            nombre, extension = os.path.splitext(self.archivo_abierto.name)
            if extension == ".jpg" or extension == ".png":
                self.direccion_imagen = os.path.abspath(self.archivo_abierto.name)
                self.img1 = tk.Label(self)
                self.img1.grid(row = 6, column=1)
                image_original = Image.open(self.direccion_imagen).resize((200,180))
                self.image_tk = ImageTk.PhotoImage(image_original)
                self.img1.config(image=self.image_tk)
                self.label_imagen_subida.config(text="Imagen subida correctamente", fg="green")

            else:
                self.label_imagen_subida.config(text="ERROR: Vuelva a subir la imagen (Solo archivos jpg o png)", fg="red")
                self.img1.destroy()
        except :
            self.label_imagen_subida.config(text="ERROR: Vuelva a subir la imagen (Solo archivos jpg o png)", fg="red")
            self.img1.destroy()
    
    def guardar_numeros(self):
        self.lista_numeros = []
        try:
            #Abrimos el explorador para subir el archivo
            self.archivo_abierto = askopenfile(mode="r")
            nombre, extension = os.path.splitext(self.archivo_abierto.name)
            if extension == ".csv":
                self.label_archivo_correctamente.config(text="Archivo subido CORRECTAMENTE", fg="green")
                with open(os.path.abspath(self.archivo_abierto.name), "r", encoding="utf-8") as csv_file:
                    csv_reader = csv.reader(csv_file, delimiter=";")
                    next(csv_reader)
                    for line in csv_reader:
                        if len(line[3]) == 10 and line[3][0] == "0":
                            self.lista_numeros.append(line[3])
            else:
                self.label_archivo_correctamente.config(text="ERROR: Suba un archivo de tipo CSV", fg="red")
        except Exception as e:
            pass


    def enviar_mensaje(self):
        answer = messagebox.askokcancel(title="Confirmación",message="¿Desea continuar? Después de dar aceptar, no podrá parar el proceso.")
        if answer:
            mensaje = self.texto_mensaje.get(1.0, tk.END).strip()
            try:
                if len(self.lista_numeros) > 0 and mensaje:
                        for numero in self.lista_numeros:
                            if len(numero) == 10:
                                if self.direccion_imagen != "":
                                    pywhatkit.sendwhats_image("+593"+numero,caption=mensaje, img_path=self.direccion_imagen, wait_time = 20 ,tab_close=True)
                                else:
                                    pywhatkit.sendwhatmsg_instantly("+593"+numero, mensaje, tab_close=True)
                        self.label_verificar.config(text="Mensajes enviados correctamente", fg="green")
                else:
                    self.label_verificar.config(text="ERROR: A ocurrido un error al enviar los mensajes, intente de nuevo", fg="red")
            except Exception:
                pass
        else: 
            pass

Menu()