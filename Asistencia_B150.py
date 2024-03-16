#!/usr/bin/env python
# -*- coding: utf-8 -*-
from tkinter import * #Interface
import pandas as pd  #Para manipulación de la base de datos en excel
import time #Para obtener la hora y fecha
from tkinter import messagebox, filedialog
from pyzbar.pyzbar import decode #Para la decodificación de codigos QR
from pyzbar.pyzbar import ZBarSymbol
import cv2 # Open CV Manipulación de la web cam
from datetime import date
import numpy #Para manejar los arreglos y arrays
from PIL import Image, ImageTk, ImageDraw #Posicionar imagenes

class App:
    def __init__(self, font_video=0):
        #Data
        self.lista_codigos = open("Data/qr_list.txt", "r")
        self.lista_usuarios = open("Data/id_list.txt", "r")
        self.a = self.lista_codigos.read().splitlines()
        self.b = self.lista_usuarios.read().splitlines()
        self.dic = zip(self.a, self.b)
        self.new_data = dict(self.dic)
        self.active_camera = False
        self.info = []
        self.codelist= []
        self.new_codelist = "".join(self.codelist)

        #Ventana
        self.appName = 'Sistema de asistencias'
        self.ventana = Tk()
        self.ventana.iconbitmap('Files/logo.ico')
        self.font_video = font_video
        self.ventana.title('Sistema de asistencias')
        self.ventana.config(bg="#F38F00")


        #Titulo

        self.Label = Label(self.ventana, text='SISTEMA DE ASISTENCIAS', font=30, bg="#F38F00" )
        self.Label.config(fg="black", font=("Arial", 30))
        self.Label.place(x=420, y=20)

        #Boton entrada
        btn_entrada = PhotoImage(file="Files/boton_entrada.png")
        boton = Button(self.ventana,image= btn_entrada,command=self.firmar_entrada)
        boton.place(x=545,y=600)
        #Boton salida
        btn_salida = PhotoImage(file="Files/boton_salida.png")
        boton_salida = Button(self.ventana,image=btn_salida, command=self.firmar_salida)
        boton_salida.place(x=690, y=600)

        #Canvas
        self.canvas = Canvas(self.ventana, bg='black', width=740, height=530)
        self.canvas.place(x=360, y=90)

        #Logo
        logo = ImageTk.PhotoImage(Image.open("Files/logo.png"))
        label_logo = Label(self.ventana, image=logo, bg="#F38F00")
        label_logo.place(x=1083, y=230)

        #Bienvenido

        self.bienvenido = Label(self.ventana,text="Bienvenido", bg="#F38F00")
        self.bienvenido.config(fg="black", font=("Arial", 17), anchor="center")
        self.bienvenido.place(x=115,y=165)

        #Detecta nombre
        self.Label2 = Label(self.ventana, bg="#F38F00", fg='black')
        self.Label2.config(fg="red", font=("Arial", 17),anchor="center")
        self.Label2.place(x=182, y=220, anchor="center")
        #Llamando funciones
        self.encontrar()
        self.ventana.geometry("1360x690")
        self.active_cam()
        self.leer()
        self.ventana.mainloop()


    def encontrar(self):
        self.new_a = "".join(self.codelist)
        for i in self.new_data:
            if i == self.new_a:
                self.usuario = self.new_data[i]
                self.usuario_print = str(self.usuario)
                return self.usuario_print

    def nombre(self):
        self.nombre_usu = self.encontrar()
        self.ventana.after(10, self.nombre)

    def leer(self):
        self.a = self.encontrar()
        self.codelist.clear()
        self.Label2["text"] = self.a
        self.ventana.after(10,self.leer)

    #FIRMAR ENTRADA
    def firmar_entrada(self):
        self.archivo_excel = pd.read_excel('Data/Asistencia.xlsx', header=0)
        self.num_fila = self.archivo_excel.shape[0]
        self.hora = time.strftime("%H:%M")
        self.hora_sal = None
        self.num_colums = self.archivo_excel.shape[1]
        self.fecha = str(date.today())
        self.nomb = self.usuario_print

        if self.num_colums > 3:
            self.archivo_excel = self.archivo_excel.drop(self.archivo_excel.columns[[0]], axis='columns')
        self.archivo_excel.loc[self.num_fila] = [ self.nomb, self.hora,self.fecha,self.hora_sal]
        self.archivo_excel.drop(self.archivo_excel.columns[[0]], axis='columns')
        self.archivo_excel.to_excel("Data/Asistencia.xlsx")

    #FIRMAR SALIDA
    def firmar_salida(self):
        self.hora = time.strftime("%H:%M")
        self.fecha = str(date.today())
        self.nomb = self.usuario_print

        self.archivo_excel = pd.read_excel('Data/Asistencia.xlsx', header=0)
        self.num_colums = self.archivo_excel.shape[1]
        if self.num_colums > 3:
            self.archivo_excel = self.archivo_excel.drop(self.archivo_excel.columns[[0]], axis='columns')
        if self.num_colums > 1:
            self.archivo_excel.loc[((self.archivo_excel["Nombres"] == self.nomb) & (self.archivo_excel["Fecha"] == self.fecha))] = self.archivo_excel.loc[((self.archivo_excel["Nombres"] == self.nomb) & (self.archivo_excel["Fecha"] == self.fecha))].fillna(self.hora)
            self.archivo_excel.drop(self.archivo_excel.columns[[0]], axis='columns')
            self.archivo_excel.to_excel("Data/Asistencia.xlsx")

    def visor(self):
        ret, frame = self.get_frame()

        if ret:
            self.photo = ImageTk.PhotoImage(image=Image.fromarray(frame))
            self.canvas.create_image(0, 0, image=self.photo, anchor=NW)
            self.ventana.after(10, self.visor)

    def active_cam(self):
        if self.active_camera == False:
            self.active_camera = True
            self.VideoCaptura()
            self.visor()


    def capta(self, frm):
        self.info = decode(frm)
        cv2.putText(frm, "Muestre el codigo delante de la camara para su lectura", (53, 40), cv2.FONT_HERSHEY_SIMPLEX,
                    0.6, (255, 255, 0), 2)
        if self.info != []:
            for code in self.info:
                if code not in self.codelist:
                    content = code[0].decode('utf-8')
                    self.codelist.append(content)
                self.draw_rectangle(frm)

    def get_frame(self):
        if self.vid.isOpened():
            verif, frame = self.vid.read()
            if verif:
                self.capta(frame)
                return (verif, cv2.cvtColor(frame, cv2.COLOR_BGR2RGB))
            else:
                messagebox.showwarning("CAMARA NO DISPONIBLE", """La cámara está siendo utilizada por otra aplicación.
                Cierrela e intentelo de nuevo.""")
                self.active_cam()
                return (verif, None)
        else:
            verif = False
            return (verif, None)

    def draw_rectangle(self, frm):
        codes = decode(frm)
        for code in codes:
            data = code.data.decode('ascii')
            x, y, w, h = code.rect.left, code.rect.top, \
                         code.rect.width, code.rect.height
            cv2.rectangle(frm, (x, y), (x + w, y + h), (0, 255, 0), 6)
            cv2.putText(frm, code.type, (x - 10, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 50, 255), 2)

    def VideoCaptura(self):
        self.vid = cv2.VideoCapture(self.font_video)
        if self.vid.isOpened():
            self.width = self.vid.get(cv2.CAP_PROP_FRAME_WIDTH)
            self.height = self.vid.get(cv2.CAP_PROP_FRAME_HEIGHT)
            self.canvas.configure(width=self.width, height=self.height)
        else:
            messagebox.showwarning("CAMARA NO DISPONIBLE", "El dispositivo está desactivado o no disponible")
            self.active_camera = False

    def __del__(self):
        if self.active_camera == True:
            self.vid.release()

if __name__ == "__main__":
    App()