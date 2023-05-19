from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from pdf2docx import Converter
import os


pdf_file = ''
archivo = ''

def main():
	ventana = Tk()
	ventana.geometry("600x500")
	
	ventana.config(bg="darkslateblue")

	estilos = ttk.Style()

	estilos.theme_use("alt")

	estilos.configure("frames.TFrame",
					background = "grey22",
					)

	estilos.configure("titulo.TLabel",
					background = "grey22",
					foreground = "white",
					font = ("Calibri", 18),
					)

	estilos.configure("boton.TButton",
					background = "#5532F1",
					font = ("Calibri", 12),
					relief = 0,
					)
				
	frameContenido = ttk.Frame(ventana, style="frames.TFrame")
	frameContenido.pack(expand=1, ipadx= 180, ipady= 100)


	lbl1 = ttk.Label(frameContenido, style= "titulo.TLabel", text= "Fichero PDF: ")
	lbl1.pack(expand=1)

	ent = ttk.Entry(frameContenido, width= 60, state='readonly')
	ent.pack()

	def seleccionarArchivo():
		global archivo
		archivo = filedialog.askopenfilename(title="seleccionar", filetypes=(("Archivos pdf","*.pdf"),))
		ent.config(state='normal')
		ent.insert(END, archivo)
		ent.config(state='readonly')
		

	btn1 = ttk.Button(frameContenido,style="boton.TButton",width= 20,text= "Cargar un archivo", command=seleccionarArchivo)
	btn1.pack(expand=1)

	lbl2 = ttk.Label(frameContenido, style= "titulo.TLabel", text= "Convertir a DOCX: ")
	lbl2.pack(expand=1)

	def convertirADOCX():
		global pdf_file
		pdf_file = ''
		pdf_file = archivo
		docx_file = 'converted.docx'
		ent.config(state='normal')
		ent.delete(0, END)
		ent.config(state='readonly')

		cv = Converter(pdf_file)
		path = os. getcwd()
		docx = cv.convert(docx_file)
		cv.close()


	btn2 = ttk.Button(frameContenido, style="boton.TButton", width=20 ,text= "Guardar y Convertir", command= convertirADOCX)
	btn2.pack(expand=1)


	ventana.mainloop()

if __name__ == '__main__':
	main()


