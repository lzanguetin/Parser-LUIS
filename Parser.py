class ProcessType(object):
	def __init__(self, regex=None, column=None):
		self.regex = regex
		self.column = column
		
import openpyxl
import re

from openpyxl import Workbook

fileName = input("Nome do arquivo:")
book = openpyxl.load_workbook(fileName)
sheet = book.active

i = 2
txt = sheet['A'+str(i)]

processTypes=[]
processTypes.append(ProcessType(r"\b\d{2}\.*\d{3}\.*\d{3}\/*\d{4}-*\d{2}\b", "C")) #CNPJ
processTypes.append(ProcessType(r"\b\d{3}\.*\d{3}\.*\d{3}-*\d{2}\b", "D")) #CPF
processTypes.append(ProcessType(r"\b\d{6}?\b", "E")) #SERIE
processTypes.append(ProcessType(r"\b9\d{3}\b", "F")) #PORTA
processTypes.append(ProcessType(r"\b\S+@\S+\.\S+\b", "G")) #EMAIL
processTypes.append(ProcessType(r"\b\d{2}\/\d{2}\b", "H")) #DATA
processTypes.append(ProcessType(r"\b[A-Za-z]{3}-*\s*\d{4}\b", "J")) #PLACA

print("Lendo...")

while (txt.value != None):
	i+=1
	txt = sheet['A'+str(i)]
	size = i-1

def processField():
	var = 2

	while (var <= size):
		for val in processTypes:
			txt = sheet['A'+str(var)]
			x = re.search(val.regex, str(txt.value))
	
			sheet[str(val.column)+str(var)] = x[0] if x else None
		var += 1

processField()

print("Salvando Arquivo")

book.save(fileName)

print("Arquivo Processado")