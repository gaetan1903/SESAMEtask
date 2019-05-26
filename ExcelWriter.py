#!/usr/bin/python
""" 
Author: 	Gaetan Jonathan 
@email: gaetan.s118@gmail.com || gaetan.jonathan.bakary@esti.mg 
@facebook /gaetan1903

"""
# -*-coding:Utf-8 -*-

import time, datetime, xlsxwriter, webbrowser, os
from tkinter import *
from tkinter.font import *
import tkinter.messagebox as tkmsg 

Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


class Citation():
	""" une class qui afffecte une valeur au variable citation
 	dont le but de pouvoir la modifié facilement depuis l interface """
	def write_cit(self):
		self.citation = '"Avec des bonnes habitudes, viennent les bonnes attitudes"' 
		#  valeur par defaut de citation


class TacheNameChanger():
	""" une class qui affecte une liste de valeur au variable taches dont le but est 
de pouvoir modifier facilment depuis l Interface """
	def  t_changer(self):
		self.taches = ['Cuisine', 0, 'Table', 0, 'Vaisselle', 0, 'Refectoire']
		#  valeur par defaut de taches


class GrandMenage():
	""" une class qui affecte des valeurs a des variables dont le but est 
de pouvoir modifier facilement depuis l interface """
	def gm_changer(self):
		self.grandM1 = 'Frigo, aspirateur, fours a gaz, cuisiniere'
		self.grandM2 = 'Tables, leviers, carrelage'
		self.grandM3 = 'Vitre intérieur et mur à la cuisine'
		self.grandM4 = 'Intendance et laver les chiffons'
		self.grandM5 = 'Locaux ordures et dehors cuisines'
		self.grandM6 = "Tables, Chaises, toiles d'arraignés, et carrelage du refectoire"
		self.grandM7 = 'Charriots au refectoire (interieur et exterieur)'
		self.grandM8 = 'Vitres refectoire (interieur et exterieur)'
		#  leurs valeurs par defaut


cit = Citation() #  crée une instance pour la class Citation
ind = TacheNameChanger() #  cree une instance pour la class
gm = GrandMenage() #  crée une instance pour cetteclass


fileLog = open("logfile.gj", "r") #  ouvre un fichier nommé logfile.gj en mode lecture seule
count = 0 #  valeur initiale affecer
for read in fileLog:  #  met dans la variable read chaque ligne du fichier ouvert 
	exec(read)  #  execute les lignes du fichiers
	count += 1 #  incremente de 1 la valeur de count a chaque ligne du fichier (<==> compte le nombre de ligne)


fileLog.close() #  ferme le fichier ouvert 

class Fr_mois():
	""" une class pour convertir le mois en français """
	def month(self, jour): #  jour est un parametr du fonction month
			mois = {"01":"Janvier", "02":"Fevrier", "03":"Mars", "04":"Avril", "05":"Mai", "06":"Juin", "07":"Juillet", "08":"Août", "09":"Septembre", "10":"Octobre", "11":"Novembre", "12":"Decembre"}
			for x, y in mois.items(): # x parcours les chiffres et y parcours les lettres
				if str(datetime.date.strftime(datetime.date.fromordinal(jour), "%m")) == x: 
					return y
				#  si la valeur du parametr jour est chiffre quelquonque  , y sera le mois correspondante



class Write_constante():

	def __init__(self):

		if time.strftime('%a') == 'Mon':
			self.jour = int( datetime.date.toordinal(datetime.date.today()) ) - 1

		elif time.strftime('%a') == 'Tue':
			self.jour = int( datetime.date.toordinal(datetime.date.today()) ) - 2

		elif time.strftime('%a') == 'Wed':
			self.jour = int( datetime.date.toordinal(datetime.date.today()) ) - 3

		elif time.strftime('%a') == 'Thu':
			self.jour = int( datetime.date.toordinal(datetime.date.today()) ) - 4

		elif time.strftime('%a') == 'Fri':
			self.jour = int( datetime.date.toordinal(datetime.date.today()) ) - 5

		elif time.strftime('%a') == 'Sat':
			self.jour = int( datetime.date.toordinal(datetime.date.today()) ) - 6

		else: 
			self.jour = int( datetime.date.toordinal(datetime.date.today()) )
		
		mon = Fr_mois()

		self.dateName = str(datetime.date.strftime(datetime.date.fromordinal(self.jour), "%Y-%m-%d"))
		self.dateFin = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + 6), "%Y-%m-%d"))


	def initial(self):

		try:
			self.Fichier_Excel = xlsxwriter.Workbook(r"Output/Tache"+"-"+str(self.dateName)+".xlsx")
			
		except FileNotFoundError:
			os.system("mkdir Output")
			tkmsg.showinfo("Dossier", "Un Dossier Output a été créé... ")
			self.Fichier_Excel = xlsxwriter.Workbook(r"Output/Tache"+"-"+str(self.dateName)+".xlsx")

		except PermissionError:
			self.Fichier_Excel = xlsxwriter.Workbook(r"Output/Tache"+"-"+str(self.dateName)+"_1.xlsx")

		self.Feuille_Excel = self.Fichier_Excel.add_worksheet(self.dateName + '_' + self.dateFin)

		self.taches = ['Cuisine', 0, 'Table', 0, 'Vaisselle', 0, 'Refectoire']
		self.vagues = ['vague I', 'vague II']
		self.MMM = ["Matin", "Midi", "Soir"]
		self.frat = list("frat {}".format(x+1) for x in range(8))
		self.CM_S = ["APS 1", "APS 2"]
		self.CM_L = ["APL 1", "APL 2"]

		self.time_start = time.time()

		self.Feuille_Excel.set_landscape()
		self.Feuille_Excel.hide_gridlines(2)


	def initial_next(self):
		
		self.dateName = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + (7*self.nextChanger)), "%Y-%m-%d"))

		try:
			self.Fichier_Excel = xlsxwriter.Workbook(r"Output/Tache"+"-"+str(self.dateName)+".xlsx")
			
		except FileNotFoundError:
			os.system("mkdir Output")
			tkmsg.showinfo("Dossier", "Un Dossier Output a été créé... ")
			self.Fichier_Excel = xlsxwriter.Workbook(r"Output/Tache"+"-"+str(self.dateName)+".xlsx")

		except PermissionError:

			self.Fichier_Excel = xlsxwriter.Workbook(r"Output/Tache"+"-"+str(self.dateName)+"_1.xlsx")

		self.Feuille_Excel = self.Fichier_Excel.add_worksheet(self.dateName + '_' + self.dateFin)

		
		self.vagues = ['vague I', 'vague II']
		self.MMM = ["Matin", "Midi", "Soir"]
		self.frat = list("frat {}".format(x+1) for x in range(8))
		self.CM_S = ["APS 1", "APS 2"]
		self.CM_L = ["APL 1", "APL 2"]

		self.time_start = time.time()

		self.Feuille_Excel.set_landscape()
		self.Feuille_Excel.hide_gridlines(2)


	def format(self):
		self.merge_format_pT = self.Fichier_Excel.add_format({'align': 'center', 'bold': 1, 'fg_color': 'gray','border':1})
		self.fin_format = self.Fichier_Excel.add_format({'align': 'center', 'bold': 1, 'font_size':14})
		self.merge_format_TV = self.Fichier_Excel.add_format({'align': 'center', 'bold': 1, 'valign': 'vcenter', 'border':1})
		self.centrer = self.Fichier_Excel.add_format({'align': 'center','valign': 'vcenter', 'border': 1})
		self.cuisine_format = self.Fichier_Excel.add_format({'align': 'center','valign': 'vcenter', 'font_size': 10, 'border':1})
		self.MMM_format = self.Fichier_Excel.add_format({'align': 'center','bg_color': 'gray', 'border':1})


	def principal_text_next(self): 
		mon = Fr_mois()
		self.dateStart = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + 7 * self.nextChanger), "%d {} %Y".format(mon.month(self.jour + 7 * self.nextChanger))))
		self.dateEnd = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + 7 * self.nextChanger + 6), "%d {} %Y".format(mon.month(self.jour + 7 * self.nextChanger + 6))))
		self.Feuille_Excel.merge_range('D1:K1', 'Repartition des taches en AtriUM (semaine {}_{})'.format(self.dateStart, self.dateEnd), self.merge_format_pT)
		
		del self.dateStart
		del self.dateEnd


	def principal_text(self): 
		mon = Fr_mois()
		self.dateStart = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + 7 * self.nextChanger), "%d {} %Y".format(mon.month(self.jour + 7 * self.nextChanger))))
		self.dateEnd = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + 7 * self.nextChanger + 6), "%d {} %Y".format(mon.month(self.jour + 7 * self.nextChanger + 6))))
		self.Feuille_Excel.merge_range('D1:K1', 'Repartition des taches en AtriUM (semaine {}_{})'.format(self.dateStart, self.dateEnd), self.merge_format_pT)
		
		del self.dateStart
		del self.dateEnd


	def writeTache(self):
		self.Feuille_Excel.set_column(0, 0, 10) 	# Elargir le cote du premiere colonne
		self.Feuille_Excel.merge_range('A3:A4', 'Taches', self.merge_format_TV)
		for x in range(5, 13, 2):
			self.Feuille_Excel.merge_range('A{}:A{}'.format(x, x+1), ind.taches[x-5], self.merge_format_TV)

		""" _________________________________________x_________________________________________________ """

		self.Feuille_Excel.merge_range('A14:A15', 'Taches', self.merge_format_TV)
		for x in range(16, 24, 2):
			self.Feuille_Excel.merge_range('A{}:A{}'.format(x, x+1), ind.taches[x-16], self.merge_format_TV)
		del x 


	def writeVague(self):
		self.Feuille_Excel.set_column(1, 1, 10) 	# elargir le colonne
		self.Feuille_Excel.merge_range('B3:B4', 'Vagues', self.merge_format_TV)
		self.Feuille_Excel.write('B5', 'cuisine', self.cuisine_format)
		self.Feuille_Excel.write('B6', 'wc-Intce', self.cuisine_format)
		for x in range(6):
			self.Feuille_Excel.write(x + 6, 1, self.vagues[x%2], self.cuisine_format)
		""" ___________________________________x____________________________________ """

		self.Feuille_Excel.merge_range('B14:B15', 'Vagues', self.merge_format_TV)
		self.Feuille_Excel.write('B16', 'cuisine', self.cuisine_format)
		self.Feuille_Excel.write('B17', 'wc-Intce', self.cuisine_format)
		for x in range(6):
			self.Feuille_Excel.write(x + 17, 1, self.vagues[x % 2], self.cuisine_format)
		del x


	def writeDate(self):
		self.i = 0
		mon = Fr_mois()
		for x in range(0, 12, 3):

			if str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Mon":
				self.days = "Lundi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Tue":
				self.days = "Mardi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Wed":
				self.days = "Mercredi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Thu":
				self.days = "Jeudi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Fri":
				self.days = "Vendredi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Sat":
				self.days = "Samedi"

			else:
				self.days = "Dimanche"
		
			self.date = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%d {} %Y".format(mon.month(self.jour + self.i))))
			self.Feuille_Excel.merge_range('{}3:{}3'.format(Alphabet[x+2], Alphabet[x+4]), "{} {}".format(self.days, self.date), self.merge_format_TV)
			self.i += 1
		""" ______________________________________________________________x_______________________________________________ """

		for x in range(0, 9, 3):

			if str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Mon":
				self.days = "Lundi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Tue":
				self.days = "Mardi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Wed":
				self.days = "Mercredi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Thu":
				self.days = "Jeudi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Fri":
				self.days = "Vendredi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Sat":
				self.days = "Samedi"

			else:
				self.days = "Dimanche"
		
			self.date = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%d {} %Y".format(mon.month(self.jour + self.i))))
			self.Feuille_Excel.merge_range('{}14:{}14'.format(Alphabet[x+2], Alphabet[x+4]), "{} {}".format(self.days, self.date), self.merge_format_TV)
			self.i += 1
			
		del x
		del self.i


	def Write_nextDate(self):
		self.jour = self.jour + (7 * self.nextChanger)
		self.i = 0
		mon = Fr_mois()

		for x in range(0, 12, 3):

			if str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Mon":
				self.days = "Lundi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Tue":
				self.days = "Mardi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Wed":
				self.days = "Mercredi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Thu":
				self.days = "Jeudi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Fri":
				self.days = "Vendredi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Sat":
				self.days = "Samedi"

			else:
				self.days = "Dimanche"

			self.date = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%d {} %Y".format(mon.month(self.jour + self.i))))
			self.Feuille_Excel.merge_range('{}3:{}3'.format(Alphabet[x+2], Alphabet[x+4]), "{} {}".format(self.days, self.date), self.merge_format_TV)
			self.i += 1
			""" ______________________________________________________________x_______________________________________________ """

		for x in range(0, 9, 3):

			if str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Mon":
				self.days = "Lundi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Tue":
				self.days = "Mardi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Wed":
				self.days = "Mercredi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Thu":
				self.days = "Jeudi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Fri":
				self.days = "Vendredi"

			elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Sat":
				self.days = "Samedi"

			else:
				self.days = "Dimanche"

			self.date = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%d {} %Y".format(mon.month(self.jour + self.i))))
			self.Feuille_Excel.merge_range('{}14:{}14'.format(Alphabet[x+2], Alphabet[x+4]), "{} {}".format(self.days, self.date), self.merge_format_TV)
			self.i += 1 
	

	def writeMMM(self):
		for x in range(12):
			self.Feuille_Excel.write('{}4'.format(Alphabet[x+2]), self.MMM[x%3], self.MMM_format)
		""" __________________________________________x_________________________________________"""

		for x in range(9):
			self.Feuille_Excel.write('{}15'.format(Alphabet[x+2]), self.MMM[x%3], self.MMM_format)
		del x


	def writeGrandMenage(self):
		self.Feuille_Excel.merge_range('L14:N14', 'Samedi de 16h à 17h', self.merge_format_TV)
		self.Feuille_Excel.merge_range('L15:N15', 'Grand Menage', self.MMM_format)

		for x in range(16, 24):
			self.Feuille_Excel.merge_range('L{}:M{}'.format(x, x), 'GM Tache {}'.format(x - 15), self.centrer)


		self.Feuille_Excel.write("A25", "GM Tache 1:  '{}' ".format(gm.grandM1))
		self.Feuille_Excel.write("A26", "GM tache 2:  '{}' ".format(gm.grandM2))
		self.Feuille_Excel.write("A27", "GM Tache 3:  '{}' ".format(gm.grandM3))
		self.Feuille_Excel.write("A28", "GM Tache 4:  '{}' ".format(gm.grandM4))
		self.Feuille_Excel.write("A29", "GM Tache 5:  '{}' ".format(gm.grandM5))
		self.Feuille_Excel.write("A30", "GM Tache 6:  '{}' ".format(gm.grandM6))
		self.Feuille_Excel.write("A31", "GM Tache 7:  '{}' ".format(gm.grandM7))
		self.Feuille_Excel.write("A32", "GM Tache 8:  '{}' ".format(gm.grandM8))

		del x


	def sesame_logo(self):
		try:
			self.Feuille_Excel.insert_image("L25", "Images/Sesame.png")
		except:
			pass
	
		
	def select_date(self):
		self.date0 = datetime.date(2019, 1, 6)	 # La date supposée comme la date du debut, l'année 0 :-p
		self.date_x = datetime.date.today()  # la variable contenant la date actuel

		if self.date0 == self.date_x:
			return 0
		else:
			return int(str(self.date_x - self.date0).split(" ")[0])	 # Le nombre de jour entre la date actuel et la date initial

	def select_week(self):
		self.temps = self.select_date()

		return self.temps // 7


	def write_GMvariable(self):
		self.week = self.select_week()
		self.week = self.week + self.nextChanger

		y = 0
		for x in range((7*self.week), ((7*self.week) + 8)):
			self.Feuille_Excel.write("N{}".format(y + 16), self.frat[x % 8], self.centrer)
			y += 1
		del x


	def write_variable(self):
		self.Frat= ["frat 1", "frat 5", "frat 2", "frat 6", "frat 3", "frat 7", "frat 4", "frat 8","frat 5", "frat 1", "frat 6", "frat 2", "frat 7", "frat 3", "frat 8", "frat 4"]
		self.Frat_inverse = ["frat 5", "frat 1", "frat 6", "frat 2", "frat 7", "frat 3", "frat 8", "frat 4","frat 1", "frat 5", "frat 2", "frat 6", "frat 3", "frat 7", "frat 4", "frat 8",]
		self.week = self.select_week()
		self.week = self.week + self.nextChanger

		y = 0
		self.week = 2 * self.week

		for x in range((7*self.week), ((7*self.week) + 8)):

			self.Feuille_Excel.write("C{}".format(y + 5), self.Frat[x % 16], self.centrer)
			self.Feuille_Excel.write("E{}".format(y + 5), self.Frat_inverse[x % 16], self.centrer)
			self.Feuille_Excel.write("F{}".format(y + 5), self.Frat_inverse[(x +2) % 16], self.centrer)
			self.Feuille_Excel.write("H{}".format(y + 5), self.Frat[(x + 2) % 16], self.centrer)
			self.Feuille_Excel.write("I{}".format(y + 5), self.Frat[(x + 4) % 16], self.centrer)
			self.Feuille_Excel.write("K{}".format(y + 5), self.Frat_inverse[(x + 4) % 16], self.centrer)
			self.Feuille_Excel.write("L{}".format(y + 5), self.Frat_inverse[(x + 6) % 16], self.centrer)
			self.Feuille_Excel.write("N{}".format(y + 5), self.Frat[(x + 6) % 16], self.centrer)

			self.Feuille_Excel.write("C{}".format(y + 16), self.Frat_inverse[x % 16], self.centrer)
			self.Feuille_Excel.write("E{}".format(y + 16), self.Frat[x % 16], self.centrer)
			self.Feuille_Excel.write("F{}".format(y + 16), self.Frat[(x + 2) % 16], self.centrer)
			self.Feuille_Excel.write("H{}".format(y + 16), self.Frat_inverse[(x + 2) % 16], self.centrer)
			self.Feuille_Excel.write("I{}".format(y + 16), self.Frat_inverse[(x + 4) % 16], self.centrer)
			self.Feuille_Excel.write("K{}".format(y + 16), self.Frat[(x + 4) % 16], self.centrer)
		
			y += 1

		del x
		del y


	def write_varMidi(self):
		self.week = self.select_week()
		self.week = self.week + self.nextChanger
		y = 0
		for x in range((7*self.week), ((7*self.week) + 8)):
			self.Feuille_Excel.write("D{}".format(y + 5), self.frat[(x + 7) % 8], self.centrer)
			self.Feuille_Excel.write("J{}".format(y + 16), self.frat[(x + 1) % 8], self.centrer)
			y += 1

		for n in range(6, 13, 3):
			self.Feuille_Excel.write("{}5".format(Alphabet[n]), "Cuisiniers", self.centrer)
			self.Feuille_Excel.write("{}9".format(Alphabet[n]), "Cuisiniers", self.centrer)
			self.Feuille_Excel.write("{}10".format(Alphabet[n]), "Cuisiniers", self.centrer)

		for n in range(3, 7, 3):
			self.Feuille_Excel.write("{}16".format(Alphabet[n]), "Cuisiniers", self.centrer)
			self.Feuille_Excel.write("{}20".format(Alphabet[n]), "Cuisiniers", self.centrer)
			self.Feuille_Excel.write("{}21".format(Alphabet[n]), "Cuisiniers", self.centrer)

	
		self.autreList = ["G","J","M","D","G2"]

		
		for autreList in self.autreList:

			if autreList == 'J':
				self.Feuille_Excel.write("{}7".format(autreList), self.CM_S[(self.week + 1) % 2], self.centrer)
				self.Feuille_Excel.write("{}8".format(autreList), self.CM_L[(self.week + 1) % 2], self.centrer)				
				self.Feuille_Excel.write("{}11".format(autreList), self.CM_S[self.week % 2], self.centrer)
				self.Feuille_Excel.write("{}12".format(autreList), self.CM_L[self.week % 2], self.centrer)
				

			elif autreList == "G" or autreList =="M":
				self.Feuille_Excel.write("{}7".format(autreList), self.CM_S[self.week % 2], self.centrer)
				self.Feuille_Excel.write("{}8".format(autreList), self.CM_L[self.week % 2], self.centrer)
				self.Feuille_Excel.write("{}11".format(autreList), self.CM_S[(self.week + 1) % 2], self.centrer)	
				self.Feuille_Excel.write("{}12".format(autreList), self.CM_L[(self.week + 1) % 2], self.centrer)
					
			elif autreList == "G2":
				self.Feuille_Excel.write("G18", self.CM_S[self.week % 2], self.centrer)
				self.Feuille_Excel.write("G19", self.CM_L[self.week % 2], self.centrer)
				self.Feuille_Excel.write("G22", self.CM_S[(self.week + 1) % 2], self.centrer)	
				self.Feuille_Excel.write("G23", self.CM_L[(self.week + 1) % 2], self.centrer)

			else:
				self.Feuille_Excel.write("{}18".format(autreList), self.CM_S[(self.week + 1) % 2], self.centrer)
				self.Feuille_Excel.write("{}19".format(autreList), self.CM_L[(self.week + 1) % 2], self.centrer)				
				self.Feuille_Excel.write("{}22".format(autreList), self.CM_S[self.week % 2], self.centrer)
				self.Feuille_Excel.write("{}23".format(autreList), self.CM_L[self.week % 2], self.centrer)
				
		del y
		del autreList
		del self.autreList
		del n 


	def __fin__(self):

		self.Feuille_Excel.merge_range('D34:L34', cit.citation, self.fin_format)
		self.Feuille_Excel.write("D6", None, self.centrer)
		self.Feuille_Excel.write("J17", None, self.centrer)
		self.Fichier_Excel.close()
		self.time_end = time.time()
		self.duree = self.time_end - self.time_start
		self.duree = float("%.2f" %self.duree)
		tkmsg.showinfo("Succes", "La tâche a été génerer en {} secondes avec succes".format(self.duree))

	


class Interface():
	
	def __init__(self):
		self.root = Tk()
		self.root.title("SESAME TASK AtriuM")
		#self.root.iconbitmap(r'Images\\favicon.ico')
		self.root.geometry("1300x625+20+20")
		filelog = open(".theme.gj", "r")
		self.count_changeTheme = 0
		for read in filelog:
			exec(read)
			self.count_changeTheme += 1
		filelog.close()
		self.nombreErreur = 1
		

	def get_start(self):

		excel = Write_constante()
		excel.initial()
		excel.format()
		excel.nextChanger = 0
		excel.principal_text()
		excel.writeTache()
		excel.writeVague()
		excel.writeDate()
		excel.writeMMM()
		excel.writeGrandMenage()
		excel.sesame_logo()
		excel.select_date()
		excel.select_week()
		excel.write_GMvariable()
		excel.write_variable()
		excel.write_varMidi()
		excel.__fin__()



		if self.nombreErreur < 2:
			tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Quelque chose a mal fonctionner")
			self.nombreErreur += 1

		elif self.nombreErreur == 2:
			tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Verifier s'il y pas de fichier Tache excel ouvert, si oui Fermer")
			self.nombreErreur += 1

		elif self.nombreErreur > 2 and self.nombreErreur < 5:
			tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Verifier si un dossier 'Output' est présent dans le dossier contenant l'excecutable, sinon créé")
			self.nombreErreur += 1
		else:
			tkmsg.showerror("Erreur", 'Veuiller Contacter le Developper pour resoudre le probleme')
			
			
	def get_next(self):
		
		try:
			excel = Write_constante()
			excel.nextChanger = 1
			excel.initial_next()
			excel.format()
			excel.principal_text_next()
			excel.writeTache()
			excel.writeVague()
			excel.Write_nextDate()
			excel.writeMMM()
			excel.writeGrandMenage()
			excel.sesame_logo()
			excel.select_date()
			excel.select_week()
			excel.write_GMvariable()
			excel.write_variable()
			excel.write_varMidi()
			excel.__fin__()
		
		except:
			if self.nombreErreur < 2:
				tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Quelque chose a mal fonctionner")
				self.nombreErreur += 1

			elif self.nombreErreur == 2:
				tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Verifier s'il y pas de fichier Tache excel ouvert, si oui Fermer")
				self.nombreErreur += 1

			elif self.nombreErreur > 2 and self.nombreErreur < 5:
				tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Verifier si un dossier 'Output' est présent dans le dossier contenant l'excecutable, sinon créé")
				self.nombreErreur += 1

			else:
				tkmsg.showerror("Erreur", 'Veuiller Contacter le Developper pour resoudre le probleme')


	def get_next_next(self):

		try: 
			test = int(self.entry_champ.get())
			test = True

		except ValueError:
			test = False
	
		if test:
			try:
				excel = Write_constante()
				excel.nextChanger = int(self.entry_champ.get())
				excel.initial_next()
				excel.format()
				excel.principal_text_next()
				excel.writeTache()
				excel.writeVague()
				excel.Write_nextDate()
				excel.writeMMM()
				excel.writeGrandMenage()
				excel.sesame_logo()
				excel.select_date()
				excel.select_week()
				excel.write_GMvariable()
				excel.write_variable()
				excel.write_varMidi()
				excel.__fin__()
				
			except:

				if self.nombreErreur < 2:
					tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Quelque chose a mal fonctionner")
					self.nombreErreur += 1

				elif self.nombreErreur == 2:
					tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Verifier s'il y pas de fichier Tache excel ouvert, si oui Fermer")
					self.nombreErreur += 1

				elif self.nombreErreur > 2 and self.nombreErreur < 5:
					tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Verifier si un dossier 'Output' est présent dans le dossier contenant l'excecutable, sinon créé")
					self.nombreErreur += 1

				else:
					tkmsg.showerror("Erreur", 'Veuiller Contacter le Developper pour resoudre le probleme')
		else:
			tkmsg.showerror("Erreur d'entrée","Il semblerait que vous avez entrer un type non valable \n L'entrée doit être un entier")

		del test


	def citation_changer(self):
		return self.entry_citation.get()


	def citationChanger(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')

		cit.citation = self.citation_changer()
		fileLog.write("\n"+"cit.citation = "+"'{}'".format(cit.citation))
		tkmsg.showinfo("Changement Effectuer", 'La citation de cette semaine a été changer en " {} " avec succes '.format(cit.citation))

		fileLog.close()


	def citation_restore(self):
		return '"Avec des bonnes habitudes viennent les bonnes attitudes"'


	def citationRestore(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')

		cit.citation = str(self.citation_restore())
		fileLog.write("\n"+"cit.citation = "+"'{}'".format(cit.citation))
		tkmsg.showinfo("Restauration Effectuer", 'La citation de cette semaine a été restaurer en " {} " avec succes '.format(cit.citation))

		fileLog.close()


	def changetache1(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')

		ind.taches[0] = str(self.entry_tache1.get())
		fileLog.write("\n"+"ind.taches[0] = "+"'{}'".format(ind.taches[0]))
		tkmsg.showinfo("Tache Changer", 'La tache a été changer en " {} " avec succes'.format(str(self.entry_tache1.get())))

		fileLog.close()


	def changetache2(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')

		ind.taches[2] = str(self.entry_tache2.get())
		fileLog.write("\n"+"ind.taches[2] = "+"'{}'".format(ind.taches[2]))
		tkmsg.showinfo("Tache Changer", 'La tache a été changer en " {} " avec succes'.format(str(self.entry_tache2.get())))

		fileLog.close()


	def changetache3(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')

		ind.taches[4] = str(self.entry_tache3.get())
		fileLog.write("\n"+"ind.taches[4] = "+"'{}'".format(ind.taches[4]))
		tkmsg.showinfo("Tache Changer", 'La tache a été changer en " {} " avec succes'.format(str(self.entry_tache3.get())))

		fileLog.close()


	def changetache4(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')

		ind.taches[6] = str(self.entry_tache4.get())
		fileLog.write("\n"+"ind.taches[6] = "+"'{}'".format(ind.taches[6]))
		tkmsg.showinfo("Tache Changer", 'La tache a été changer en " {} " avec succes'.format(str(self.entry_tache4.get())))

		fileLog.close()


	def grandM_changer1(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')
		
		gm.grandM1 = self.entry_gm1.get()
		tkmsg.showinfo("Changement de Tâche", "La tâche a été changer en '{}' avec succes".format(gm.grandM1))
		fileLog.write("\n"+"gm.grandM1 = '{}'".format(gm.grandM1))

		fileLog.close()

	
	def grandM_changer2(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')
		
		gm.grandM2 = self.entry_gm2.get()
		tkmsg.showinfo("Changement de Tâche", "La tâche a été changer en '{}' avec succes".format(gm.grandM2))
		fileLog.write("\n"+"gm.grandM2 = '{}'".format(gm.grandM2))

		fileLog.close()

	
	def grandM_changer3(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')
		
		gm.grandM3 = self.entry_gm3.get()
		tkmsg.showinfo("Changement de Tâche", "La tâche a été changer en '{}' avec succes".format(gm.grandM3))
		fileLog.write("\n"+"gm.grandM3 = '{}'".format(gm.grandM3))

		fileLog.close()


	def grandM_changer4(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')
		
		gm.grandM4 = self.entry_gm4.get()
		tkmsg.showinfo("Changement de Tâche", "La tâche a été changer en '{}' avec succes".format(gm.grandM4))
		fileLog.write("\n"+"gm.grandM4 = '{}'".format(gm.grandM4))

		fileLog.close()


	def grandM_changer5(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')
		
		gm.grandM5 = self.entry_gm5.get()
		tkmsg.showinfo("Changement de Tâche", "La tâche a été changer en '{}' avec succes".format(gm.grandM5))
		fileLog.write("\n"+"gm.grandM5 = '{}'".format(gm.grandM5))

		fileLog.close()


	def grandM_changer6(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')
		
		gm.grandM6 = self.entry_gm6.get()
		tkmsg.showinfo("Changement de Tâche", "La tâche a été changer en '{}' avec succes".format(gm.grandM6))
		fileLog.write("\n"+"gm.grandM6 = '{}'".format(gm.grandM6))

		fileLog.close()

	
	def grandM_changer7(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')
		
		gm.grandM7 = self.entry_gm7.get()
		tkmsg.showinfo("Changement de Tâche", "La tâche a été changer en '{}' avec succes".format(gm.grandM7))
		fileLog.write("\n"+"gm.grandM7 = '{}'".format(gm.grandM7))

		fileLog.close()


	def grandM_changer8(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')
		
		gm.grandM8 = self.entry_gm8.get()
		tkmsg.showinfo("Changement de Tâche", "La tâche a été changer en '{}' avec succes".format(gm.grandM8))
		fileLog.write("\n"+"gm.grandM8 = '{}'".format(gm.grandM8))

		fileLog.close()



	def reset_taches(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')

		ind.t_changer() 
		tkmsg.showinfo("Restaurer les tâches", "Les tâches on été restaurer avec succès")
		fileLog.write("\n"+"ind.t_changer")
		

		fileLog.close()

	
	def reset_gm_taches(self):
		fileLog = open("logfile.gj", "a", encoding = 'Utf-8')

		gm.gm_changer()
		fileLog.write("\n"+"ind.t_changer")
		tkmsg.showinfo("Restaurer les tâches du Grand Menage", "Les tâches du grand menage ont été restaurer avec succes")

		fileLog.close()


	def reset_all(self):
		fileLog = open("logfile.gj", "w", encoding = 'Utf-8')
		fileLog.write( "cit.write_cit()" + "\n"+"ind.t_changer()" + "\n"+"gm.gm_changer()" )
		tkmsg.showinfo("Restaurer tous", "Tout les changements seront restaurer après redemarrage du logiciel")
		fileLog.close()


	def menu_button(self):

		self.menubutton = Menu(self.root)

		self.sous_menubutton_1 = Menu(self.menubutton, tearoff =0)
		self.sous_menubutton_2 = Menu(self.menubutton, tearoff =0)
		self.sous_menubutton_3 = Menu(self.menubutton, tearoff =0)
		self.sous_menubutton_4 = Menu(self.menubutton, tearoff =0)
		self.sous_menubutton_5 = Menu(self.menubutton, tearoff =0)

		self.menubutton.add_cascade(label = "Fichier"  , menu = self.sous_menubutton_1)
		self.menubutton.add_cascade(label = "Restaurer"  , menu = self.sous_menubutton_3)
		self.menubutton.add_cascade(label = "Theme"  , menu = self.sous_menubutton_4)
		self.menubutton.add_cascade(label = "Documentation"  , menu = self.sous_menubutton_5)
		self.menubutton.add_cascade(label = "Aide"  , menu = self.sous_menubutton_2)
		

		self.sous_menubutton_1.add_command(label ="Nouvelle fenetre", command = self.new_fen)
		self.sous_menubutton_1.add_command(label ="Quitter", command = self.Confirmer)

		self.sous_menubutton_2.add_command(label ="A propos du logiciel", command = self.Apropos_Logiciel)
		self.sous_menubutton_2.add_command(label ="A propos du programmeur", command = self.Apropos_Developper)
		self.sous_menubutton_2.add_command(label ="Open Source", command = self.CodeSource)

		self.sous_menubutton_3.add_command(label ="Restaurer les taches", command = self.reset_taches)
		self.sous_menubutton_3.add_command(label ="Restaurer les taches du Grand Menages", command = self.reset_gm_taches)
		self.sous_menubutton_3.add_command(label ="Restaurer tous",command = self.reset_all)

		self.sous_menubutton_4.add_command(label ="Light (Default)", command = self.fond_blanc)
		self.sous_menubutton_4.add_command(label ="Dark", command = self.fond_noir)
		self.sous_menubutton_4.add_command(label ="Fond Jaune",command = self.fond_jaune)
		self.sous_menubutton_4.add_command(label ="Fond Bleu",command = self.fond_bleu)

		self.sous_menubutton_5.add_command(label ="Lisez-Moi", command = self.readme)
		self.sous_menubutton_5.add_command(label ="Me Contatcer", command = self.contact)
		

		self.root.config(menu = self.menubutton)

	def contact(self):
		webbrowser.open("www.facebook.com/gaetan1903",new = 2, autoraise = True)


	def readme(self):
		webbrowser.open("index.html", new = 3, autoraise = True )


	def fond_blanc(self):
		fileLog = open(".theme.gj", "w", encoding = 'Utf-8')	
		self.root['bg'] = 'white'
		fileLog.write("self.root['bg'] = 'white'")
		fileLog.close()


	def fond_noir(self):
		fileLog = open(".theme.gj", "w", encoding = 'Utf-8')	
		self.root['bg'] = 'black'
		fileLog.write("self.root['bg'] = 'black'")
		fileLog.close()


	def fond_jaune(self):
		fileLog = open(".theme.gj", "w", encoding = 'Utf-8')	
		self.root['bg'] = 'yellow'
		fileLog.write("self.root['bg'] = 'yellow'")
		fileLog.close()


	def fond_bleu(self):
		fileLog = open(".theme.gj", "w", encoding = 'Utf-8')	
		self.root['bg'] = 'blue'
		fileLog.write("self.root['bg'] = 'blue'")
		fileLog.close()


	def new_fen(self):
		fen = Interface()
		fen.menu_button()
		fen.labeltext()
		fen.button()
		fen.label_bord()
		fen.__final__()


	def Apropos_Logiciel(self):
			tkmsg.showinfo("A propos de ce logiciel", "C'est un simple outil permettant de generer la liste des Tache en AtriUM SESAME.")


	def Apropos_Developper(self):
			tkmsg.showinfo("A propos du Developper", " Nom: \t BAKARY \n Prenoms: Gaetan Jonathan \n Ville: \t Toamasina \n Promo: \t SESAME P18 \n Ecole: \t ESTI Antanimena \n Mail(1): \t gaetan.s118@gmail.com \n Mail(2): \t gaetan.jonathan.bakary@esti.mg ")


	def Confirmer(self):
		self.fermer = tkmsg.askquestion("Confirmer la fermeture!", "Voulez-vous vraiment quitter?")
		if self.fermer == "yes":
			self.root.quit()
		else:
			pass


	def CodeSource(self):
			tkmsg.showinfo("Code Source", 'Le code source est à voir sur le site \n "https://github.com/gaetan1903/SESAMEtache" ')
			webbrowser.open("..\\dist\\cache.gj", new = 2, autoraise = True)
			

	def button(self):
		self.button_start_font = Font(family ="Times New Rowan", size = -20, weight = "bold")
		self.button_start = Button(self.root, text = "GENERER LA TACHE", font = self.button_start_font, bg = "gray", height = 3, width = 20, fg = 'green',  command = self.get_start)
		self.button_start.place(x ="525", y = "475")

		self.button_start_font = Font(family ="Times New Rowan", size = -20, weight = "bold")
		self.button_start = Button(self.root, text = "GENERER LA TACHE DE LA SEMAINE PROCHAINE", font = self.button_start_font, bg = "gray", height = 3, width = 45, fg = 'green',  command = self.get_next)
		self.button_start.place(x ="385", y = "575")


	def labeltext(self):
		mon = Fr_mois()
		self.table_bord = Label(self.root, text = " ", width = 70, height = 100, bg = '#e8e8c8')
		self.table_bord.place(x="1100", y="50")

		self.textNone = Label(self.root, text = " ", width = 200, height = 4, bg = "grey").place(x ='0', y = '0')

		self.sesame_icone = PhotoImage(file='Images/favicon.png')
		self.sesame_icon_label = Label(self.root, image = self.sesame_icone)
		self.sesame_icon_label.place(x = '0', y = '4')

		self.gj_image = PhotoImage(file='Images/gj.png')
		self.gj_label = Label(self.root, image = self.gj_image)
		self.gj_label.place(x = '1320', y = '5')

		self.font_date_actuel = Font(family ="Times New Rowan", size = -20, weight = "bold", underline = True)
		self.date_actuel = Label(self.root, text = time.strftime("%d {} %Y".format(mon.month(datetime.date.toordinal(datetime.date.today())))), fg = "orange",  bg = "grey", height = 2, width = '15', font = self.font_date_actuel)
		self.date_actuel.place(x = "550", y = "2") 

		self.photo = PhotoImage(file = "Images/Sesame.png")
		self.images_logo = Label(self.root, image = self.photo).place(x = "0", y= "510")

		self.photo_decor = PhotoImage(file = "Images/priv.png")
		self.images_decor = Label(self.root, image = self.photo_decor).place(x="385", y='70')

		self.taches_photo = PhotoImage(file = "Images/logo1.png")
		self.taches_image = Label(self.root, image = self.taches_photo).place(x = "0", y ="65")


	def label_bord(self):
		self.font_text1 = Font(family = "Arial", size = -12, weight = "bold", underline = True)
		self.text1 = Label(self.root, text = "Avancer de plusieurs semaines?", bg = "white", font = self.font_text1)
		self.text1.place(x = "1128", y = "70")

		self.entry_variable = IntVar()
		self.entry_champ = Entry(self.root, textvariable = self.entry_variable, width = 5)
		self.entry_champ.place(x="1230", y ="100")
		self.entry_text_s = Label(self.root,font = Font(family = 'Arial', size = -14, weight = 'bold'), bg = "#e8e8c8", text = "Avancer de ").place(x="1135", y ="100")
		self.entry_button = Button(self.root, text= "GENERER",font = Font(family = 'Arial', size = -14, weight = 'bold'), command = self.get_next_next ).place(x="1130", y="125")
		self.entry_text_e = Label(self.root, font = Font(family = 'Arial', size = -14, weight = 'bold'), bg = "#e8e8c8", text = "semaines").place(x="1265", y ="100")

		self.text_citation = Label(self.root, text = "Changer la citation de la semaine ?", bg = "white", font = self.font_text1)
		self.text_citation.place(x = "1128", y = "175")

		self.entry_citation_variable = StringVar()
		self.entry_citation = Entry(self.root, textvariable = self.entry_citation_variable, bg = "white", font = self.font_text1, width = 30)
		self.entry_citation.place(x = "1128", y = "200")

		self.entry_citation_button_change = Button(self.root, text = "CHANGER", font = Font(family = 'Arial', size = -14, weight = 'bold'), command = self.citationChanger )
		self.entry_citation_button_change.place(x="1130", y="225")
		self.entry_citation_button_restore = Button(self.root, text = "RESTAURER", font = Font(family = 'Arial', size = -14, weight = 'bold'), command = self.citationRestore )
		self.entry_citation_button_restore.place(x = "1230", y="225")

		self.text_tache = Label(self.root, text = "Changer les  Tâches?", bg = "white", font = self.font_text1)
		self.text_tache.place(x = "1128", y = "275")

		font_text_tache = Font(family = "Times New Rowan", size = -12)

		self.text_tache1 = Label(self.root, text = "T1 (Cuisine): ", font = font_text_tache )
		self.text_tache1.place(x = "1128", y = "300")
		self.entry_tache1 = Entry(self.root, textvariable = StringVar(), bg = "white", font = self.font_text1, width = 15)
		self.entry_tache1.place(x = "1210", y = "301")
		self.button_tache1 = Button(self.root, text = "OK", font = Font(family = "Arial", size = -11 , weight = "bold"),command = self.changetache1)
		self.button_tache1.place(x = "1325", y ="297")

		self.text_tache2 = Label(self.root, text = "T2 (Table): ", font = font_text_tache )
		self.text_tache2.place(x = "1128", y = "335")
		self.entry_tache2 = Entry(self.root, textvariable = StringVar(), bg = "white", font = self.font_text1, width = 15)
		self.entry_tache2.place(x = "1210", y = "336")
		self.button_tache2 = Button(self.root, text = "OK", font = Font(family = "Arial", size = -11 , weight = "bold"), command = self.changetache2)
		self.button_tache2.place(x = "1325", y ="332")

		self.text_tache3 = Label(self.root, text = "T3 (Vaisselle): ", font = font_text_tache )
		self.text_tache3.place(x = "1128", y = "370")
		self.entry_tache3 = Entry(self.root, textvariable = StringVar(), bg = "white", font = self.font_text1, width = 15)
		self.entry_tache3.place(x = "1210", y = "371")
		self.button_tache3 = Button(self.root, text = "OK", font = Font(family = "Arial", size = -11 , weight = "bold"), command = self.changetache3)
		self.button_tache3.place(x = "1325", y ="367")

		self.text_tache4 = Label(self.root, text = "T4 (Refectoire): ", font = font_text_tache )
		self.text_tache4.place(x = "1128", y = "405")
		self.entry_tache4 = Entry(self.root, textvariable = StringVar(), bg = "white", font = self.font_text1, width = 15)
		self.entry_tache4.place(x = "1210", y = "406")
		self.button_tache4 = Button(self.root, text = "OK", font = Font(family = "Arial", size = -11 , weight = "bold"), command = self.changetache4)
		self.button_tache4.place(x = "1325", y ="402")
		
		self.text_gm = Label(self.root, text = "Changer les Taches du Grand Menage?", bg = "white", font = self.font_text1)
		self.text_gm.place(x = "1128", y = "452")

		self.text_gm1 = Label(self.root, text ="1", font = font_text_tache)
		self.text_gm1.place(x="1128", y="485")
		self.entry_gm1 = Entry(self.root, textvariable = StringVar(), bg = "white", width = 29)
		self.entry_gm1.place(x="1145", y="487")
		self.button_gm1 = Button(self.root, text = "OK", font = Font(family = "Arial", size = -11 , weight = "bold"), command = self.grandM_changer1)
		self.button_gm1.place(x="1325", y="483")
		
		self.text_gm2 = Label(self.root, text ="2", font = font_text_tache)
		self.text_gm2.place(x="1128", y="510")
		self.entry_gm2 = Entry(self.root, textvariable = StringVar(), bg = "white", width = 29)
		self.entry_gm2.place(x="1145", y="512")
		self.button_gm2 = Button(self.root, text = "OK", font = Font(family = "Arial", size = -11 , weight = "bold"), command = self.grandM_changer2)
		self.button_gm2.place(x="1325", y="513")

		self.text_gm3 = Label(self.root, text ="3", font = font_text_tache)
		self.text_gm3.place(x="1128", y="535")
		self.entry_gm3 = Entry(self.root, textvariable = StringVar(), bg = "white", width = 29)
		self.entry_gm3.place(x="1145", y="537")
		self.button_gm3 = Button(self.root, text = "OK", font = Font(family = "Arial", size = -11 , weight = "bold"), command = self.grandM_changer3)
		self.button_gm3.place(x="1325", y="538")

		self.text_gm4 = Label(self.root, text ="4", font = font_text_tache)
		self.text_gm4.place(x="1128", y="560")
		self.entry_gm4 = Entry(self.root, textvariable = StringVar(), bg = "white", width = 29)
		self.entry_gm4.place(x="1145", y="562")
		self.button_gm4 = Button(self.root, text = "OK", font = Font(family = "Arial", size = -11 , weight = "bold"), command = self.grandM_changer4)
		self.button_gm4.place(x="1325", y="563")

		self.text_gm5 = Label(self.root, text ="5", font = font_text_tache)
		self.text_gm5.place(x="1128", y="585")
		self.entry_gm5 = Entry(self.root, textvariable = StringVar(), bg = "white", width = 29)
		self.entry_gm5.place(x="1145", y="587")
		self.button_gm5 = Button(self.root, text = "OK", font = Font(family = "Arial", size = -11 , weight = "bold"), command = self.grandM_changer5)
		self.button_gm5.place(x="1325", y="588")

		self.text_gm6 = Label(self.root, text ="6", font = font_text_tache)
		self.text_gm6.place(x="1128", y="610")
		self.entry_gm6 = Entry(self.root, textvariable = StringVar(), bg = "white", width = 29)
		self.entry_gm6.place(x="1145", y="612")
		self.button_gm6 = Button(self.root, text = "OK", font = Font(family = "Arial", size = -11 , weight = "bold"), command = self.grandM_changer6)
		self.button_gm6.place(x="1325", y="613")

		self.text_gm7 = Label(self.root, text ="7", font = font_text_tache)
		self.text_gm7.place(x="1128", y="635")
		self.entry_gm7 = Entry(self.root, textvariable = StringVar(), bg = "white", width = 29)
		self.entry_gm7.place(x="1145", y="637")
		self.button_gm7 = Button(self.root, text = "OK", font = Font(family = "Arial", size = -11 , weight = "bold"), command = self.grandM_changer7)
		self.button_gm7.place(x="1325", y="638")

		self.text_gm8 = Label(self.root, text ="8", font = font_text_tache)
		self.text_gm8.place(x="1128", y="660")
		self.entry_gm8 = Entry(self.root, textvariable = StringVar(), bg = "white", width = 29)
		self.entry_gm8.place(x="1145", y="662")
		self.button_gm8 = Button(self.root, text = "OK", font = Font(family = "Arial", size = -11 , weight = "bold"), command = self.grandM_changer8)
		self.button_gm8.place(x="1325", y="663")



		del self.font_text1
		del font_text_tache
		

	def __final__(self):
		self.root.mainloop()
		


if __name__ == "__main__":
	fen = Interface()
	fen.menu_button()
	fen.labeltext()
	fen.button()
	fen.label_bord()
	if count >= 40:
		fen.reset_all()

	fen.__final__()
