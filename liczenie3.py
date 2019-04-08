# -*- coding: utf-8 -*-
"""
Created on Thu Feb 14 14:37:38 2019

@author: marek-mleczko
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Jan 10 10:24:54 2019

@author: marek-mleczko
"""
import time
import docx
#import excel_reading as ex
import math
import pandas as pd
import numpy as np
#from matplotlib import pyplot as mp
from openpyxl import load_workbook

start_time = time.time()
    
tab=[]
row = 0
qb1 = pd.read_excel('mdc/QB.xlsm', sheet_name = "Compensation P&P")
question_bank_cc1 = pd.read_excel('mdc/QB.xlsm', sheet_name = "Company Cars")
qb2 = pd.read_excel('mdc/QB.xlsm', sheet_name = "Retirement")
qb3 = pd.read_excel('mdc/QB.xlsm', sheet_name = "Insurance Medical")
qb4 = pd.read_excel('mdc/QB.xlsm', sheet_name = "Other Benefits")


ben = pd.read_excel('mdc/2018 Example Databases/DK/POLPRAC2.2.xlsx')
ben_cc = pd.read_excel('mdc/2018 Example Databases/DK/BENEFITS2.xlsx')
#ben_cc.applymap(str)
ben1 = pd.read_excel('mdc/2018 Example Databases/DK/BENEFITS1.xlsx')
ben2 = pd.read_excel('mdc/2018 Example Databases/DK/BENEFITS3.xlsx')

frames =[question_bank_cc1, qb1, qb2, qb3, qb4]
question_bank = pd.concat(frames)
#ben4 = pd.read_excel('mdc/2018 Example Databases/DK/BENEFITS4.xlsx')


#ben5 = pd.read_excel('mdc/2018 Example Databases/DK/BENEFITS5.xlsx')
#ben6 = pd.read_excel('mdc/2018 Example Databases/DK/BENEFITS6.xlsx')
#ben7 = pd.read_excel('mdc/2018 Example Databases/DK/BENEFITS7.xlsx')
#ben8 = pd.read_excel('mdc/2018 Example Databases/DK/BENEFITS8.xlsx')

dimensionsets = pd.read_excel('mdc/dimension sets.xlsx')


a = question_bank.loc[question_bank['Countries'].str.contains("DK")]
#c = question_bank_cc.loc[question_bank_cc['Countries'].str.contains("DK")] 
dimensionssets = dimensionsets.loc[dimensionsets['countries'].str.contains("DK")]
#puste = ben_cc.isnull().sum(axis = 1)


class question():
        def __init__(self, QBcode, dimensionSetName, Label, referenceTable, referenceTableLabels, Type_of_question):
            self.freq=[]
            self.QBcode = QBcode
            #self.baza = baza
            self.dimensionSetName =dimensionSetName
            self.Label = Label
            self.referenceTable = str(referenceTable)
            self.referenceTablesplit = self.referenceTable.split('|')
            if "NA" in self.referenceTablesplit:
                    del self.referenceTablesplit[-1]
            
            self.referenceTableLabels = str(referenceTableLabels)
            self.referenceTableLabelssplit = self.referenceTableLabels.split('|')
            if "Not applicable" in self.referenceTableLabelssplit:
                    del self.referenceTableLabelssplit[-1]
            
            self.Type_of_question = Type_of_question
            self.grupy_pracownicze = []
            self.dane= []
            self.labele_grup = []
            self.dimensions = []
            self.grupy = self.ilegrup()    
            self.ile_grup = len(self.grupy_pracownicze)
            self.ile_kolumn = len(self.referenceTablesplit)
            #self.dane4 = [] 
            self.dane6 = [] #[] for y in range(len(self.dimensions))
            
            self.o = self.ktory_obiekt()
            
            self.dane1 =[]
            if self.Type_of_question == "radio_buttons" and self.o is not None:
                self.dane = self.laduj_dane_text()
                self.wyniki = self.liczenie_RB()
            elif self.Type_of_question == "text" and self.o is not None:   
                self.dane = self.laduj_dane_text()
                self.wyniki = self.liczenie_text()
            elif (self.Type_of_question == "double" or self.Type_of_question == "percentage") and self.o is not None:   
                self.dane = self.laduj_dane_num()
                self.wyniki = self.liczenie_percentile() 
            elif self.Type_of_question == "checkboxes" and self.o is not None:   
                self.dane = self.laduj_dane_checkboxes()
                self.wyniki = self.liczenie_checkboxes()          
            else:
                self.wyniki=[]
                self.wyniki1=[]
                
#                self.dane = self.laduj_dane_nowy_typ and self.o is not None() 
#                self.wyniki = self.liczenie()
        
        

               
            
            self.wyniki=[]
            #self.tabelaprocenty, self.odp, self.odpcount = self.procenty(self.wyniki, self.ile_kolumn )
            self.insufficient = self.show_table(self.freq)
            
            
            self.sekcja = "Company Cars"
            if (self.Type_of_question == "checkboxes" or self.Type_of_question == "radio_buttons") and self.o is not None:
                self.rounding = self.checkrounding()
            if self.Type_of_question == "checkboxes" and self.o is not None:
                self.exceeds100 = self.exceeds100()
            self.example = False
            self.section_name = False
            self.subsection_name = False
            
            if  len(self.wyniki1)!=0:
                self.mozna = True 
            elif self.wyniki1 is None:     
                self.mozna = False
            else: 
                self.mozna = False
        
        
        def show_table(self, odp):
            maska= bool()
            if self.Type_of_question == "Text":
                return maska
            else:
               if all([x[-1] < 3 for x in odp]):
                   maska= True
               return maska         
      
        def ktory_obiekt(self):
            if self.QBcode in ben_cc.columns:  
                return ben_cc
            elif self.QBcode in ben.columns:
                return ben
            elif self.QBcode in ben1.columns:
                return ben1
            elif self.QBcode in ben2.columns:
                return ben2
        
        def checkrounding(self):
            rounding = bool()
            for x in self.wyniki1_round:  
                if sum(x)%100 !=0:
                    rounding =True
            return rounding
        def exceeds100(self):
            rounding = bool()
            for x in self.freq:  
              if sum(x[0:-1]) > x[-1]:
                rounding =True
                print (self.QBcode)
            return rounding            
            
        def laduj_dane_text(self):
            if len(self.dimensions) == 0:
                pass
                #self.dane2 = ben_cc[(ben_cc["DIM_NAME"]==(self.dimensions[0]))]
                self.dane2 = self.o[["SMI_CODE", self.QBcode]]                               # self.o  nazwa obiektu
                self.dane3 = self.dane2.drop_duplicates(["SMI_CODE"])#.unique()
                self.dane5 = (self.dane3[self.QBcode].tolist())
                self.dane6.append(pd.Series(v for v in self.dane5))
            else:
                  i=0
                  for x in self.dimensions:
                    self.dane2 = self.o[(self.o["DIM_NAME"]==(x))]#& (ben_cc[self.QBcode]==self.QBcode)]
                    self.dane7 = self.dane2[self.QBcode].dropna()
                    self.dane3 = self.dane7.tolist()
                    if all(isinstance(n, float) for n in self.dane3):
                         self.dane50 = [ int(x) for x in self.dane3 ] 
                         self.dane5 = [ str(x) for x in self.dane50 ]    
                    else:
                         self.dane5 = [ str(x) for x in self.dane7 ] 
                         
                    self.dane51 = pd.Series((v for v in self.dane5))
                    
                    self.dane6.append(self.dane51)
              
                    i+=1
        def laduj_dane_num(self):
            if len(self.dimensions) == 0:
                pass
                #self.dane2 = ben_cc[(ben_cc["DIM_NAME"]==(self.dimensions[0]))]
                self.dane2 = self.o[["SMI_CODE", self.QBcode]]
                self.dane3 = self.dane2.drop_duplicates(["SMI_CODE"])#.unique()
                self.dane5 = (self.dane3[self.QBcode].tolist())
                self.dane6.append(pd.Series(v for v in self.dane5))
            else:
                  i=0
                  for x in self.dimensions:
                    self.dane2 = self.o[(self.o["DIM_NAME"]==(x))]#& (ben_cc[self.QBcode]==self.QBcode)]
                    self.dane7 = self.dane2[self.QBcode].dropna()
                    self.dane3 = self.dane7.tolist()
                    

                    self.dane51 = pd.Series((v for v in self.dane3))
                    
                    self.dane6.append(self.dane51)
                  
                    i+=1

        def laduj_dane_checkboxes(self):
            
            self.wyniki=[]
            self.wyniki1=[] 
            if len(self.dimensions) == 0:
                for x in self.referenceTablesplit: 
                    self.dane2 = self.o[["SMI_CODE", self.QBcode + "_" + x]]
                    self.dane3 = self.dane2.drop_duplicates(["SMI_CODE"])#.unique()
                    self.dane5 = (self.dane3[self.QBcode].tolist())
                    self.dane6.append(pd.Series(v for v in self.dane5))
            else:
                                  
                    for x in self.dimensions:
                        chbox = pd.DataFrame()
                        for num in self.referenceTablesplit: 
                            self.dane2 = self.o[(self.o["DIM_NAME"]==(x))]#& (ben_cc[self.QBcode]==self.QBcode)]
                            self.dane1 = self.dane2[self.QBcode + "_"+ num]  #.dropna()
                            chbox = pd.concat([chbox, self.dane1],  axis=1, join='outer')
                        self.dane6.append(chbox.fillna(0.0).astype(int))
                                           
                 
        def ilegrup(self):
            if pd.isnull(self.dimensionSetName):
                return "bez grup"
            else:
                
                self.grupy_pracownicze = dimensionssets.loc[dimensionssets.dimensionSet == self.dimensionSetName,'dimension'].tolist()
                self.ile_grup_prac =  len(self.grupy_pracownicze)
                for x in range(self.ile_grup_prac):
                    self.labele_grup.append(self.grupy_pracownicze[x].split(':')[-1])
                    self.dimensions.append(self.grupy_pracownicze[x].split(':')[0])
                #return int(self.zakres[0][5])
            #self.grupy_pracownicze=
            
        def liczenie(self): 
            i=0
            self.perc = [[] for y in range(len(self.dane4))]
            for x in self.dane4:
                x.dropna()
                self.perc[i].append(np.nanpercentile(x, 25)) 
                self.perc[i].append(np.nanpercentile(x, 50))
                self.perc[i].append(np.mean(x))
                self.perc[i].append(np.nanpercentile(x, 75))
                self.perc[i].append(x.count())
                i+=1
        def liczenie_RB(self): 
            a =0
            i=0
            self.perc = [[] for y in range(len(self.dane6))]
            self.wyniki1 = [[] for y in range(len(self.dane6))]
            self.wyniki1_raw = [[] for y in range(len(self.dane6))]
            self.wyniki1_round = [[] for y in range(len(self.dane6))]
            self.freq = [[] for y in range(len(self.dane6))]
            for i in range(len(self.dane6)):
                     #self.dane6[i].tolist()
                     for k in range(len(self.referenceTablesplit)):
                           self.freq[i].append(self.dane6[i].tolist().count(self.referenceTablesplit[k]))
                           
                           #self.freq[i].append(self.dane6[i].value_counts(normalize=True))
                           #self.freq[i].append(self.dane6[i].value_counts())
                           #self.freq[i].append(self.dane6[i].iloc[self.dane6[i].iloc[0] == self.referenceTablesplit[k], self.dane6[i].iloc[0]].count())
                           #i+=1 self.freq[i].append(self.dane6[i].iloc[self.dane6[i][0]== self.referenceTablesplit[k]].sum())
                     #self.freq[i].append(sum(self.freq[i]))  
                     self.freq[i].append(len(self.dane6[i].tolist()))
                     for l in range(len(self.referenceTablesplit)): 
                          if int(self.freq[i][-1]) < 3:
                              self.wyniki1[i].insert(a, "--")
                          else:
                              sumatmp = int(self.freq[i][l])/int(self.freq[i][-1])
                              self.wyniki1[i].insert(a, "{0:.0f}%".format(sumatmp*100))
                              self.wyniki1_raw[i].insert(a, sumatmp*100)
                              self.wyniki1_round[i].insert(a, round(sumatmp*100))
                        #i+=self.ile_grup
                              a+=1
                     #return self.wyniki     
        def liczenie_checkboxes(self): 
            a=0
            i=0
            self.perc = [[] for y in range(len(self.dane6))]
            self.wyniki1 = [[] for y in range(len(self.dane6))]
            self.wyniki1_raw = [[] for y in range(len(self.dane6))]
            self.wyniki1_round = [[] for y in range(len(self.dane6))]            
            self.freq = [[] for y in range(len(self.dane6))]
            for i in range(len(self.dane6)):
                     
                     for k in self.referenceTablesplit:
                           
                         self.freq[i].append(self.dane6[i][self.QBcode+"_"+k].astype(bool).sum(axis=0))

                     self.freq[i].append(self.dane6[i].astype(bool).sum(axis=1).sum(axis=0))
                     for l in range(len(self.referenceTablesplit)): 
                          if int(self.freq[i][-1]) < 3:
                              self.wyniki1[i].insert(a, "--")
                          else:
                              sumatmp = int(self.freq[i][l])/int(self.freq[i][-1])
                              self.wyniki1[i].insert(a, "{0:.0f}%".format(sumatmp*100))
                              self.wyniki1_raw[i].insert(a, sumatmp*100)
                              self.wyniki1_round[i].insert(a, round(sumatmp*100))
                            
        def procenty(self, value):
           tabelaprocenty=[]
           for i in value:
               if i <= 1:  
                  tabelaprocenty.append( "{0:.0f}%".format(sumatmp*100))    
           return tabelaprocenty                 
                           

        def licz(self):
            for x in range(len(self.referenceTablesplit)):
                self.freq = self.dane4[i][self.dane4[i] == self.referenceTablesplit[x]].count()
                self.wyniki.append(self.freq/self.count)

       
        def liczenie_text(self): 
            #print("text")
            #def laduj_dane_checkboxes(self):
            self.wyniki=[]
            self.wyniki1=[]
            pass
        
        
        
        def liczenie_percentile(self):                              # seocond arument for round could be precision of rounding
            i=0
            
            self.wyniki1 = [[] for y in range(len(self.dane6))]
            self.freq = [[self.dane6[y].size] for y in range(len(self.dane6))]         #  Freq is list in list to align with other Qs types
            for x in self.dane6:
                
                if self.Type_of_question == "percentage":
                    if self.freq[i][0] >4:
                    
                        self.wyniki1[i].append("{0:.0f}%".format(round(float(np.nanpercentile(x, 25))))) 
                        self.wyniki1[i].append("{0:.0f}%".format(round(float(np.nanpercentile(x, 50)))))
                        self.wyniki1[i].append("{0:.0f}%".format(round(np.mean(x))))
                        self.wyniki1[i].append("{0:.0f}%".format(round(float(np.nanpercentile(x, 75)))))
                        self.wyniki1[i].append(x.count())
                        i+=1
                    elif self.freq[i][0] == 4:    
                        self.wyniki1[i].append("--") 
                        self.wyniki1[i].append("{0:.0f}%".format(round(float(np.nanpercentile(x, 50)))))
                        self.wyniki1[i].append("{0:.0f}%".format(round(np.mean(x))))
                        self.wyniki1[i].append("--")
                        self.wyniki1[i].append(x.count())
                        i+=1
                    elif self.freq[i][0] == 3:
                        self.wyniki1[i].append("--") 
                        self.wyniki1[i].append("--")
                        self.wyniki1[i].append("{0:.0f}%".format(round(np.mean(x))))
                        self.wyniki1[i].append("--")
                        self.wyniki1[i].append(x.count())
                        i+=1
                    elif self.freq[i][0] < 3:
                        self.wyniki1[i].append("--") 
                        self.wyniki1[i].append("--")
                        self.wyniki1[i].append("--")
                        self.wyniki1[i].append("--")
                        self.wyniki1[i].append(x.count())    
                        i+=1
                else:    
                    if self.freq[i][0] >4:
                        self.wyniki1[i].append("{0:,.0f}".format(round(float(np.nanpercentile(x, 25))))) 
                        self.wyniki1[i].append("{0:,.0f}".format(round(float(np.nanpercentile(x, 50)))))
                        self.wyniki1[i].append("{0:,.0f}".format(round(np.mean(x))))
                        self.wyniki1[i].append("{0:,.0f}".format(round(float(np.nanpercentile(x, 75)))))
                        self.wyniki1[i].append(x.count())
                        i+=1
                    elif self.freq[i][0] == 4:    
                        self.wyniki1[i].append("--") 
                        self.wyniki1[i].append("{0:,.0f}".format(round(float(np.nanpercentile(x, 50)))))
                        self.wyniki1[i].append("{0:,.0f}".format(round(np.mean(x))))
                        self.wyniki1[i].append("--")
                        self.wyniki1[i].append(x.count())
                        i+=1
                    elif self.freq[i][0] == 3:
                        self.wyniki1[i].append("--") 
                        self.wyniki1[i].append("--")
                        self.wyniki1[i].append("{0:,.0f}".format(round(np.mean(x))))
                        self.wyniki1[i].append("--")
                        self.wyniki1[i].append(x.count())
                        i+=1
                    elif self.freq[i][0] < 3:
                        self.wyniki1[i].append("--") 
                        self.wyniki1[i].append("--")
                        self.wyniki1[i].append("--")
                        self.wyniki1[i].append("--")
                        self.wyniki1[i].append(x.count())    
                        i+=1               
                    
                    
            


        
pytanie = a["Question code"].tolist()
#pytaniecc = c["Question code"].tolist()
ile_tab = a["Question code"].tolist()
pytania = a.loc[a['Type'].str.contains("Question")]
#pytaniacc = c.loc[c['Type'].str.contains("Question")]


for x in range(len(pytania)):
    
  #  if not in 
    
    tab.append(question(pytania.iloc[x, 2],pytania.iloc[x, 10],pytania.iloc[x, 4],pytania.iloc[x, 7],pytania.iloc[x, 8],pytania.iloc[x, 9]))


print("--- %s seconds ---" % (time.time() - start_time))



