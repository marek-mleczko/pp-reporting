# -*- coding: utf-8 -*-
"""
Created on Fri Mar 29 17:05:55 2019

@author: marek-mleczko
"""
import pandas as pd
import json

question_bank = pd.read_excel('mdc/QB.xlsm', sheet_name = "Company Cars")
a = question_bank.loc[question_bank['Countries'].str.contains("SI")]
pytania = a.loc[a['Type'].str.contains("Question")]
#x = "sample-submissions.json"

jason_file = open("sample-submissions.json", "r")

connector = json.load(jason_file)

#print(connector["CAR1"])
#print(type(connector))
#for value in connector:
#    print("Key:")
#    print(value)
#for v, k in connector.items():
#    print(v)
#    print(k)    
    
def myprint(d):
  for k, v in d.items():
    if isinstance(v, dict):
      myprint(v)
    else:
      print("{0} : {1}".format(k, v))  
data = connector["data"]
submissions = data["submissions"]
lista=[]
for id in submissions:
      lista.append(id["answers"])
      
my_data=lista[2]

      