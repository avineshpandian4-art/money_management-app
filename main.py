import pandas as pd
import numpy as np
from datetime import datetime,date
import time
import openpyxl,os
import streamlit as st
import plotly.express as px

data={
       "BUDGET":[],
       "DATE":[],
       "PETROL":[],
       "SNACKS":[],
       "BREAKFAST":[],
       "LUNCH":[],
       "DINNER":[],
       "SHOPPING":[],
       "BUS":[],
       "OTHERS":[],
       "SAVINGS":[]
       }

file_path="output.xlsx"
sheet_n="User"




def update_excel(update,file_path,sheet_n):
 workbook=openpyxl.load_workbook(file_path)
 sheet=workbook[sheet_n]
 new_row=update
 sheet.append(new_row)
 workbook.save(file_path)



def create_excel(df,file_path,sheet):
 df.to_excel(file_path,sheet_name=sheet_n,index=False)

def check_excel():
   if not os.path.exists(file_path):
        df=pd.DataFrame(data)
        create_excel(df,file_path,sheet_n)

check_excel()

def clean_excel(path,sheet_name):
    try:
      data=pd.read_excel(path)
      data.drop_duplicates(inplace=True)
      with pd.ExcelWriter(file_path, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
         data.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as e:
      return 
   
def get_input():
 budget=st.number_input("Enter todays budget:")
 dat=date.today()
 petrol=st.number_input("Enter amout spent on petrol:")
 snacks=st.number_input("Enter amount spent on snacks:")
 breakfast=st.number_input("Enter amount spent on breakfast:")
 lunch=st.number_input("Enter amount spent on lunch:")
 dinner=st.number_input("Enter amount spent on dinner:")
 shopping=st.number_input("Enter amount spent on shopping:")
 bus=st.number_input("Enter amount spent on bus:")
 others=st.number_input("Enter amount spent on any others:")

 total=petrol+snacks+breakfast+lunch+dinner+shopping+bus+others
 savings=max(0,budget-total)
 return [budget,dat,petrol,snacks,breakfast,lunch,dinner,shopping,bus,others,savings]

def dele_last_entry(file_name,sheet_n):

   wb = openpyxl.load_workbook(file_name)
   sheet = wb[sheet_n]
   if sheet.max_row > 1:
      sheet.delete_rows(sheet.max_row)
      wb.save(file_name)
   
def show(tab):
   st.dataframe(tab)
   
   
def main():
   st.title("MONEY MANAGEMENT 🐂")
   update=get_input()
   #clean_excel(file_path,sheet_n)
   tab=pd.read_excel(file_path)
          
   booln=st.button("submit")
   if booln:
      update_excel(update,file_path,sheet_n)
      tab=pd.read_excel(file_path)
      total_saving=tab["SAVINGS"].sum()
   clean_excel(file_path,sheet_n)
   show(tab)
          
          
   #if st.button("DELETE ENTRY"):
    #  dele_last_entry(file_path,sheet_n)
   
   try:
      total_saving=tab["SAVINGS"].sum()
      delt=tab["SAVINGS"].iloc[-1]
   except Exception as e:
      total_saving,delt=0,0
   with st.container(horizontal=True):
      if st.button("DELETE ENTRY"):dele_last_entry(file_path,sheet_n)
      st.space("xxlarge")
      st.space("xxlarge")
      st.space("xxlarge")
      st.space("small")
      st.metric(label="Total Saving", value=f"₹{total_saving} ",delta=f"+{delt}")
   try:
      df=pd.DataFrame(
         {
         "Category":["PETROL","SNACKS","BREAKFAST","LUNCH","DINNER","SHOPPING","BUS","OTHERS"],
         "Values":[
            tab["PETROL"].mean(),
            tab["SNACKS"].mean(),
            tab["BREAKFAST"].mean(),
            tab["LUNCH"].mean(),
            tab["DINNER"].mean(),
            tab["SHOPPING"].mean(),
            tab["BUS"].mean(),
            tab["OTHERS"].mean()
         ]}
   )
      #df["Values"]=tab[df["Category"]].mean()
      fig = px.pie(
         df,
         values="Values",
         names="Category",
         title="Distribution of Readings"
        )

      st.plotly_chart(fig, use_container_width=True)

   except Exception as e:
      st.write(e)
   
   
   


if __name__=="__main__":
   main()
   
   

       
       
       
