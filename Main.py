import os
import sys
import tkinter
import ipdb  # отладчик

from collections import *   # отсюда берем класс OrderedDict
import sqlite3   

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.Qt import *

import OpenExelFrm   # импортируем главную форму
import SaveBDFrm     # форма сохранения базы данных
import EditeDBFrm    # форма работы с БД 
import threading     # многопоточность
import time          # для паузы

import pyexcel    # для работы с Excel
from pyexcel import *

global window
global excel_book
excel_book = [] # тут будет храниться книга Excel
global keys
keys = []  # массив для названия страниц
global Flag

########################################################################################################################################################
#######################################################################################################################################################
                                                             ### Окно(модуль) Редактирования БД ###

class DBEditeApp(QtWidgets.QDialog, EditeDBFrm.Ui_DBEditeForm):
   def __init__(self):
      super().__init__()
      self.setupUi(self)


##################################                      Конец                     ####################################################################
##################################            Окна (модуля) Редактирования БД     ####################################################################
#######################################################################################################################################################
#######################################################################################################################################################

#######################################################################################################################################################
#######################################################################################################################################################
                                  ####### Форма (модуль) Сохранения БД   ############

class DialogApp(QtWidgets.QDialog, SaveBDFrm.Ui_DBDialog):  # Форма сохранения БД
   def __init__(self):
     super().__init__()
     self.setupUi(self)

     self.DBFileBtn.clicked.connect(self.SelDBFileFunc) # выбор файла для сохранения
     self.accepted.connect(self.DBSaveFunc) # если нажато ОК, то сохраняем
     self.DBRb.clicked.connect(self.ChangeInputFunc) # изменение параметров ввода
     self.ExcelRb.clicked.connect(self.ChangeInputFunc) # изменение параметров ввода

###########################  изменение параметров ввода ############################3
   def ChangeInputFunc(self):
     if self.DBRb.isChecked() :
       self.RowRSSB.setValue(2)
       self.ColRSSB.setValue(1)
       self.KeyRSSB.setValue(1)
       self.ColRowRSSB.setValue(len(excel_book[keys[self.KeyRSSB.value()-1]])-1)
     
       self.RowReqSB.setValue(2)
       self.ColReqSB.setValue(1)
       self.KeyReqSB.setValue(2)
       self.ColRowReqSB.setValue(len(excel_book[keys[self.KeyReqSB.value()-1]])-1)
     
       self.RowInvSB.setValue(2)
       self.ColInvSB.setValue(1)
       self.KeyInvSB.setValue(3)
       self.ColRowInvSB.setValue(len(excel_book[keys[self.KeyInvSB.value()-1]])-1)
      
       self.RowZIPSB.setValue(2)
       self.ColZIPSB.setValue(1)
       self.KeyZIPSB.setValue(4)
       self.ColRowZIPSB.setValue(len(excel_book[keys[self.KeyZIPSB.value()-1]])-1)
      
       self.RowDepsSB.setValue(2)
       self.ColDepsSB.setValue(1)
       self.KeyDepsSB.setValue(5)
       self.ColRowDepsSB.setValue(len(excel_book[keys[self.KeyDepsSB.value()-1]])-1)
      
       self.RowCatsSB.setValue(2)
       self.ColCatsSB.setValue(1)
       self.KeyCatsSB.setValue(6)
       self.ColRowCatsSB.setValue(len(excel_book[keys[self.KeyCatsSB.value()-1]])-1)

       self.TabStatusCB.setChecked(False)
     else:
       self.TabStatusCB.setChecked(True)

       self.RowRSSB.setValue(5)
       self.ColRSSB.setValue(2)
       self.KeyRSSB.setValue(1)
       self.ColRowRSSB.setValue(35427)

       self.RowReqSB.setValue(5)
       self.ColReqSB.setValue(6)
       self.KeyReqSB.setValue(1)
       self.ColRowReqSB.setValue(35427)

       self.RowInvSB.setValue(5)
       self.ColInvSB.setValue(8)
       self.KeyInvSB.setValue(1)
       self.ColRowInvSB.setValue(35427)

       self.RowZIPSB.setValue(4)
       self.ColZIPSB.setValue(1)
       self.KeyZIPSB.setValue(3)
       self.ColRowZIPSB.setValue(567)

       self.RowDepsSB.setValue(3)
       self.ColDepsSB.setValue(3)
       self.KeyDepsSB.setValue(3)
       self.ColRowDepsSB.setValue(29)

       self.RowCatsSB.setValue(606)
       self.ColCatsSB.setValue(1)
       self.KeyCatsSB.setValue(3)
       self.ColRowCatsSB.setValue(31)

################################## выбор файла для сохранения #####################################
   def SelDBFileFunc(self): # выбираем файл БД
     filename, ok = QFileDialog.getSaveFileName(self, 
                             "Сохранить файл",
                             ".",
                             "SQLite Files(*.db)") 
     self.DBFileNameLb.setText(filename)

#################################### непосредственно сохранение таблицы в БД ###########################################
   def DBSaveFunc(self): # сохраняем таблицу в БД
       global window
       if self.DBFileNameLb.text() != "" :
          global excel_book
          global keys
          filename = self.DBFileNameLb.text()
          conn = sqlite3.connect(filename)  # подключаем БД
          cursor = conn.cursor()
          window.AllPB.show()

# начинаем заполнять таблицы БД
       #таблица RS (р/с)

          if self.RSChB.isChecked() :
           try :
             cursor.execute("""CREATE TABLE IF NOT EXISTS RS (  
               Type TEXT,
               Name TEXT,
               Ser_Num TEXT,
               Man_Date DATE);
             """)  
             cursor.execute("DELETE FROM RS")  # Удаляем старые записи
           except:
             QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка обращения к таблице РАДИОСТАНЦИИ. Данные могут быть не сохранены!", QMessageBox.Ok).exec()                 
           conn.commit()  # подтверждаем удаление
           
           #проверяем правильность первоночальных условий и пишем данные
           if  len(keys) >= self.KeyRSSB.value() and len(excel_book[keys[self.KeyRSSB.value()-1]]) >= self.RowRSSB.value()+self.ColRowRSSB.value()-1 and len(excel_book[keys[self.KeyRSSB.value()-1]][self.RowRSSB.value()-1]) >= self.ColRSSB.value() + 3 :
            RD = 0
            WD = 0
            DUB = 0
            window.AllPB.setValue(0)
            window.AllPB.setFormat("%v/"+str(self.RowRSSB.value()+self.ColRowRSSB.value()-1))
            window.AllPB.setMaximum(self.RowRSSB.value()+self.ColRowRSSB.value()-1)
            for i in range (self.RowRSSB.value()-1, self.RowRSSB.value()+self.ColRowRSSB.value()-1) :
              row = []
              for j in range (self.ColRSSB.value()-1, self.ColRSSB.value() + 3) : 
                try: 
                  row.append(excel_book[keys[self.KeyRSSB.value()-1]][i][j])
                except:
                  RD += 1
                  #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка чтения данных из таблицы! Проверьте правильность заданных параметров!!!", QMessageBox.Ok).exec()                 
              try:
                #эта пороверка сильно замедляет, но необходима... возможность добавления без проверок можно добавить в форму "Работа с БД"
                if self.DoublChB.isChecked() :
                 cursor.execute("SELECT * FROM RS WHERE Type = ? AND Name = ? AND Ser_Num = ? AND Man_Date = ? LIMIT 1", row)
                 if len(cursor.fetchall()) == 0 :
                  cursor.execute("INSERT INTO RS values (?,?,?,?);", row)
                 else:
                  DUB += 1
                else:
                 cursor.execute("INSERT INTO RS values (?,?,?,?);", row)
              except:
                WD += 1
                #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка записи строки в базу данных! Запись в БД (РАДИОСТАНЦИИ) добавлена не будет", QMessageBox.Ok).exec()                 
              window.AllPB.setValue(i)
            if RD != 0 or WD != 0 or DUB != 0 :
             QMessageBox(QMessageBox.Warning, "Внимание!!!" , "При формировании таблицы РАДИОСТАНЦИИ зафиксированы ошибки! Ошибок чтения "+str(RD)+"! Ошибок Записи "+str(WD)+"! Удалено дублей "+str(DUB)+"!", QMessageBox.Ok).exec()                 
            conn.commit()     
           else :
            QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Неправильно заданы начальные условия!!! Записи в БД (РАДИОСТАНЦИИ) добавлены не будут", QMessageBox.Ok).exec()                 

        #таблица Req (заявки)
          if self.ReqChB.isChecked() :
           try:       
             cursor.execute("""CREATE TABLE IF NOT EXISTS Req (
               Dep      TEXT,
               Req_date DATE,
               Ser_Num  TEXT,
               Type     TEXT,
               Name     TEXT
               );
             """)
             cursor.execute("DELETE FROM Req")  # Удаляем старые записи
           except:
             QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка обращения к таблице ЗАЯВКИ. Данные могут быть не сохранены!", QMessageBox.Ok).exec()                 
           conn.commit()  # подтверждаем удаление

           #проверяем правильность первоночальных условий и пишем данные
           if  len(keys) >= self.KeyReqSB.value() and len(excel_book[keys[self.KeyReqSB.value()-1]]) >= self.RowReqSB.value()+self.ColRowReqSB.value()-1 and len(excel_book[keys[self.KeyReqSB.value()-1]][self.RowReqSB.value()-1]) >= self.ColReqSB.value() + 1 :
            RD = 0
            WD = 0                      
            window.AllPB.setValue(0)
            window.AllPB.setFormat("%v/"+str(self.RowReqSB.value()+self.ColRowReqSB.value()-1))
            window.AllPB.setMaximum(self.RowReqSB.value()+self.ColRowReqSB.value()-1)
            ZeroLine = 0
            if self.TabStatusCB.isChecked() :
             sv = 1
            else :
             sv = 4
            for i in range (self.RowReqSB.value()-1, self.RowReqSB.value()+self.ColRowReqSB.value()-1) : 
              row = []
              for j in range (self.ColReqSB.value()-1,self.ColReqSB.value() + sv) :
                  try:
                    row.append(excel_book[keys[self.KeyReqSB.value()-1]][i][j])
                  except:
                    #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка чтения данных из таблицы! Проверьте правильность заданных параметров!!!", QMessageBox.Ok).exec()                 
                    RD += 1
              try:
                if self.TabStatusCB.isChecked() :
                 row.append(excel_book[keys[self.KeyRSSB.value()-1]][i][self.ColRSSB.value()-1+2])  
                 row.append(excel_book[keys[self.KeyRSSB.value()-1]][i][self.ColRSSB.value()-1])  
                 row.append(excel_book[keys[self.KeyRSSB.value()-1]][i][self.ColRSSB.value()-1+1])  

              except:
                #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка чтения данных из таблицы! Проверьте правильность заданных параметров!!!", QMessageBox.Ok).exec()                 
                RD += 1
              try:
                if row[0] != "" and row[1] != "" :
                  cursor.execute("INSERT INTO Req values (?,?,?,?,?);", row)
                else :
                  ZeroLine += 1
              except:
                #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка записи строки в базу данных! Запись в БД ЗАЯВКИ) добавлена не будет", QMessageBox.Ok).exec()                 
                WD += 1
              window.AllPB.setValue(i)
             #print(row)
            conn.commit()
            if RD != 0 or WD != 0 or ZeroLine != 0 :
             QMessageBox(QMessageBox.Warning, "Внимание!!!" , "При формировании таблицы ЗАЯВКИ зафиксированы ошибки! Ошибок чтения "+str(RD)+"! Ошибок Записи "+str(WD)+"! Пустых строк "+str(ZeroLine)+"!", QMessageBox.Ok).exec()                 

           else :
            QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Неправильно заданы начальные условия!!! Записи в БД (ЗАЯВКИ) добавлены не будут", QMessageBox.Ok).exec()                 
     
        # таблица Inv (счета)
          if self.InvChB.isChecked() :
           try:
             cursor.execute("""CREATE TABLE IF NOT EXISTS Inv (
               Inv_Num  TEXT,
               Inv_date DATE,
               Sum      REAL,
               Ser_Num  TEXT
               );
             """)
             cursor.execute("DELETE FROM Inv")  # Удаляем старые записи
           except:
             QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка обращения к таблице СЧЕТА. Данные могут быть не сохранены!", QMessageBox.Ok).exec()                 
           conn.commit()  # подтверждаем удаление

           #проверяем правильность первоночальных условий и пишем данные
           if  len(keys) >= self.KeyInvSB.value() and len(excel_book[keys[self.KeyInvSB.value()-1]]) >= self.RowInvSB.value()+self.ColRowInvSB.value()-1 and len(excel_book[keys[self.KeyInvSB.value()-1]][self.RowInvSB.value()-1]) >= self.ColInvSB.value() + 2 :
            RD = 0
            WD = 0
            ZeroLine = 0
            window.AllPB.setValue(0)
            window.AllPB.setFormat("%v/"+str(self.RowInvSB.value()+self.ColRowInvSB.value()-1))
            window.AllPB.setMaximum(self.RowInvSB.value()+self.ColRowInvSB.value()-1)

            if self.TabStatusCB.isChecked() :
              sv = 2
            else :
              sv = 5
            for i in range (self.RowInvSB.value()-1, self.RowInvSB.value()+self.ColRowInvSB.value()-1) : 
              row = []
              for j in range (self.ColInvSB.value()-1,self.ColInvSB.value() + sv) : 
                  try:
                    row.append(excel_book[keys[self.KeyInvSB.value()-1]][i][j])
                  except:
                    #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка чтения данных из таблицы! Проверьте правильность заданных параметров!!!", QMessageBox.Ok).exec()                 
                    RD += 1
              try:
                if self.TabStatusCB.isChecked() :
                 row.append(excel_book[keys[self.KeyRSSB.value()-1]][i][self.ColRSSB.value()-1+2])  
                 row.append(excel_book[keys[self.KeyRSSB.value()-1]][i][self.ColRSSB.value()-1])  
                 row.append(excel_book[keys[self.KeyRSSB.value()-1]][i][self.ColRSSB.value()-1+1])  

              except:
                #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка чтения данных из таблицы! Проверьте правильность заданных параметров!!!", QMessageBox.Ok).exec()                 
                RD += 1
              try:             
                if row[0] != "" and row[1] != "" and row[2] != 0 :
                  cursor.execute("INSERT INTO Inv values (?,?,?,?,?,?);", row)
                else :
                  ZeroLine += 1
              except:
                #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка записи строки в базу данных! Запись в БД СЧЕТА) добавлена не будет", QMessageBox.Ok).exec()                 
                WD +=1
              window.AllPB.setValue(i)
              #print(row)
            conn.commit()
            if RD != 0 or WD != 0 or ZeroLine != 0 :
             QMessageBox(QMessageBox.Warning, "Внимание!!!" , "При формировании таблицы СЧЕТА зафиксированы ошибки! Ошибок чтения "+str(RD)+"! Ошибок Записи "+str(WD)+"! Пустых строк "+str(ZeroLine)+"!", QMessageBox.Ok).exec()                 

           else :
            QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Неправильно заданы начальные условия!!! Записи в БД (СЧЕТА) добавлены не будут", QMessageBox.Ok).exec()                 

        # таблица ZIP (ЗИП - запчасти)
          if self.ZIPChB.isChecked() : 
           try:
             cursor.execute("""CREATE TABLE IF NOT EXISTS ZIP (
               Art   TEXT,
               Price REAL
               );
             """)
             cursor.execute("DELETE FROM ZIP")  # Удаляем старые записи
           except:
             QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка обращения к таблице БАЗА ЗИП. Данные могут быть не сохранены!", QMessageBox.Ok).exec()                 
           conn.commit()  # подтверждаем удаление

           #проверяем правильность первоночальных условий и пишем данные
           if  len(keys) >= self.KeyZIPSB.value() and len(excel_book[keys[self.KeyZIPSB.value()-1]]) >= self.RowZIPSB.value()+self.ColZIPSB.value()-1 and len(excel_book[keys[self.KeyZIPSB.value()-1]][self.RowInvSB.value()-1]) >= self.ColZIPSB.value() + 1 :
            RD = 0
            WD = 0
            window.AllPB.setValue(0)
            window.AllPB.setFormat("%v/"+str(self.RowZIPSB.value()+self.ColRowZIPSB.value()-1))
            window.AllPB.setMaximum(self.RowZIPSB.value()+self.ColRowZIPSB.value()-1)

            for i in range (self.RowZIPSB.value()-1, self.RowZIPSB.value()+self.ColRowZIPSB.value()-1) : 
              row = []
              for j in range (self.ColZIPSB.value()-1,self.ColZIPSB.value() + 1) : 
                  try:
                    row.append(excel_book[keys[self.KeyZIPSB.value()-1]][i][j])
                  except:
                    #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка чтения данных из таблицы! Проверьте правильность заданных параметров!!!", QMessageBox.Ok).exec()                 
                    RD += 1
              #row.append(excel_book[keys[self.KeyRSSB.value()-1]][i][self.ColRSSB.value()-1+2])  
              try:
                cursor.execute("INSERT INTO ZIP values (?,?);", row)
              except:
                #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка записи строки в базу данных! Запись в БД (БАЗА ЗИП) добавлена не будет", QMessageBox.Ok).exec()                 
                WD += 1
              window.AllPB.setValue(i)
              #print(row)
            conn.commit()
            if RD != 0 or WD != 0 :
             QMessageBox(QMessageBox.Warning, "Внимание!!!" , "При формировании таблицы БАЗА ЗИП зафиксированы ошибки! Ошибок чтения "+str(RD)+"! Ошибок Записи "+str(WD)+"!", QMessageBox.Ok).exec()                 

           else :
            QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Неправильно заданы начальные условия!!! Записи в БД (БАЗА ЗИП) добавлены не будут", QMessageBox.Ok).exec()                 

        # таблица Deps (Округа/подразделения)
          if self.DepsChB.isChecked() :
           try:
             cursor.execute("""CREATE TABLE IF NOT EXISTS Deps (
               Dep TEXT
               );
             """)
             cursor.execute("DELETE FROM Deps")  # Удаляем старые записи
           except:
             QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка обращения к таблице ОКРУГА. Данные могут быть не сохранены!", QMessageBox.Ok).exec()                 
           conn.commit()  # подтверждаем удаление

           #проверяем правильность первоночальных условий и пишем данные
           if  len(keys) >= self.KeyDepsSB.value() and len(excel_book[keys[self.KeyDepsSB.value()-1]]) >= self.RowDepsSB.value()+self.ColDepsSB.value()-1 and len(excel_book[keys[self.KeyDepsSB.value()-1]][self.RowDepsSB.value()-1]) >= self.ColDepsSB.value() :
            RD = 0
            WD = 0
            window.AllPB.setValue(0)
            window.AllPB.setFormat("%v/"+str(self.RowDepsSB.value()+self.ColRowDepsSB.value()-1))
            window.AllPB.setMaximum(self.RowDepsSB.value()+self.ColRowDepsSB.value()-1)

            for i in range (self.RowDepsSB.value()-1, self.RowDepsSB.value()+self.ColRowDepsSB.value()-1) : 
              row = []
              #for j in range (self.ColDepsSB.value()-1,self.ColDepsSB.value() + 0) : 
              #    row.append(excel_book[keys[self.KeyDepsSB.value()-1]][i][j])
              #row.append(excel_book[keys[self.KeyRSSB.value()-1]][i][self.ColRSSB.value()-1+2])  

              try:        
                row.append(excel_book[keys[self.KeyDepsSB.value()-1]][i][self.ColDepsSB.value()-1])
              except:
                #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка чтения данных из таблицы! Проверьте правильность заданных параметров!!!", QMessageBox.Ok).exec()                 
                RD += 1
              try:
                cursor.execute("INSERT INTO Deps values (?);", row ) #row
              except:
                #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка записи строки в базу данных! Запись в БД (ОКРУГА) добавлена не будет", QMessageBox.Ok).exec()                 
                WD += 1
              window.AllPB.setValue(i)
              #print(row)
            conn.commit()
            if RD != 0 or WD != 0 :
             QMessageBox(QMessageBox.Warning, "Внимание!!!" , "При формировании таблицы ОКРУГА зафиксированы ошибки! Ошибок чтения "+str(RD)+"! Ошибок Записи "+str(WD)+"!", QMessageBox.Ok).exec()                 

           else :
            QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Неправильно заданы начальные условия!!! Записи в БД (Округа) добавлены не будут", QMessageBox.Ok).exec()                 

        # таблица Cats (Категории работ)
          if self.CatsChB.isChecked() :
           try:
             cursor.execute("""CREATE TABLE IF NOT EXISTS Cats (
               Name TEXT,
               Cat1 REAL,
               Cat2 REAL,
               Cat3 REAL
               );
             """)
             cursor.execute("DELETE FROM Cats")  # Удаляем старые записи
           except:
             QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка обращения к таблице КАТЕГОРИИ РАБОТ. Данные могут быть не сохранены!", QMessageBox.Ok).exec()                 
           conn.commit()  # подтверждаем удаление

           #проверяем правильность первоночальных условий и пишем данные
           if  len(keys) >= self.KeyCatsSB.value() and len(excel_book[keys[self.KeyCatsSB.value()-1]]) >= self.RowCatsSB.value()+self.ColCatsSB.value()-1 and len(excel_book[keys[self.KeyCatsSB.value()-1]][self.RowCatsSB.value()-1]) >= self.ColCatsSB.value() + 3 :
            WD = 0
            RD = 0
            window.AllPB.setValue(0)
            window.AllPB.setFormat("%v/"+str(self.RowCatsSB.value()+self.ColRowCatsSB.value()-1))
            window.AllPB.setMaximum(self.RowCatsSB.value()+self.ColRowCatsSB.value()-1)

            for i in range (self.RowCatsSB.value()-1, self.RowCatsSB.value()+self.ColRowCatsSB.value()-1) : 
              row = []
              for j in range (self.ColCatsSB.value()-1,self.ColCatsSB.value() + 3) : 
                  try:
                    row.append(excel_book[keys[self.KeyCatsSB.value()-1]][i][j])
                  except:
                    #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка чтения данных из таблицы! Проверьте правильность заданных параметров!!!", QMessageBox.Ok).exec()                 
                    RD += 1

              #row.append(excel_book[keys[self.KeyRSSB.value()-1]][i][self.ColRSSB.value()-1+2])  
              try:
                cursor.execute("INSERT INTO Cats values (?,?,?,?);", row)
              except:
                #QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Ошибка записи строки в базу данных! Запись в БД (КАТЕГОРИИ РАБОТ) добавлена не будет", QMessageBox.Ok).exec()                 
                WD += 1
              window.AllPB.setValue(i)
              #print(row)

            conn.commit() #подтверждаем изменения
            if RD != 0 or WD != 0 :
             QMessageBox(QMessageBox.Warning, "Внимание!!!" , "При формировании таблицы КАТЕГОРИИ РАБОТ зафиксированы ошибки! Ошибок чтения "+str(RD)+"! Ошибок Записи "+str(WD)+"!", QMessageBox.Ok).exec()                 

           else :
            QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Неправильно заданы начальные условия!!! Записи в БД (КАТЕГОРИИ РАБОТ) добавлены не будут", QMessageBox.Ok).exec()                 

           conn.close()  #и закрываем БД
       else :
          QMessageBox(QMessageBox.Warning, "Внимание" , "Имя файла не указано! Операция выполнена не будет!", QMessageBox.Ok).exec()                 
       window.AllPB.setValue(0)
       window.AllPB.setFormat("%v/0")
       window.AllPB.setMaximum(100)
       window.AllPB.hide()



#######################################################################################################################################################
#######################################################################################################################################################
###################################       Главное окно(модуь)        ##################################################################################


class TableApp(QtWidgets.QMainWindow, OpenExelFrm.Ui_MainWindow):

   def __init__(self):    #инициализация главного окна
     super().__init__()
     self.setupUi(self)

     self.OpenBtn.clicked.connect(self.OpenFileFunc)
     self.SaveBtn.clicked.connect(self.SaveFileFunc)
     self.sheetSpin.valueChanged.connect(self.ChangeSheetFunc)
     self.SaveAllBtn.clicked.connect(self.SaveAllFunc)
     self.DBbtn.clicked.connect(self.DBSelFunc)
     self.LoadDBBtn.clicked.connect(self.DBLoadFunc)
     self.EditeDBBt.clicked.connect(self.EditeDBFrmOpenFunc)
   

##########################   Поток для параллельной загрузки данных из excel #############################
   class OpenExcelThread (threading.Thread):
      def __init__(self, name, counter):
         threading.Thread.__init__(self)
         self.threadID = counter
         self.name = name
         self.counter = counter

      def run(self):
         global Flag
         global excel_book
         Flag = 1
         excel_book = pyexcel.get_book_dict(file_name=window.SelFileName.text()) # сохраняем всю книгу в коллекцию
         Flag = 0

##########################   Поток для параллельной выгрузки данных в excel #############################
   class SaveExcelThread (threading.Thread):
      Inpfilename=""
      def __init__(self, name, counter, filename):
         threading.Thread.__init__(self)
         self.threadID = counter
         self.name = name
         self.counter = counter
         self.Inpfilename = filename

      def run(self):
         global Flag
         global excel_book
         Flag = 1
         pyexcel.save_book_as(bookdict=excel_book, dest_file_name=self.Inpfilename)
         Flag = 0

################################# прокрутка прогресс-бара ####################################
   def RollPB(self):
         global Flag
         self.AllPB.show() #setViseble(True)
         self.AllPB.setFormat("")
         self.AllPB.setValue(0)
         while Flag != 0 :
          i = 0
          while Flag != 0 and i < 100:
           i += 1
           self.AllPB.setValue(i)
           time.sleep(0.01)
          #i=99
          self.AllPB.setInvertedAppearance(True)
          while Flag != 0 and i > 0:
            i -= 1
            self.AllPB.setValue(i)
            time.sleep(0.01)
          self.AllPB.setInvertedAppearance(False)
         self.AllPB.setValue(0)
         self.AllPB.hide() #setVisible(False)

###############   Открытие окна редактирования БД ##############################
   def EditeDBFrmOpenFunc(self):
     self.EDBFrm=DBEditeApp()
     self.EDBFrm.show()

###########################  Загрузка данных в таблицу из БД #######################################
   def DBLoadFunc(self):  #загрузка данных из БД
    global excel_book
    global keys
#    my_dict = [] 

    filename, filetype = QFileDialog.getOpenFileName(self,
                             "Выбрать файл",
                             ".",
                             "SQLite Files(*.db);;\
                             ") # диалог открытия файла
    if filename != "" :

      keys=['Радиостанции','Заявки','Счета','База ЗИП','Округа','Категории работ'] 

      #открываем БД
      conn = sqlite3.connect(filename)
      cursor = conn.cursor()
    
      self.AllPB.show()

      #Заполняем excel_book из БД ключи-имена таблиц
      cursor.execute("SELECT * FROM RS")
      result = cursor.fetchall()
  #    my_dict.append({keys[0]:result})
      excel_book = OrderedDict() #создаем OrderedDict 
      result.insert(0, ["Тип","Наименование","Серийный номер","Дата изготовления"])
      excel_book[keys[0]]= result  # добавляем и заполняем первый ключ
      self.AllPB.setValue(16)
      

      self.FullingTable(excel_book[keys[0]])  # выводим резултат в таблицу

      cursor.execute("SELECT * FROM Req")
      result = cursor.fetchall()
      result.insert(0,["Подразделение","Дата заявки","Серийный номер","Тип","Наименование"])
  #    my_dict.append({keys[1]:result})
      excel_book[keys[1]]= result #my_dict[1] # добавляем и заполняем второй ключ
      self.AllPB.setValue(32)

  # и так далее для всех ключей
      cursor.execute("SELECT * FROM Inv")
      result = cursor.fetchall()
      result.insert(0,["Номер счета","Дата счета","Стоимость","Серийный номер","Тип","Наименование"])
  #    my_dict.append({keys[2]:result})
      excel_book[keys[2]]= result
      self.AllPB.setValue(49)

      cursor.execute("SELECT * FROM ZIP")
      result = cursor.fetchall()
      result.insert(0,["Артикул","Цена"])
  #    my_dict.append({keys[3]:result})
      excel_book[keys[3]]= result
      self.AllPB.setValue(65)

      cursor.execute("SELECT * FROM Deps")
      result = cursor.fetchall()
      result.insert(0,["Округ"])
  #    my_dict.append({keys[4]:result})
      excel_book[keys[4]]= result
      self.AllPB.setValue(81)

      cursor.execute("SELECT * FROM Cats")
      result = cursor.fetchall()
      result.insert(0,["Устройство","Кат. 1","Кат. 2","Кат. 3"])
  #    my_dict.append({keys[5]:result})
      excel_book[keys[5]]= result
      self.AllPB.setValue(100)

      conn.close   # закрываем БД
      self.sheetSpin.setEnabled(True)  # активируем выбор страниц
      self.sheetSpin.setMaximum(len(keys)) # устанавливаем количество страниц
      self.AllPB.setValue(0)
      self.AllPB.hide()

################################ Открытие окна сохранения БД  ######################################
   def DBSelFunc(self): #открываем окно сохранения БД
      if self.tableWd.rowCount() != 0 and self.tableWd.columnCount() != 0 :
        self.DBwindow = DialogApp()
        self.DBwindow.show()
      else:
        QMessageBox(QMessageBox.Warning, "Внимание!!!" , "Таблица не заполнена!!!", QMessageBox.Ok).exec()
        

########################################### Изменение выбранной страницы #######################################    
   def ChangeSheetFunc(self): # изменение отображаемой страницы в таблице
      global excel_book
      global keys
      if self.sheetSpin.isEnabled() :
       my_array = excel_book[keys[self.sheetSpin.value() - 1]]    # записываем выбранную страницу в массив
# выводим массив в таблицу 
       self.FullingTable(my_array)
   
###################################  Функция заполнения таблицы  #############################################   
   def FullingTable(self, input_array) :  # записываем входящий массив в таблицу
      global keys
      self.SheetLb.setText(str(keys[self.sheetSpin.value() - 1]))
      self.tableWd.setRowCount(0)
      self.tableWd.setColumnCount(0) # обнуляем таблицу
      r = 0
      l = 0
      for line in input_array:  # записываем массив в таблицу - берем строку
       self.tableWd.insertRow(l) # добавляем в таблицу строку
       for item in line:  # берем элемент из строки
        if l == 0 : # если это первая строка - добавляем еще и столбцы 
         self.tableWd.insertColumn(r)
        self.tableWd.setItem(l, r, QTableWidgetItem(str(item))) # пишем элемент в ячейку
        r = r + 1
       l = l + 1
       r = 0    

##################################   Открытие EXCEL файла ####################################   
   def OpenFileFunc(self):  #открытие Excel файла
     global Flag
     global excel_book # объявляем переменные как глобальные
     global keys      # чтобы иметь возможность передавать значения в другие функции


     filename, filetype = QFileDialog.getOpenFileName(self,
                             "Выбрать файл",
                             ".",
                             "Exel Files(*.xlsx *.xls *.xlsm);;\
                             ") # диалог открытия файла

     if filename != "" :   # если файл выбран
      keys = []
      self.sheetSpin.setDisabled(True)
      self.sheetSpin.setValue(1)
      self.tableWd.setRowCount(0)
      self.tableWd.setColumnCount(0) # обнуляем таблицу
      self.SelFileName.setText(filename) # наш выбранный файл

#для файла из одной страницы      
      #my_array = pyexcel.get_array(file_name=filename) # пишем файл в массив
#предусматриваем много страниц в книге

      self.AllPB.show()
      Flag = 1
      threadOE = self.OpenExcelThread("trPB", 1)
      threadOE.start()

#      excel_book = pyexcel.get_book_dict(file_name=filename) # сохраняем всю книгу в коллекцию
      self.RollPB()

      Flag = 0
      threadOE.join()
  
      for key, item in excel_book.items() :
       keys.append(key)                      # создаем массив названия страниц (ключей массива)
      my_array = excel_book[keys[self.sheetSpin.value() - 1]]    # записываем первую страницу в массив
# выводим массив в таблицу 
      self.FullingTable(my_array)



#      r = 0
#      l = 0
#      for line in my_array:  # записываем массив в таблицу - берем строку
#       self.tableWd.insertRow(l) # добавляем в таблицу строку
#       for item in line:  # берем элемент из строки
#        if l == 0 : # если это первая строка - добавляем еще и столбцы 
#         self.tableWd.insertColumn(r)
#        self.tableWd.setItem(l, r, QTableWidgetItem(str(item))) # пишем элемент в ячейку
#        r = r + 1
#       l = l + 1
#       r = 0
# эту часть мы перенесли в отдельную функцию
      self.sheetSpin.setMaximum(len(keys))
      self.sheetSpin.setEnabled(True)

############################# Сохранение текущей страницы в EXCEL #########################################
   def SaveFileFunc(self):
     global excel_book
     global keys
     self.AllPB.show()
     if self.tableWd.columnCount() != 0 and self.tableWd.rowCount() !=0 :
      save_array=[[0]*self.tableWd.columnCount()]*self.tableWd.rowCount() #инициализируем
      array_line=[0]*self.tableWd.columnCount()                           #массивы
      filename, ok = QFileDialog.getSaveFileName(self, 
                             "Сохранить файл",
                             ".",
                             "Excel Files(*.xls *.xlsx)")    #Диалог выбора файла для сохранения
      #ipdb.set_trace()  #вообще это отладчик-был нужен(((
      if filename != "" :
       try:
        r=0
        l=0
        self.AllPB.setMaximum(self.tableWd.rowCount())
        self.AllPB.setFormat("%v/"+str(self.tableWd.rowCount()))
        self.AllPB.setValue(0)
        while l < self.tableWd.rowCount() :
         while r < self.tableWd.columnCount() :     
          array_line[r] = self.tableWd.item(l, r).text() # набираем строку
          r=r+1
         save_array[l] = array_line  # записываем строку в выходной массив
         array_line = [0]*self.tableWd.columnCount() # обнуляем массив-строку, иначе, при его изменении
         r=0                                        # продолжаеи меняться и выходной массив...
         l=l+1
         self.AllPB.setValue(l);
       # try : сделано для всего блока
        pyexcel.save_as(array=save_array, dest_file_name=filename) # сохраняем массив в файл
       except :
        QMessageBox(QMessageBox.Warning, "Ошибка" , "Имя файла указано неверно!", QMessageBox.Ok).exec()
      filename = ""
      self.AllPB.setValue(0)
      self.AllPB.setFormat("%v/0")
      self.AllPB.hide()

############################################ Сохранение всей книги в EXCEL ###############################################      
   def SaveAllFunc(self):
      global excel_book
      global keys  
      if self.tableWd.columnCount() != 0 and self.tableWd.rowCount() !=0 :
        filename, ok = QFileDialog.getSaveFileName(self, 
                             "Сохранить файл",
                             ".",
                             "Excel Files(*.xls *.xlsx)")    #Диалог выбора файла для сохранения
        if filename != "" :
         
#         try :
#           pyexcel.save_book_as(bookdict=excel_book, dest_file_name=filename)
#         except :
#           QMessageBox(QMessageBox.Warning, "Ошибка" , "Имя файла указано неверно!", QMessageBox.Ok).exec()

          self.AllPB.show()
          Flag = 1
          threadSE = self.SaveExcelThread("trSE", 2, filename)
          threadSE.start()

#          excel_book = pyexcel.get_book_dict(file_name=filename) # сохраняем всю книгу в коллекцию
          self.RollPB()

          Flag = 0
          threadSE.join()
        filename = ""

###########################################  нициализация Главного окна ######################################################
def main():
   global window
   app = QtWidgets.QApplication(sys.argv)
   window = TableApp()
   window.show()
   window.AllPB.hide()
   app.exec_()

if __name__ == '__main__':
   main()