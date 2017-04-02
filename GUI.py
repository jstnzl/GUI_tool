import Tkinter
from Tkinter import *
import ttk
from glob import glob
import os
import xlwt
from xlwt.Utils import cell_to_rowcol2

class simpleapp_tk(Tkinter.Tk):
   def __init__(self,parent):
      Tkinter.Tk.__init__(self,parent)
      self.parent=parent
      self.initialize()

   def initialize(self):
      self.grid()
      self.grid_columnconfigure(2,weight=1)
      self.resizable(False,False)
      self.geometry('{}x{}'.format(900, 600))
      
      # Refresh, save, save_list buttons
      refresh = Tkinter.Button(self,text=u"Refresh",command=self.rButton,bg="mediumspringgreen")
      refresh.grid(column=3,row=0, pady=10, padx=(10,0))
      
      save = Tkinter.Button(self,text=u"Save",command=self.sButton)
      save.grid(column=4,row=1)
      
      saveList = Tkinter.Button(self,text=u"Save List",command=self.sLButton)
      saveList.grid(column=4,row=0, pady=10,padx=5)      
      
      # Protocol Label
      self.lVar = Tkinter.StringVar()
      label = Tkinter.Label(self,textvariable= self.lVar,fg="black",bg="yellow",font = "Helvetica 13 bold")
      label.grid(column=2,row=1,columnspan=2,sticky='EW', padx=10)
      self.lVar.set("Protocol")
      
      # Protocol Combobox
      self.pVar = Tkinter.StringVar()
      self.pVar.set(u"Protocols")
      self.pCombo = ttk.Combobox(self,textvariable=self.pVar, state='readonly')
      self.pCombo.grid(column=2,row=0,sticky='EW', pady=10)
      self.pCombo.bind("<<ComboboxSelected>>", self.pEvent)
      
      # Data File Combobox
      self.dVar = Tkinter.StringVar()
      self.dVar.set(u"Data Files")
      self.files = []
      start_dir = os.getcwd()
      pattern = "*.txt"
      for file in [f for f in os.listdir('DataFiles') if f.endswith('.txt')]:
         self.files.append(file)  
      self.dCombo = ttk.Combobox(self,textvariable=self.dVar,values=self.files,state='readonly')
      self.dCombo.grid(column=0,row=1,sticky='EW', padx=10, columnspan=2)
      self.dCombo.bind("<<ComboboxSelected>>", self.dEvent)
      
      # Clinical Sections LabelFrame
      self.cSec = Tkinter.LabelFrame(self,text="Clinical Sections",labelanchor='n',width=200,height=300,font="Helvetica 10 bold")
      self.cSec.grid(column=0, row=2, rowspan=5, columnspan=2,padx=20,sticky='W')
      
      # CT Scanning Parameters LabelFrame
      self.CT = LabelFrame(self,text="CT Scanning Parameters",width=700,height=300,font="Helvetica 11 bold")
      self.CT.grid(column=2, row=2, rowspan=10, columnspan=10,padx=20,pady=10,sticky='NSEW')
      
      # Image Reconstruction Parameters LabelFrame
      self.imageRec = LabelFrame(self, text = "Image Reconstruction Parameters", width=700,height=200, font="Helvetica 11 bold")
      self.imageRec.grid(column=2, row=15, rowspan=5, columnspan=5,padx=20,sticky='NSEW')
      
      # Checkboxes in Clinical Sections
      self.checkDict = {'Abdomen/Pelvis':0,'Chest':0,'Neuro':0,'Extremities':0,'Miscellaneous':0}
      count = 0
      for key in self.checkDict:
         self.checkDict[key] = IntVar()
         aCheckButton = Checkbutton(self.cSec, text=key, font="Helvetica 10 bold", variable=self.checkDict[key])
         aCheckButton.grid(column=0,row=count,padx=5,sticky= 'W')
         count = count+1
      
      # Radiobuttons
      self.radio = IntVar()
      adult = Radiobutton(self, text = "Adult ", bg = "green2", font="Helvetica 10 bold", variable = self.radio, value=1)
      adult.grid(column=0,row=7,padx=(20,0),sticky= 'W')
      ped = Radiobutton(self, text = "Pediatric ", bg = "palevioletred1", font="Helvetica 10 bold", variable = self.radio, value=2)
      ped.grid(column=1,row=7,sticky= 'W')
      
      
   def dEvent(self, event):
      self.var_data = self.dCombo.get()     

   def sButton(self):
      fname = Tkinter.StringVar(value=str(self.var_protocol)+".xls")
      sheet = Tkinter.StringVar(value= "Parameters")
      row = col = max = 0
      wb = xlwt.Workbook()
      ws = wb.add_sheet(sheet.get())
      #CT
      for col in range(self.CTcols):
         for row in range(0,11):
            elem = self.CTcolumns[col][row]
            ws.write(row,col,elem)
            row = row+1
            count = len(elem)
            if count > max:
               max = count
         col=col+1
      # RECON
      col1 = max1 = 0
      for col in range(self.reconCols):
         row1 = 14
         for row in range(0,5):
            elem = self.reconColumns[col][row]
            ws.write(row1,col,elem)
            row1 = row1+1
            count = len(elem)
            if count > max1:
               max1 = count
            if max1 > max:
               max = max1
         col1=col1+1
      if col1 > col:
         col = col1
      for col in range(col):
         ws.col(col).width = (max)*275
      wb.save(fname.get())

   def sLButton(self):
      fname = Tkinter.StringVar(value="saveList.xls")
      sheet = Tkinter.StringVar(value= self.var_data)
      row = col = max = 0
      wb = xlwt.Workbook()
      ws = wb.add_sheet(sheet.get())
      for elem in self.valueList:
         ws.write(row,col,elem)
         row = row+1
         count = len(elem)
         if count > max:
            max = count
      ws.col(0).width = (max+1)*275
      wb.save(fname.get())
      
   def pEvent(self, event):
      self.CTlabels = ["","kV","Rotation time","Type","QRM","Ref kV","Eff.mAs","Collimation","Pitch","CarekV","CareDose4D"]
      self.reconLabels = ["","Thickness","Inc","Kernel","IR"]
      self.CTcolumns=[self.CTlabels]
      self.reconColumns = [self.reconLabels]
      self.CTcols,self.reconCols = 1,1
      self.var_protocol = self.pCombo.get()
      self.lVar.set(self.var_protocol)       # change label
      i = 0
      for value in self.valueList:           # find index for label
         if value == self.var_protocol:
            break
         i = i+1
      start = self.countList[i]+3              # startline
      for num in self.endList:               
         if start < num:
            end = num                           # endline
            break
      with open(self.var_data, 'r') as file:
         lines = file.readlines()
         for count in range(0,len(lines)):
            line = lines[count]
            if count >= end:
               break
            elif count >= start:
               if "+" == line[0]:
                  CTtemp = []
                  self.CTcols = self.CTcols+1
                  CTtemp.append(line[2:-1])
                  CTtemp1 = (lines[count+1])[:-1]
                  for splits in CTtemp1.split("\t"):
                     CTtemp.append(splits)
                  self.CTcolumns.append(CTtemp)
                  reconLine = count+2
                  while ("+"!=lines[reconLine][0])&(reconLine<end):
                     self.reconCols = self.reconCols+1
                     reconTemp=[]
                     reconTemp.append(line[2:-1])
                     reconTemp1 = (lines[reconLine][:-1])
                     for splits in reconTemp1.split("\t"):
                        reconTemp.append(splits)
                     self.reconColumns.append(reconTemp)
                     reconLine = reconLine+1
      self.CTTableMaker(11,self.CTcols)
      self.reconTableMaker(5,self.reconCols)
      
   def CTTableMaker(self,rows,cols):
      CTmaxcol = 7
      self.CTtable = Frame(self.CT,bg="black")  
      self.CTtable.grid(row=3, column=3,padx=10,pady=(5,10),sticky='NSEW')
      for column in range(CTmaxcol):
         if column in range(cols):
            for row in range(rows):
               label = Tkinter.Label(self.CTtable,text=self.CTcolumns[column][row],width=11,bg="white",font="Helvetica 10 bold",borderwidth=0) 
               self.gridFormat(label,row,column)
         # fill empty columns
         else:
            for row in range(rows):
               label = Tkinter.Label(self.CTtable,text="",width=11,bg="white",font="Helvetica 10 bold",borderwidth=0)
               self.gridFormat(label,row,column)

   def reconTableMaker(self,rows,cols):
      reconMaxcol = 7
      self.reconTable = Frame(self.imageRec,bg="black")  
      self.reconTable.grid(row=3, column=3,padx=10,pady=(5,10),sticky='NSEW')
      for column in range(reconMaxcol):
         if column in range(cols):
            for row in range(rows):
               if column==0:
                  label = Tkinter.Label(self.reconTable,text=self.reconColumns[column][row],width=8,bg="white",font="Helvetica 10 bold",borderwidth=0)
               else:
                  label = Tkinter.Label(self.reconTable,text=self.reconColumns[column][row],width=11,bg="white",font="Helvetica 10 bold",borderwidth=0) 
               self.gridFormat(label,row,column)
         # fill empty columns
         else:
            for row in range(rows):
               label = Tkinter.Label(self.reconTable,text="",width=11,bg="white",font="Helvetica 10 bold",borderwidth=0)
               self.gridFormat(label,row,column)               
   
   def gridFormat(self,label,row,column):               
      if row==0 and column==0:
         label.grid(row=row, column=column, sticky="NSEW", padx=1, pady=1)
      elif row==0 and column!=0:
         label.grid(row=row, column=column, sticky="NSEW", padx=(0,1), pady=1)
      elif row!=0 and column==0:
         label.grid(row=row, column=column, sticky="NSEW", padx=1, pady=(0,1))   
      else:
         label.grid(row=row, column=column, sticky="NSEW", padx=(0,1), pady=(0,1))
         
   def rButton(self):
      self.protocolList,self.pedList,self.adultList,self.cList,self.cpedList,self.caList = [],[],[],[],[],[]
      self.keyList,self.endList,self.valueList,countList = [],[],[],[]
      self.neuroList,self.chestList,self.abList,self.exList,self.miscList= [],[],[],[],[]
      self.neuroCount,self.chestCount,self.abCount,self.exCount,self.miscCount= [],[],[],[],[]
      # search file
      with open(self.var_data, 'r') as file:
         lines = file.readlines()
         for count in range(0,len(lines)):
            line = lines[count]
            if "*" == line[0]:
               value = line[1:-1]
               self.protocolList.append(value)
               self.cList.append(count)
               if "Child" in lines[count+1]:
                  self.pedList.append(value)  
                  self.cpedList.append(count)
               elif "Adult" in lines[count+1]:
                  self.adultList.append(value)
                  self.caList.append(count)
               check = lines[count+2]
               if ("Head" in check)|("Spine" in check)|("Neck" in check)|("Shoulder" in check):
                  self.neuroList.append(value)
                  self.neuroCount.append(count)
               elif ("Thorax" in check)|("Cardiac" in check)|("Vascular" in check):
                  self.chestList.append(value)
                  self.chestCount.append(count)
               elif ("Abdomen" in check)|("Pelvis" in check):
                  self.abList.append(value)
                  self.abCount.append(count)
               elif "Extremities" in check:
                  self.exList.append(value)
                  self.exCount.append(count)
               else:
                  self.miscList.append(value)
                  self.miscCount.append(count)
            elif "#" == line[0]:
               self.endList.append(count)
      # Which protocols to list (radiobuttons)
      self.listofLists_check = [self.neuroList,self.chestList,self.abList,self.exList,self.miscList]
      self.listofLists_count = [self.neuroCount,self.chestCount,self.abCount,self.exCount,self.miscCount]
      if self.radio.get() == 1:
         self.valueList = self.adultList
         self.countList = self.caList
         self.checkingBox()
      elif self.radio.get() == 2:
         self.valueList = self.pedList
         self.countList = self.cpedList
         self.checkingBox()
      else:
         self.valueList = self.protocolList
         self.countList = self.cList
         self.checkingBox()
   
   # find which protocols to use with checkboxes
   def checkingBox(self):  
      self.checkUnion,self.countUnion = [],[]
      count,boxChecked = 0,0        
      for key, value in self.checkDict.items():
         if value.get() != 0:
            checkSet = set(self.checkUnion)
            countSet = set(self.countUnion)
            buttonSet = set(self.listofLists_check[count])
            buttonCount = set(self.listofLists_count[count])
            self.checkUnion = list(set.union(checkSet,buttonSet))
            self.countUnion = list(set.union(countSet,buttonCount))
            boxChecked=boxChecked+1
         count=count+1
      valueSet = set(self.valueList)
      checkSet = set(self.checkUnion)
      countList_set = set(self.countList)
      countSet = set(self.countUnion)
      if boxChecked!=0:
         self.valueList = list(set.intersection(valueSet,checkSet))
         self.countList = list(set.intersection(countList_set,countSet))
      self.pCombo['values'] = (self.valueList)
      
if __name__ == "__main__":
   app = simpleapp_tk(None)
   app.title('GUI')
   app.mainloop()
   