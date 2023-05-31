import matplotlib.pyplot as plt 
from matplotlib import style 
import pandas as pd
import numpy as np
import xlsxwriter 
from emoji import emojize
from sklearn import linear_model
from win32com.client import Dispatch

def say(text):
    speak= Dispatch("SAPI.SpVoice")
    print(text)
    speak.Speak(text)

def sayforinput(text):                       
    speak= Dispatch("SAPI.SpVoice")            
    print(text)                                     
    speak.Speak(text)                             
    a=input()                                     
    return a

def line(a):
    if int(a)==1:
        n=int(sayforinput("Enter number of lines: "))
        
        for i in range(1,n+1):
            x=eval(sayforinput("Enter x points of line " + str(i) ))
            y=eval(sayforinput("Enter y points of line " + str(i) ))
            l=sayforinput("Do you want to add legend for this line? ").lower()
            le=l.split()
            if le[0]=='yes':
                la=sayforinput("Enter name for the legend of this line: ")
            else:
                la=' '
            plt.plot(x,y,label=la)
            plt.legend()
            
        plt.title(sayforinput("Enter title for line graph: "))
        plt.xlabel(sayforinput("Enter label for x axis: "))
        plt.ylabel(sayforinput("Enter label for y axis: "))
        plt.show()
        x1=pd.DataFrame(x)
        reg=linear_model.LinearRegression()
        reg.fit(x1,y)
        pre=sayforinput("Do you want to predict any value of y for given x points? ").lower()
        predict=pre.strip()
        if predict=='yes':
            p=reg.predict([[float(sayforinput('Enter the value for which you need prediction: '))]])
            say('Corresponding value of y= '+str(p[0]))
        say("Looks like you dont have a excel file for given data thus you inputted the data here")
        
        fm=sayforinput("Do you want us to make you a excel file with the given data? ").lower()
        filemaking=fm.strip()
        if filemaking=='yes':
            filename=sayforinput("What name should we give to your file? ").strip()
            f=pd.DataFrame({"xcoordinates":x,"ycoordinates":y})  
            fileexcel=pd.ExcelWriter(filename)
            f.to_excel(fileexcel,sheet_name="Sheet 1")
            fileexcel.save()
        
            
    
    if int(a) == 2:
        n=int(sayforinput("Enter number of lines: "))
            
        for i in range(1,n+1):
            x1=sayforinput("Name of column containing values of x axis: ").strip()
            y1=sayforinput("Name of column containing values of y axis: ").strip()
            x= rf[x1]
            y =rf[y1]
            l=sayforinput("Do you want to add legend for this line? ").lower()
            le=l.split()
            if le[0]=='yes':
                la=sayforinput("Enter name for the legend of this line: ")
            else:
                la=' '
            plt.plot(x,y,label=la)
            plt.legend()
            
        plt.title(sayforinput("Enter title for line graph: "))
        plt.xlabel(sayforinput("Enter label for x axis: "))
        plt.ylabel(sayforinput("Enter label for y axis: "))
        plt.show()
        pre=sayforinput("Do you want to predict any y value for given x points?  ").lower()
        predict=pre.strip()
        reg = linear_model.LinearRegression()
        reg.fit(rf[[x1]],y)
        if predict=='yes':
            p=reg.predict([[float(sayforinput('Enter the value for which you need prediction: '))]])
            say('Corresponding value of y= '+str(p[0]))
    if int(a)==3:
        say("Value before delimeter will be used as x point and Value after delimeter as y point")
        x1,y1 = np.loadtxt(tf,unpack=True,delimiter=m)
        x=[]
        y=[]
        s=int(sayforinput('row number from which the machine should start taking the values'))
        e=int(sayforinput('row number from which the machine should stop taking the values'))
        for i in range(s-1,e):
            x.append(x1[i])
            y.append(y1[i])
        x1=pd.DataFrame(x) 
        l=sayforinput("Do you want to add legend for this line? ").lower()
        le=l.split()
        if le[0]=='yes':
            la=sayforinput("Enter name for the legend of this line: ")
        else:
            la=' '
        plt.plot(x,y,label=la)
        plt.legend()
        plt.title(sayforinput("Enter title for line graph: "))
        plt.xlabel(sayforinput("Enter label for x axis: "))
        plt.ylabel(sayforinput("Enter label for y axis: "))
        plt.show()
        reg=linear_model.LinearRegression()
        reg.fit(x1,y)
        pre=sayforinput("Do you want to predict any y value for given x points?  ").lower()
        predict=pre.strip()
        if predict=='yes':
            p=reg.predict([[float(sayforinput('Enter the value for which you need prediction: '))]])
            say('Corresponding value of y= '+ str(p[0]))


def bar(a):
    if int(a)==1:
        n=int(sayforinput("Enter number of bar charts required: "))
        label=[]
        co=[]
        for i in range(1,n+1):
            x=sayforinput("Enter x label of bar number "+ str(i))
            y=int(sayforinput("Enter corresponding y coordinate: "))
            label.append(x)
            co.append(y)
            plt.bar(label,co)
                
                
        plt.title(sayforinput("Enter title for bar graph: "))
        plt.xlabel(sayforinput("Enter label for x axis: "))
        plt.ylabel(sayforinput("Enter label for y axis: "))
        plt.show()
        say("Looks like you dont have a excel file for given data thus you inputted the data here")
        fm=sayforinput("Do you want us to make you a excel file with the given data? ").lower()
        filemaking=fm.strip() 
        if filemaking=='yes':
            filename=sayforinput("What name should we give to your file? ").strip()
            f=pd.DataFrame({"label":label,"value":co})  
            fileexcel=pd.ExcelWriter(filename)
            f.to_excel(fileexcel,sheet_name="Sheet 1")
            fileexcel.save()
        
    if int(a)==2:
        x1=sayforinput("Name of column containing values of x axis: ").strip()
        y1=sayforinput("Name of column containing values of y axis: ").strip()
        x= rf[x1]
        y =rf[y1]
                
        plt.bar(x,y)
            
        plt.title(sayforinput("Enter title for bar graph: "))
        plt.xlabel(sayforinput("Enter label for x axis: "))
        plt.ylabel(sayforinput("Enter label for y axis: "))
        plt.show()
    if int(a)==3:
        say("Name of x label should be before delimeter and y coordinate after delimeter")
        x1 = np.loadtxt(tf,unpack=True,dtype=str,delimiter=m,usecols=(0))
        y1 = np.loadtxt(tf,unpack=True,delimiter=m,usecols=(1))
        x=[]
        y=[]
        s=int(sayforinput('Row number from which the machine should start taking the values: '))
        e=int(sayforinput('Row number from which the machine should stop taking the values: '))
        for i in range(s-1,e):
            x.append(x1[i])
            y.append(y1[i])
            
        plt.bar(x,y)
        plt.legend()
        plt.title(sayforinput("Enter title for bar graph: "))
        plt.xlabel(sayforinput("Enter label for x axis: "))
        plt.ylabel(sayforinput("Enter label for y axis: "))
        plt.show()

def pie(a):
    if int(a)==1:
        item=[]
        color=[]
        number=eval(sayforinput("Enter list of percentages: "))
            
        for i in range(1,len(number)+1):
            l=sayforinput("Enter name for label number" + str(i)).strip()
            item.append(l)
            
        for i in range(1,len(number)):
            c=sayforinput("Enter colour for label number" + str(i)).strip()
            color.append(c)
            
        plt.pie(number,labels=item,colors=color,startangle=90,autopct='%1.2f%%',shadow=True)
        plt.legend()
        plt.title(sayforinput("Enter title: "))
        plt.show()
        say("Looks like you dont have a excel file for given data thus you inputted the data here")
        fm=sayforinput("Do you want us to make you a excel file with the given data?").lower()
        filemaking=fm.strip() 
        if filemaking=='yes':
            filename=sayforinput("What name should we give to your file? ").strip()
            f=pd.DataFrame({"Percentages":number,"Label":item})  
            fileexcel=pd.ExcelWriter(filename)
            f.to_excel(fileexcel,sheet_name="Sheet 1")
            fileexcel.save()
    if int(a)==2:
        color=[]
        x1=sayforinput("Name of column containing percentages: ").strip()
        y1=sayforinput("Name of column containing names of label: ").strip()
        x= rf[x1]
        y =rf[y1]
            
        for i in range(1,len(y1)):
            c=sayforinput("Enter colour for label number" + str(i)).strip()
            color.append(c)
            
        plt.pie(x,labels=y,colors=color,startangle=90,autopct='%1.2f%%',shadow=True)
        plt.legend()
        plt.title(sayforinput("Enter title: "))
        plt.show()
    if int(a)==3:
        say("Label should be before delimeter and percentage after delimeter")
        x1 = np.loadtxt(tf,unpack=True,dtype=str,delimiter=m,usecols=(0))
        y1 = np.loadtxt(tf,unpack=True,delimiter=m,usecols=(1))
        x=[]
        y=[]
        color=[]
        s=int(sayforinput('Row number from which the machine should start taking the values: '))
        e=int(sayforinput('Row number from which the machine should stop taking the values: '))
        for i in range(s-1,e):
            x.append(x1[i])
            y.append(y1[i])
        for i in range(1,len(y)+1):
            c=sayforinput("Enter colour for label number" + str(i) ).strip()
            color.append(c)
            
        plt.pie(y,labels=x1,colors=color,startangle=90,autopct='%1.2f%%',shadow=True)
        plt.legend()
        plt.title(sayforinput("Enter title: "))
        plt.show()

def scatter(a):
    if int(a)==1:
        x=eval(sayforinput("Enter x axis values: "))
        y=eval(sayforinput("Enter corresponding y axis values: "))
        x1=pd.DataFrame(x)
        plt.title(sayforinput("Enter title for scatter graph: "))
        plt.xlabel(sayforinput("Enter label for x axis: "))
        plt.ylabel(sayforinput("Enter label for y axis: "))
        plt.scatter(x,y)
        reg=linear_model.LinearRegression()
        reg.fit(x1,y)
        bf=sayforinput("Do you want line of best fit? ").lower()
        bof=bf.strip()
        if bof=='yes':
            bestfit= reg.predict(x1)
            plt.plot(x,bestfit)
        plt.show()
        pre=sayforinput("Do you want to predict any y value for given x points? ").lower()
        predict=pre.strip()
        if predict=='yes':
            p=reg.predict([[float(sayforinput('Enter the value for which you need prediction: '))]])
            say('Corresponding value of y= '+ str(p[0]))
        say("Looks like you dont have a excel file for given data thus you inputted the data here")
        fm=sayforinput("Do you want us to make you a excel file with the given data? ").lower()
        filemaking=fm.strip() 
        if filemaking=='yes':
            filename=sayforinput("What name should we give to your file? ").strip()
            first=sayforinput("Input name of column 1: ").strip()
            second=sayforinput("Input name of column 2: ").strip()
            f=pd.DataFrame({first:x,second:y})  
            fileexcel=pd.ExcelWriter(filename)
            f.to_excel(fileexcel,sheet_name="Sheet 1")
            fileexcel.save()
    if int(a)==2:
        x1=sayforinput("Name of column containing values of x axis: ").strip()
        y1=sayforinput("Name of column containing values of y axis: ").strip()
        x= rf[x1]
        y =rf[y1]
        plt.title(sayforinput("Enter title for scatter graph: "))
        plt.xlabel(sayforinput("Enter label for x axis: "))
        plt.ylabel(sayforinput("Enter label for y axis: "))
        reg = linear_model.LinearRegression()
        reg.fit(rf[[x1]],y)
        plt.scatter(x,y)
        bf=sayforinput("Do you want line of best fit? ").lower()
        bof=bf.strip()
        if bof=='yes':
            bestfit= reg.predict(rf[[x1]])
            plt.plot(x,bestfit)
        plt.show()
        pre=sayforinput("Do you want to predict any y value for given x points?  ").lower()
        predict=pre.strip()
        if predict=='yes':
            p=reg.predict([[float(sayforinput('Enter the value for which you need prediction: '))]])
            say('Corresponding value of y= '+ str(p[0]))
    if int(a)==3:
        say("Value before delimeter will be used as x point and Value after delimeter as y point")
        x1,y1 = np.loadtxt(tf,unpack=True,delimiter=m)
        x=[]
        y=[]
        s=int(sayforinput('Row number from which the machine should start taking the values: '))
        e=int(sayforinput('Row number from which the machine should stop taking the values: '))
        for i in range(s-1,e):
            x.append(x1[i])
            y.append(y1[i])
        x1=pd.DataFrame(x)
        plt.title(sayforinput("Enter title for scatter graph: "))
        plt.xlabel(sayforinput("Enter label for x axis: "))
        plt.ylabel(sayforinput("Enter label for y axis: "))
        plt.scatter(x,y)
        reg=linear_model.LinearRegression()
        reg.fit(x1,y)
        bf=sayforinput("Do you want line of best fit? ").lower()
        bof=bf.strip()
        if bof=='yes':
            bestfit= reg.predict(x1)
            plt.plot(x,bestfit)
        plt.show()
        pre=sayforinput("Do you want to predict any y value for given x points?  ").lower()
        predict=pre.strip()
        if predict=='yes':
            p=reg.predict([[float(sayforinput('Enter the value for which you need prediction: '))]])
            say('Corresponding value of y= '+ str(p[0]))
    


w=('WELCOME TO GRAPH MAKER')
say(w.center(50,'-'))

while True:
    t=sayforinput("\n Type 1 if you want to input the data here \n Type 2 if you want to make a graph from a csv file or a excel file \n Type 3 if you want to make a graph from a TEXT file")
    
    
    if int(t)==1:
        d=sayforinput("\n Graph you need \n 1 for Line Graph \n 2 for Bar Chart \n 3 for Pie Chart \n 4 for Scatter Graph")
        
        STY=sayforinput("Enter style (Bmh or Classic or Dark_background) ")
        styleuse=STY.lower()
        style.use(styleuse.split()) 
        if d=='1' :
            line(1)
        if d=='2':
            bar(1)
        if d=='3':
            pie(1)
        if d=='4':
            scatter(1)
        
    
    if int(t)==2:
        ty = sayforinput("Is it a csv file or excel file? ").split()
        type=ty[0].lower()
        if type == 'csv' or type == 'csvfile':
            f = sayforinput("Enter your file name: ").strip()
            rf = pd.read_csv(f)
        elif type == 'excel' or type == 'excelfile':
            f = sayforinput("Enter your file name: ").strip()
            rf = pd.read_excel(f)
        
        d=int(sayforinput("\n Graph you need \n 1 for Line Graph \n 2 for Bar Chart \n 3 for Pie Chart \n 4 for Scatter Graph"))
       
        STY=sayforinput("Enter style (bmh or classic or dark_background)").lower()
        style.use(STY.split()) 
        if d==1:
            line(2)
        if d==2:
            bar(2)
        if d==3:
            pie(2)
        if d==4:
            scatter(2)
    if int(t)==3:
        say("\n So that now you have choosen to make graph from a text file we would like to tell you a rule about it \n the x and y coordianates needs to be seperated either by ,(comma) or :(colon) or any other marking")

        tf = sayforinput("Enter your file name: ").strip()
        m = sayforinput("Enter the seperation marking eg,(comma) ").strip()
        
        d=int(sayforinput("\n Graph you need \n 1 for Line Graph \n 2 for Bar Chart \n 3 for Pie Chart \n 4 for Scatter Graph"))
        
        STY=sayforinput("Enter style (bmh or classic or dark_background)").lower()
        style.use(STY.split()) 
        if d==1:
            line(3)
        if d==2:
            bar(3)
        if d==3:
            pie(3)
        if d==4:
            scatter(3)
    
    con=(sayforinput("Do you want to continue?(Yes or No): ").upper())
    ucon=con.split()
    if ucon[0] =='NO' :
        say("Thankyou")
        print(emojize(":thumbs_up:"))
        break
    elif ucon[0] =='YES':
        say("Here we go again!!")
        print("\U0001f600")
        continue
    else:
        break
