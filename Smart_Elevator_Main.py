##**************************************************************************
#ELEVATOR CALCULATOR DESIGNED AND PROGRAMMED BY :
#SANTHOSH JAYARAJAN .
#PROGRMING START ON 16/11/2021.
#NO EARLIER REVISIONS.
#NOTES:
#Checked on 7/12/2021
#USED 3 NESTED LOOPS TO CALCULATE FOR VARIOUS SPEEDS
#WAITING INTERVALS AND PASSENGER CAPACITY
#ADDED A COMBO BOX FOR SELECTION OF PROJECT TYPE
#
#**************************************************************************#
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import sys
import math
import xlwt
import datetime
from datetime import date

###VARIABLE ASSIGNMENT#








####START OF PROGRAM

root=Tk()
root.wm_title("Smart Elevator Calculator Ver 1.0 Nov 2021 Tata Realty and Infrastructure")
root.resizable(0,0)
root.config(bg="LIGHTGREEN")

# EXCEL FILE NAME GENERATION
Current_Time = datetime.datetime.now()
Hr = Current_Time.hour
Min = Current_Time.minute
Sec = Current_Time.second
date_today=date.today()

File_Name_Excel = "Elevator_Design_Interface " + str(Hr) + "_" + str(Min) + "_" + str(Sec)+".xls"
Excel_Workbook = xlwt.Workbook()

#SETTING OF FONT0 FOR EXCEL
Font0=xlwt.Font()
Font0.name="Calibri"
Font0.bold=True
Font0.height=600

#SETTING FONT1
Font1=xlwt.Font()
Font1.name="Calibri"
Font1.bold=True
Font1.height=200

#SETTING OF FONT 2,3 FOR EXCEL)
Font2=xlwt.Font()
Font2.name="Calibri"
Font2.height=180

Font3=xlwt.Font()
Font3.name="Calibri"
Font3.height=400

Font4=xlwt.Font()
Font4.name="Calibri"
Font4.bold=True
Font4.height=200


#EXCEL STYLE SETTINGS
Style0=xlwt.XFStyle()
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['light_green']
Style0.pattern = pattern
Style0.font=Font0

Style3=xlwt.XFStyle()
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['green']
Style3.pattern = pattern
Style3.font=Font2

Style4=xlwt.XFStyle()
pattern = xlwt.Pattern()
pattern.pattern = xlwt.Pattern.SOLID_PATTERN
pattern.pattern_fore_colour = xlwt.Style.colour_map['yellow']
Style4.pattern = pattern
Style4.font=Font4
Style4.alignment.wrap=1

Style1=xlwt.XFStyle()
Style1.font=Font1
Style2=xlwt.XFStyle()
Style2.font=Font2


#ELEVATOR CALCULATOR SUB ROUTINE
def Calc_Elevator(root):
    #READING ALL THE VARAIABLES
    if Entry.get(Ent1) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                     "Please Enter the Name of Project...  ")
        return
    if Entry.get(Ent2) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Total Site Population...  ")
        return

    if Entry.get(Ent3) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Number of Floors...  ")
        return

    if Entry.get(Ent4) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Floor Height in Meters...  ")
        return

    if Entry.get(Ent5) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Door Open Time in Secs...  ")
        return

    if Entry.get(Ent6) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Door Close Time in Secs...  ")
        return

    if Entry.get(Ent7) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Passenger Transfer Time in Secs...  ")
        return

    if Entry.get(Ent8) == "":
        ret_err1 = messagebox.showinfo("Data Entry Error",
                                       "Please Enter the Extra Time or 0.0 for zero time..  ")
        return

    # ADD THE PROJECT NAME WORKSHEET

    Excel_WorkSheet = Excel_Workbook.add_sheet('Project_' + str(Entry.get(Ent1)))
    Excel_Summary = Excel_Workbook.add_sheet('Project_' + str(Entry.get(Ent1)+"_Summary"))
    Excel_WorkSheet.write_merge(0, 1, 0, 6, "Tata Realty and Infrastructure - Elevator Design Data", Style0)
    Excel_WorkSheet.write(2, 0, "Project:" + str(Entry.get(Ent1)), Style1)
    Excel_WorkSheet.write(3, 0, "Date of Report:" + str(date_today), Style1)
    Excel_WorkSheet.write(4, 0, "Time of Report:" + str(Hr) + ":" + str(Min) + ":" + str(Sec), Style1)
    Excel_WorkSheet.write(6, 0, "Sl.No", Style4)
    Excel_WorkSheet.write(6, 1, "Speed-M/s", Style4)
    Excel_WorkSheet.write(6, 2, "Number of Elevators", Style4)
    Excel_WorkSheet.write(6, 3, "Passenger Capacity", Style4)
    Excel_WorkSheet.write(6, 4, "RTT(Secs)", Style4)
    Excel_WorkSheet.write(6, 5, "Waiting Interval (Secs)", Style4)
    Excel_WorkSheet.write(6, 6, "5 Min Handling Capacity-%", Style4)
    Excel_WorkSheet.write(6, 7, "H-Highest Reversal Floor", Style4)
    Excel_WorkSheet.write(6, 8, "Time for Traversal of full Height", Style4)

    #SUMMARY SHEET FILLING DATA
    Excel_Summary.write_merge(0, 1, 0, 6, "Results Summary", Style0)
    Excel_Summary.write(2, 0, "Project:" + str(Entry.get(Ent1)), Style1)
    Excel_Summary.write(3, 0, "Date of Report:" + str(date_today), Style1)
    Excel_Summary.write(4, 0, "Time of Report:" + str(Hr) + ":" + str(Min) + ":" + str(Sec), Style1)
    Excel_Summary.write(6, 0, "Sl.No", Style4)
    Excel_Summary.write(6, 1, "Speed-M/s", Style4)
    Excel_Summary.write(6, 2, "Waiting Interval (Secs)", Style4)
    Excel_Summary.write(6, 3, "5 Min Handling Capacity(%)", Style4)
    Excel_Summary.write(6, 4, "RTT(Secs)", Style4)
    Excel_Summary.write(6, 5, "Number Of Elevators)", Style4)
    Excel_Summary.write(6, 6, "Traverse Time(Secs)", Style4)
    Excel_Summary.write(6, 7, "Average Number of Stops(S)", Style4)

    #DATA AREA
    Speed_Range=[0.75,1.0,1.5,1.75,2.0,2.5,3.0]
    Elev_Capacity=[6,8,12,16,18,20,24]
    Waiting_Interval=[item for item in range(1,60,1)]
    Population = int(Entry.get(Ent2))
    Excel_WorkSheet.write(5, 0, "Site Population:"+str(Population), Style1)
    inc=1
    inc_sum=1

    #GETTING THE BUILDING TYPE
    build_type=building_type.get()
    if build_type ==' Multi Tenant Commercial':
        FMHCMin=10
    if build_type ==' Single Tenant Commercial':
        FMHCMin = 15
    if build_type ==' Residential':
        FMHCMin=7.5

        # GETTING THE MINIMUM ELEVATOR SPEED
    min_elev_speed = Min_Elev_Speed.get()
    if min_elev_speed == '0.75':
        MINELEVSPD = 0.75
    if min_elev_speed == '1.0':
        MINELEVSPD = 1.0
    if min_elev_speed == '1.5':
        MINELEVSPD = 1.5
    if min_elev_speed == '1.75':
        MINELEVSPD = 1.75
    if min_elev_speed == '2.0':
        MINELEVSPD = 2.0
    if min_elev_speed == '2.5':
        MINELEVSPD = 2.5
    if min_elev_speed == '3.0':
        MINELEVSPD = 3.0

        # GETTING THE WAITING INTERVAL CATEGORY

    wait_int_cat = Wait_Int_Cat.get()
    if wait_int_cat == 'Excellent':
        Waiting_Int_Min = 0.0
        Waiting_Int_Max = 25
    if wait_int_cat == 'Very Good':
        Waiting_Int_Min = 0.0
        Waiting_Int_Max = 30.0
    if wait_int_cat == 'Good':
        Waiting_Int_Min = 0.0
        Waiting_Int_Max = 35.0
    if wait_int_cat == 'Fair':
        Waiting_Int_Min = 0.0
        Waiting_Int_Max = 40.0
    if wait_int_cat == 'Poor':
        Waiting_Int_Min = 0.0
        Waiting_Int_Max = 45.0
    if wait_int_cat == 'Unsatisfactory':
        Waiting_Int_Min = 45.0
        Waiting_Int_Max = 100.0

    #START OF NESTED LOOPS
    for k in Waiting_Interval:
        for j in Elev_Capacity:
            for i in Speed_Range:
    # RETRIVING ALL DATA FROM FORM AND CALCULATING
                Tex=float(Entry.get(Ent8))
                N=int(Entry.get(Ent3))-1
                tp=float(Entry.get(Ent5))
                Toc=float(Entry.get(Ent5))+float(Entry.get(Ent6))
                tp=float(Entry.get(Ent7))
                Population=int(Entry.get(Ent2))
                Speed=i
                Acc=1.0
                Pmax=j
                P=0.8*Pmax
                tv=float(Entry.get(Ent4))/Speed
                tf1=Speed/Acc
                T=Toc+tf1+Tex
                INT=k
                Elevator_Speed_Min=MINELEVSPD

                #TOTAL HEIGHT OF TRAVEL

                TOTHEIGHT=int(Entry.get(Ent3))*float(Entry.get(Ent4))

                #CALCULATION OF H - AVEARGE HIGHEST REVERSAL FLOOR
                HSum=0
                for i in range(1,N-1):
                    HSum=HSum+(i/N)**P
                H=N-HSum

                #CALCULATION OF S
                Scratch1= (1-1/N)**P
                S=N*(1-(Scratch1))

                #CALCULATION OF RTT
                RTT=2*H*tv+(S+1)*(T-tv)+2*P*tp

                #CALACULATION OF NUMBER OF ELEVATORS

                L=RTT/INT

                #CAKLCULATION OF 5 MIN HANDLING CAPACITY

                FIVEMINHC=(300*P*L*100)/(RTT*Population)

                TRAV_TIME=(int(Entry.get(Ent3))*float(Entry.get(Ent4)))/Speed

                #SEND DATA TO EXCEL

                Excel_WorkSheet.write(inc + 6, 0, inc, Style2)
                Excel_WorkSheet.write(inc + 6, 1,Speed, Style2)
                Excel_WorkSheet.write(inc + 6, 2,math.ceil(L), Style2)
                Excel_WorkSheet.write(inc + 6, 3, round(Pmax, 1), Style2)
                Excel_WorkSheet.write(inc + 6, 4,round(RTT,2), Style2)
                Excel_WorkSheet.write(inc + 6, 5,round(INT,2), Style2)

                if FIVEMINHC >= FMHCMin:
                    Excel_WorkSheet.write(inc + 6, 6, round(FIVEMINHC, 2), Style3)
                else:
                    Excel_WorkSheet.write(inc + 6, 6, round(FIVEMINHC, 2), Style2)

                Excel_WorkSheet.write(inc + 6, 7, round(H,2), Style2)
                Excel_WorkSheet.write(inc + 6, 8, round(TRAV_TIME, 2), Style2)
                inc=inc+1
                #SUMMARY DATA
                if (Speed>=Elevator_Speed_Min and FIVEMINHC>=FMHCMin and (INT>=Waiting_Int_Min and INT<=Waiting_Int_Max)and TRAV_TIME<=60.0 ):
                    Excel_Summary.write(inc_sum + 6, 0, inc_sum, Style2)
                    Excel_Summary.write(inc_sum + 6, 1, Speed, Style2)
                    Excel_Summary.write(inc_sum + 6, 2, INT, Style2)
                    Excel_Summary.write(inc_sum + 6, 3, round(FIVEMINHC,3), Style2)
                    Excel_Summary.write(inc_sum + 6, 4, round(RTT,3), Style2)
                    Excel_Summary.write(inc_sum + 6, 5, math.ceil(L), Style2)
                    Excel_Summary.write(inc_sum + 6, 6, round(TRAV_TIME, 3), Style2)
                    Excel_Summary.write(inc_sum + 6, 7, round(S, 3), Style2)
                    inc_sum=inc_sum+1


    # BASIC RESULTS DISPLAY
                Ent9 = Label(text="{:.2f}".format(H), relief=RIDGE, width=13).grid(row=18, column=1)
                Ent10 = Label(text="{:.2f}".format(S), relief=RIDGE, width=13).grid(row=19, column=1)
                Ent11 = Label(text="{:.2f}".format(TOTHEIGHT), relief=RIDGE, width=13).grid(row=20, column=1)
                Ent12 = Label(text="{:.2f}".format(tv), relief=RIDGE, width=13).grid(row=21, column=1, padx=10, pady=1, )
                Ent13 = Label(text="{:.2f}".format(T), relief=RIDGE, width=13).grid(row=22, column=1, padx=10, pady=1)

    Excel_WorkSheet.write(inc + 8, 0, " Note: As per NBC 2016", Style2)
    Excel_WorkSheet.write(inc + 9, 0, " 5 Min Handling Capacity:", Style2)
    Excel_WorkSheet.write(inc + 10, 0, " ---Multi Tenant = 10 -15 %", Style2)
    Excel_WorkSheet.write(inc + 11, 0, " ---Single Tenant = 15-25 %", Style2)
    Excel_WorkSheet.write(inc + 12, 0, " Waiting Interval:", Style2)
    Excel_WorkSheet.write(inc + 13, 0, " ---Excellent = Less Than 25 Secs", Style2)
    Excel_WorkSheet.write(inc + 14, 0, " ---Very Good = 25 to less than 30 Secs", Style2)
    Excel_WorkSheet.write(inc + 15, 0, " ---Good = 30 to less than 35 Secs", Style2)
    Excel_WorkSheet.write(inc + 16, 0, " ---Fair = 35 to less than 40 Secs", Style2)
    Excel_WorkSheet.write(inc + 17, 0, " ---Poor = 40 to less than 45 Secs", Style2)
    Excel_WorkSheet.write(inc + 18, 0, " ---Unsatisfactory = Greater than 45", Style2)

#SUMMARY SHEET FILLING
    Excel_Summary.write(inc_sum + 8, 0, " Note: As per NBC 2016", Style2)
    Excel_Summary.write(inc_sum + 9, 0, " 5 Min Handling Capacity:", Style2)
    Excel_Summary.write(inc_sum + 10, 0, " ---Multi Tenant = 10 -15 %", Style2)
    Excel_Summary.write(inc_sum + 11, 0, " ---Single Tenant = 15-25 %", Style2)
    Excel_Summary.write(inc_sum + 12, 0, " Waiting Interval:", Style2)
    Excel_Summary.write(inc_sum + 13, 0, " ---Excellent = Less Than 25 Secs", Style2)
    Excel_Summary.write(inc_sum + 14, 0, " ---Very Good = 25 to less than 30 Secs", Style2)
    Excel_Summary.write(inc_sum + 15, 0, " ---Good = 30 to less than 35 Secs", Style2)
    Excel_Summary.write(inc_sum + 16, 0, " ---Fair = 35 to less than 40 Secs", Style2)
    Excel_Summary.write(inc_sum + 17, 0, " ---Poor = 40 to less than 45 Secs", Style2)
    Excel_Summary.write(inc_sum + 18, 0, " ---Unsatisfactory = Greater than 45", Style2)



    Excel_Workbook.save(File_Name_Excel)
    done = messagebox.showinfo("Results Data..","Results Stored in Excel File :"+str(File_Name_Excel)+" With "+ str(inc)+" of Combinations...Please check Excel file")


def About_Prog(root):
    rss=messagebox.showinfo("About Elevator Calculator","Designed and Programmed by Santhosh Jayarajan in November 2021")

def Quitprog(root):
    sys.exit()

##CATEGORY ELEVATOR PROJECT DETAILS

Col_main_lbl1=Label(text="Category:Basic Project Details",bg="lightblue",fg="red",font=16, relief=RIDGE,width=55).grid(row=0,column=0)

Lbl1=Label(text='Name of Project:', relief=RIDGE,width=55).grid(row=1,column=0)
Ent1 = Entry(bg='yellow', relief=SUNKEN, width=35)
Ent1.grid(row=1,column=1,padx=0,pady=0)

Lbl2=Label(text='Total Population:', relief=RIDGE,width=55).grid(row=2,column=0)
Ent2= Entry(bg='yellow', relief=SUNKEN, width=15)
Ent2.grid(row=2,column=1,padx=0,pady=0)


#CATEGORY BUILDING DATA

Col_main_lbl2=Label(text="Category: Building Data",bg="lightblue",fg="red",font=16, relief=RIDGE,width=55).grid(row=3,column=0)
Lbl3=Label(text='Number of Floors:', relief=RIDGE,width=55).grid(row=5,column=0)
Ent3 = Entry(bg='yellow', relief=SUNKEN, width=15)
Ent3.grid(row=5,column=1)

Lbl4=Label(text='Floor Height in Meters:', relief=RIDGE,width=55).grid(row=6,column=0)
Ent4 = Entry(bg='yellow', relief=SUNKEN, width=15)
Ent4.grid(row=6,column=1)


#CATEGORY ELEVATOR OPERATION DATA
Col_main_lbl3=Label(text="Category: Elevator Operation Data",bg="lightblue",fg="red",font=16, relief=RIDGE,width=55).grid(row=8,column=0)
Lbl5=Label(text='Door Open Time in Secs:', relief=RIDGE,width=55).grid(row=9,column=0)
Ent5 = Entry(bg='yellow', relief=SUNKEN, width=15)
Ent5.grid(row=9,column=1)

Lbl6=Label(text='Door Close Time in Secs:', relief=RIDGE,width=55).grid(row=10,column=0)
Ent6 = Entry(bg='yellow', relief=SUNKEN, width=15)
Ent6.grid(row=10,column=1)

Lbl7=Label(text='Passenger Transfer Time tp in Secs:', relief=RIDGE,width=55).grid(row=11,column=0)
Ent7 = Entry(bg='yellow', relief=SUNKEN, width=15)
Ent7.grid(row=11,column=1)

Lbl8=Label(text='Extra Time in Secs:', relief=RIDGE,width=55).grid(row=12,column=0)
Ent8 = Entry(bg='yellow', relief=SUNKEN, width=15)
Ent8.grid(row=12,column=1)

Lbl9=Label(text='Building Type Selection', relief=RIDGE,width=55).grid(row=14,column=0)

building_type = ttk.Combobox(root, width=27,state="readonly")
building_type.grid(row=14,column=1)
building_type['values'] = (' Multi Tenant Commercial',
                          ' Single Tenant Commercial',
                          ' Residential'
                          )
#MINIMUM ELEVATOR SPEED
Lbl14=Label(text='Minimum Elevator Speed(m/s)', relief=RIDGE,width=55).grid(row=15,column=0)
Min_Elev_Speed = ttk.Combobox(root, width=27,state="readonly")
Min_Elev_Speed.grid(row=15,column=1)
Min_Elev_Speed['values'] = ('0.75',
                          '1.0',
                          '1.5',
                          '1.75',
                          '2.0',
                          '2.5',
                          '3.0'
                          )

#WAITING INTERVAL CATEGORY
Lbl15=Label(text='Waiting Interval Category', relief=RIDGE,width=55).grid(row=16,column=0)
Wait_Int_Cat = ttk.Combobox(root, width=27,state="readonly")
Wait_Int_Cat.grid(row=16,column=1)
Wait_Int_Cat['values'] = ('Excellent',
                          'Very Good',
                          'Good',
                          'Fair',
                          'Poor',
                          'Unsatisfactory'
                          )

#CATEGORY BASIC RESULTS
Col_main_lbl4=Label(text="Category: Basic Elevator Results:",bg="lightblue",fg="red",font=16, relief=RIDGE,width=55).grid(row=17,column=0)
Lbl9=Label(text='Highest Reversal Floor (H):', relief=RIDGE,width=55,bg="lightgreen").grid(row=18,column=0)
Ent9=Label(text='', relief=RIDGE,width=13).grid(row=18,column=1)


Lbl10=Label(text='Average Number of Stops (S) :', relief=RIDGE,width=55,bg="lightgreen").grid(row=19,column=0)
Ent10=Label(text='', relief=RIDGE,width=13).grid(row=19,column=1)

Lbl11=Label(text='Total Height of Travel :', relief=RIDGE,width=55,bg="lightgreen").grid(row=20,column=0)
Ent11=Label(text='', relief=RIDGE,width=13).grid(row=20,column=1)

Lbl12=Label(text='Single Floor Transit Time (Tv) in Secs', relief=RIDGE,width=55,bg="lightgreen").grid(row=21,column=0)
Ent12=Label(text='', relief=RIDGE,width=13).grid(row=21,column=1,padx=10, pady=1,)

Lbl13=Label(text='Total Door Open Time (T) in Secs', relief=RIDGE,width=55,bg="lightgreen").grid(row=22,column=0)
Ent13=Label(text='', relief=RIDGE,width=13).grid(row=22,column=1,padx=10, pady=1)

# SPACE AT THE END OF THE DATA ENTRY FORM
Col_main_sep4=Label(text="", width=50,bg="LIGHTGREEN",).grid(row=23,column=0)


Calculate_Elevator = Button(root,text='Calculate Elevator Data')
Calculate_Elevator.bind("<ButtonRelease-1>", Calc_Elevator)
Calculate_Elevator.grid(row=1,column=3,padx=5, pady=2)

About_Elevator = Button(root,text= '   About  ')
About_Elevator.bind("<ButtonRelease-1>", About_Prog)
About_Elevator.grid(row=2,column=3,padx=5, pady=2)

Quit1 = Button(root,text='         Quit        ')
Quit1.bind("<ButtonRelease-1>", Quitprog)
Quit1.grid(row=3,column=3,padx=5, pady=2)


root.mainloop()
