import customtkinter as ctk
from tkinter import filedialog
from weeklyCheck.weeklyJig import generateTestJigCert, exportJigCert
import os
from PIL import Image

def uiTestJig(mainScreen):
    
    global benchType, testJigType, voltageType, dataType, tempType, generateButton, openCertButton, exportCert, meterPointType
    
    lineColor = ctk.CTkImage(
    light_image=Image.open(r"images\linecolor.png"),
    dark_image=Image.open(r"images\linecolor.png"),
    size=(150,1),
    )
    lineColorPic = ctk.CTkLabel(mainScreen, image=lineColor, text="", bg_color="#2B2B2B")
    lineColorPic.place(x=0,y=20)

    lineColor2 = ctk.CTkImage(
    light_image=Image.open(r"images\linecolor.png"),
    dark_image=Image.open(r"images\linecolor.png"),
    size=(550,2),
    )
    lineColorPic2 = ctk.CTkLabel(mainScreen, image=lineColor2, text="", bg_color="#2B2B2B")
    lineColorPic2.place(x=0,y=250)

    darkbackground = ctk.CTkImage(
    light_image=Image.open(r"images\backgroundcolor.png"),
    dark_image=Image.open(r"images\backgroundcolor.png"),
    size=(550,255),
    )
    darkbackgroundl = ctk.CTkLabel(mainScreen, image=darkbackground, text="")
    darkbackgroundl.place(x=0,y=266)

    title = ctk.CTkLabel(
        mainScreen,
        text="Required Data",
        font=("Arial", 20, "bold"),
        text_color="white",
    )
    title.place(x=0, y=5)

    add1 = ctk.CTkLabel(mainScreen, text="➤", text_color="white", font=("Arial", 20))
    add1.place(x=0, y=50)

    tempText = ctk.CTkLabel(
        mainScreen, text="Room Temp in C:", font=("Arial", 13), text_color="white"
    )
    tempText.place(x=30, y=50)
    
    tempType = ctk.CTkEntry(mainScreen)
    tempType.place(x=135, y=50)

    add2 = ctk.CTkLabel(mainScreen, text="➤", text_color="white", font=("Arial", 20))
    add2.place(x=0, y=90)

    benchText = ctk.CTkLabel(
        mainScreen, text="Bench:", font=("Arial", 13), text_color="white"
    )
    benchText.place(x=50, y=90)
    
    benchType = ctk.CTkComboBox(mainScreen, values=['WECO2350', 'WECO4150'], command=changetestJig)
    benchType.place(x=135, y=90)
    benchType.set("")


    add3 = ctk.CTkLabel(mainScreen, text="➤", text_color="white", font=("Arial", 20))
    add3.place(x=0, y=130)

    testJigText = ctk.CTkLabel(
        mainScreen, text="Test Jig:", font=("Arial", 13), text_color="white"
    )
    testJigText.place(x=50, y=130)
    
    testJigType = ctk.CTkComboBox(mainScreen, values=[""], command=voltageSelection)
    testJigType.place(x=135, y=130)

    add4 = ctk.CTkLabel(mainScreen, text="➤", text_color="white", font=("Arial", 20))
    add4.place(x=0, y=170)

    voltageText = ctk.CTkLabel(
        mainScreen, text="Voltage:", font=("Arial", 13), text_color="white"
    )
    voltageText.place(x=50, y=170)
    
    voltageType = ctk.CTkComboBox(mainScreen, values=[""])
    voltageType.place(x=135, y=170)

    add4 = ctk.CTkLabel(mainScreen, text="➤", text_color="white", font=("Arial", 20))
    add4.place(x=0, y=210)

    meterPoint = ctk.CTkLabel(
        mainScreen, text="Meter Points:", font=("Arial", 13), text_color="white"
    )
    meterPoint.place(x=45, y=210)
    
    meterPointType = ctk.CTkEntry(mainScreen)
    meterPointType.place(x=135, y=210)

    title2 = ctk.CTkLabel(
        mainScreen, text="Submit Files", font=("Arial", 20, "bold"), text_color="white", bg_color="#061E40"
    )
    title2.place(x=210, y=275)

    dataType = ctk.CTkButton(
        mainScreen, text="Data File (.xlsx)", fg_color="#061E40", command=submitData, border_color="white", border_width=2
    )
    dataType.place(x=200, y=330)

    generateButton = ctk.CTkButton(
        mainScreen, text="Generate Certificate", fg_color="#061E40", command=genJigCert, border_color="white", border_width=2
    )

    generateButton.place(x=200, y=380)

    openCertButton = ctk.CTkButton(
        mainScreen, text="Open Certificate", command=openCert,  fg_color="#061E40", border_color="white", border_width=2
    )


    exportCert = ctk.CTkButton(
        mainScreen, text="Export", command=exportCertificate,  fg_color="#061E40", border_color="white", border_width=2
    )


    openExcelButton = ctk.CTkButton(
        mainScreen, text="Open Empty Certificate & Errors", fg_color="#061E40", border_color="white", border_width=2, command=openExcel
    )
    openExcelButton.place(x=180, y=450)

def changetestJig(*args):
    
    if benchType.get() == "WECO2350":
        testJigType.configure(values=['TRIACTA-6312-JIG-XP-01', 'TRIACTA-6312-JIG-3P-01', 'TRIACTA-6320-JIG-3P-01', 'TRIACTA-GATEWAY-GT-XP-01'])					

    elif benchType.get() == "WECO4150":
        testJigType.configure(values=['TRIACTA-6312-JIG-XP-02', 'TRIACTA-6312-JIG-3P-02', 'TRIACTA-6320-JIG-XP-01'])

    else:
        testJigType.configure(values=[""])

def voltageSelection(*args):

    if testJigType.get() == "TRIACTA-6320-JIG-3P-01" or testJigType.get() == "TRIACTA-6320-JIG-XP-01":
        voltageType.configure(values=["120"])

    elif testJigType.get() == "TRIACTA-6312-JIG-XP-01" or testJigType.get() == "TRIACTA-6312-JIG-3P-01"  or testJigType.get() == 'TRIACTA-GATEWAY-GT-XP-01':
        voltageType.configure(values=["120", "347", "600"])

    elif testJigType.get() == "TRIACTA-6312-JIG-XP-02" or testJigType.get() == "TRIACTA-6312-JIG-3P-02":
        voltageType.configure(values=["120", "240", "277", "347", "480", "600"])

    else:
        voltageType.configure(values=[""])


def submitData():

    global dataFilePath

    dataFilePath = filedialog.askopenfilename()

    if dataFilePath and dataFilePath.endswith(".xlsx"):
        dataType.configure(
            text="Data File Selected", border_color="green", font=("Calibri", 15, "bold")
        )

    else:
        dataType.configure(text="Wrong File Type (.xlsx)", border_color="red")


def genJigCert():

    tempSelection = tempType.get()
    benchSelection = benchType.get()
    testJigSelection = testJigType.get()
    voltageSelect = voltageType.get()

    if testJigSelection == "TRIACTA-6320-JIG-XP-01" or testJigSelection == "TRIACTA-6320-JIG-3P-01":
        testJigSelection = "6320"

    elif testJigSelection == "TRIACTA-6312-JIG-3P-01" or testJigSelection == "TRIACTA-6312-JIG-3P-02":
        testJigSelection = "6312-3P"

    elif testJigSelection == "TRIACTA-6312-JIG-XP-01" or testJigSelection == "TRIACTA-6312-JIG-XP-02":
        testJigSelection = "6312-XP"

    elif testJigSelection ==  'TRIACTA-GATEWAY-GT-XP-01':
        testJigSelection = "GT-XP-01"


    if (tempSelection != "" and dataFilePath != "" and benchSelection in ["WECO2350", "WECO4150"] 
    and voltageSelect in ["120", "277", "347", "480", "600"] and testJigSelection in ["6312-3P", "6312-XP", "6320"]
    ):
        generateButton.place_forget()
        generateTestJigCert(tempSelection, dataFilePath, benchSelection, voltageSelect, testJigSelection)
        openCertButton.place(x=200, y=330)

    else:
        generateButton.configure(text="Error Occured (Check Data)", border_color= "red")


def openCert():
    """
        Allows user to see the certificate that was just created
    """
    filePath = r"weeklyCheck\excelFile\Modified.xlsx"

    if os.path.exists(filePath):
        os.startfile(filePath)
        print("Opened certificate")
        exportCert.place(x=200, y=370)
        
    
    else:
        print("Unable to open certificate")


def exportCertificate():

    meterSelection = meterPointType.get()
    exportJigCert(meterSelection)
    print("Export Successful")
    refreshjigUI()


def refreshjigUI():
    benchType.set("")
    testJigType.set("")
    voltageType.set("")
    dataType.configure(border_color="white", border_width=2, font=("Calibri", 15, "normal"))
    dataType.place(x=200, y=330)
    tempType.delete(0, ctk.END)
    generateButton.place(x=200, y=380)
    openCertButton.place_forget()
    exportCert.place_forget()
    meterPointType.delete(0, ctk.END)


def openExcel():
    os.startfile(r"weeklyCheck\excelFile\TestJigCert.xlsx")

    
    