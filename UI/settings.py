import customtkinter as ctk
from components.settingFunctions import openCertTemp, openMeterError, saveJigDate, setBench, openFeedback, openLoop
from UI.main import checkBench
from PIL import Image

def uiSettings(mainScreen):
    """
    Sets up the UI elements for the Settings tab.

    Parameters:
    - mainScreen: The main screen (CTkTabview tab) where UI elements will be placed.
    """

    def changeSave():
        """
        Function triggered when the Save button is clicked.
        Changes the appearance of the Save button temporarily,
        saves the current date and selected bench, and reverts
        the Save button appearance after 1 second.
        """
        
        color = saveButton.cget('fg_color')

        saveButton.configure(fg_color='green')

        saveJigDate(compareDate.get())
        setBench(benchEntry.get())
        checkBench()

        mainScreen.after(1000, lambda: saveButton.configure(fg_color=color))

    lineColor = ctk.CTkImage(
    light_image=Image.open(r"images\linecolor.png"),
    dark_image=Image.open(r"images\linecolor.png"),
    size=(575,2),
    )
    lineColorPic = ctk.CTkLabel(mainScreen, image=lineColor, text="", bg_color="#2B2B2B")
    lineColorPic.place(x=0,y=175)

    lineColor2 = ctk.CTkImage(
    light_image=Image.open(r"images\linecolor.png"),
    dark_image=Image.open(r"images\linecolor.png"),
    size=(575,2),
    )
    lineColorPic2 = ctk.CTkLabel(mainScreen, image=lineColor2, text="", bg_color="#2B2B2B")
    lineColorPic2.place(x=0,y=330)

    settingIcon = ctk.CTkImage(
    light_image=Image.open(r"images\settings.png"),
    dark_image=Image.open(r"images\settings.png"),
    size=(70,70),
    )
    settingPic = ctk.CTkLabel(mainScreen, image=settingIcon, text="", bg_color="#2B2B2B")
    settingPic.place(x=425,y=20)

    title = ctk.CTkLabel(
        mainScreen,
        text="Test Jig",
        font=("Arial",16, "bold"),
        text_color="white",
        
    )
    title.place(x=225, y=0)

    compareText = ctk.CTkLabel(
        mainScreen,
        text="Last Comparison Test:",
        font=("Arial",14),
        text_color="white"
        
    )
    compareText.place(x=10, y=45)

    compareDate = ctk.CTkEntry(mainScreen, width=142)  
    compareDate.place(x=170, y=45)

    benchText = ctk.CTkLabel(
        mainScreen,
        text="Bench:",
        font=("Arial",14),
        text_color="white"
        
    )
    benchText.place(x=10, y=85)
    
    benchEntry = ctk.CTkComboBox(mainScreen, values= ["WECO4150", "WECO2350"])
    benchEntry.place(x=170, y=85)


    saveButton = ctk.CTkButton(mainScreen, text="Save", width=125, command=changeSave, fg_color="#061E40")
    saveButton.place(x=200, y=140)

    title2 = ctk.CTkLabel(
        mainScreen,
        text="Update Files",
        font=("Arial",16, "bold"),
        text_color="white",
    )
    title2.place(x=210, y=210)

    certTemp = ctk.CTkLabel(
        mainScreen,
        text="Certificate Template:",
        font=("Arial",14),
        text_color="white"
        
    )
    certTemp.place(x=90, y=250)
    
    openButton1 = ctk.CTkButton(mainScreen, text="Open", width=115, command=openCertTemp, fg_color="#061E40")
    openButton1.place(x=100, y=290)

    certTemp = ctk.CTkLabel(
        mainScreen,
        text="Meter Errors File:",
        font=("Arial",14),
        text_color="white"
        
    )
    certTemp.place(x=310, y=250)
     
    openButton3 = ctk.CTkButton(mainScreen, text="Open", width=115, command= lambda: openMeterError(benchEntry.get()), fg_color="#061E40")
    openButton3.place(x=310, y=290)


    title3 = ctk.CTkLabel(
        mainScreen,
        text="Support",
        font=("Arial",16, "bold"),
        text_color="white",
    )
    title3.place(x=225, y=355)

    
    certTemp = ctk.CTkLabel(
        mainScreen,
        text="Feedback Forum:",
        font=("Arial",14),
        text_color="white"
        
    )
    certTemp.place(x=100, y=400)
    
    openButton4 = ctk.CTkButton(mainScreen, text="Open",  width=115, command=openFeedback, fg_color="#061E40")
    openButton4.place(x=100, y=435)


    looppage = ctk.CTkLabel(
        mainScreen,
        text="Documentation:",
        font=("Arial",14),
        text_color="white"
        
    )
    looppage.place(x=315, y=400)
    
    loopbutton = ctk.CTkButton(mainScreen, text="Open", width=115, command=openLoop, fg_color="#061E40")
    loopbutton.place(x=315, y=435)