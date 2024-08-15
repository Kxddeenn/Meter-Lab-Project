from PIL import Image
import customtkinter as ctk
import tkinter as tk
from components.updateCompleted import refreshCompleted
from components.settingFunctions import openSealinglog


# Background color = #2B2B2B
# Metergy Colors = #061E40
# #00C3D5

def uiCompleted(mainScreen):
    

    #UI DESIGN

    add1 = ctk.CTkLabel(mainScreen, text="__________________________________________________", text_color="white", font=("Arial", 20))
    add1.place(x=0, y=410)

    add2 = ctk.CTkLabel(mainScreen, text="➤", text_color="white", font=("Arial", 20))
    add2.place(x=35, y=453)

    backgroundColor = ctk.CTkImage(
        light_image=Image.open(r"images\backgroundcolor.png"),
        dark_image=Image.open(r"images\backgroundcolor.png"),
        size=(600,430)
    )
    backgroundColorPic = ctk.CTkLabel(mainScreen, image=backgroundColor, text="")
    backgroundColorPic.place(x=0,y=-10)


    logo = ctk.CTkImage(
        light_image=Image.open(r"images\meterseals.png"),
        dark_image=Image.open(r"images\meterseals.png"),
        size=(350, 350),
    )
    imageLogo = ctk.CTkLabel(mainScreen, image=logo, text="")
    imageLogo.place(x=85, y=15)

    # Functionalities

    refreshButton = ctk.CTkButton(mainScreen, command=refreshCompleted, text="Refresh Data", width=20, height=30, fg_color="#061E40", border_color="white", border_width=2)
    refreshButton.place(x=210, y=380)

    openButton2 = ctk.CTkButton(mainScreen, text="Open", width=100, command=openSealinglog, fg_color="#061E40")
    openButton2.place(x=210, y=453)

    certTemp = ctk.CTkLabel(
        mainScreen,
        text="Open Sealing Log:",
        font=("Arial",14),
        text_color="white"
        
    )
    certTemp.place(x=70, y=453)
