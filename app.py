import customtkinter as ctk
from UI.main import uiMain
from UI.settings import uiSettings
from UI.completed import uiCompleted
from UI.jigtest import uiTestJig
from PIL import Image

ctk.set_appearance_mode("dark")
app = ctk.CTk()
app.title("Certificate Generation")
app.geometry("800x800")
app.resizable(False, False)
app.iconbitmap(r"UI\images\metergyicon.ico")

welcomeLabel = ctk.CTkLabel(app, text="Meter Lab Certificate Generation", font=('Calibri', 30, "bold"))
welcomeLabel.pack(pady=20)

# Create a tab view for main and settings screens
appTabs = ctk.CTkTabview( 
    app,
    width=600,
    height=600,
    corner_radius=30,
    segmented_button_fg_color="#2B2B2B",
    border_width=5,
    border_color="#2B2B2B",
    segmented_button_selected_color="white",
    segmented_button_selected_hover_color="white",
    segmented_button_unselected_color="lightgrey",
    text_color="black",
)

appTabs.pack()

mainScreen = appTabs.add("Main")
testJigScreen = appTabs.add("Weekly Jig")
settingsScreen = appTabs.add("Settings")
completedScreen = appTabs.add("Completed")

lineColor = ctk.CTkImage(
    light_image=Image.open(r"images\linecolor.png"),
    dark_image=Image.open(r"images\linecolor.png"),
    size=(595,3),
    )
lineColorPic = ctk.CTkLabel(app, image=lineColor, text="", bg_color="#2B2B2B")
lineColorPic.place(x=101,y=120)


# Call functions to populate main and settings screens
uiMain(mainScreen)
uiSettings(settingsScreen)
uiCompleted(completedScreen)
uiTestJig(testJigScreen)

app.mainloop()


