from pathlib import Path
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import filedialog
from pathlib import Path
from moviepy.editor import *
import os
import ffmpeg
from pydub import AudioSegment
from docx2pdf import convert
from pdf2docx import Converter
from docx import Document
import re
import pypandoc
from tkinter import ttk


from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage


OUTPUT_PATH = Path(__file__).parent
current_dir = Path(__file__).parent
ASSETS_PATH = current_dir / "assets" / "images"


file_path = [None]
dr = [None]
window = tk.Tk()
value_inside = tk.StringVar(window)
value_inside.set("Convert to")

text_var = tk.StringVar()
text_var.set("Select file")

text_var1 = tk.StringVar()
text_var1.set("Save file")

default_list = ["mp3", "mp4", "png", "jpg", "wov", "pdf", "txt", "webm", "worddoc"]
video_list = ["only audio", "mp4", "webp"]  #all done
image_list = ["png", "jpg",] # done 
sound_list = ["mp3", "wav"] # done 
text_list = ["txt", "pdf", "docx"] #all done. 
combined_list = [video_list, image_list, sound_list, text_list]
selected_list = default_list


def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)




window.geometry("600x500")
window.configure(bg = "#1C1F1D")


canvas = Canvas(
    window,
    bg = "#1C1F1D",
    height = 500,
    width = 600,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)

canvas.place(x = 0, y = 0)
canvas.create_text(
    125.0,
    18.0,
    anchor="nw",
    text="Stumpy’s Converter ",
    fill="#FFFFFF",
    font=("Commissioner Regular", 36 * -1)
)

canvas.create_rectangle(
    49.0,
    69.0,
    550.0000007152539,
    70.00001376867203,
    fill="#FFFFFF",
    outline="")

canvas.create_rectangle(
    50.0,
    125.0,
    550.0,
    161.0,
    fill="#414441",
    outline="")

canvas.create_text(
    50.0,
    100.0,
    anchor="nw",
    text="Input File",
    fill="#FFFFFF",
    font=("Commissioner Regular", 14 * -1)
)

canvas.create_rectangle(
    512.0,
    129.0,
    513.0,
    155.0,
    fill="#000000",
    outline="")

canvas.create_rectangle(
    50.0,
    199.0,
    550.0,
    235.0,
    fill="#414441",
    outline="")

save_directory_text = canvas.create_text(
    61.0,
    208.0,
    anchor="nw",
    text="C:/Photo/",
    fill="#FFFFFF",
    font=("Inter", 16 * -1)
)

canvas.create_text(
    50.0,
    174.0,
    anchor="nw",
    text="Output  Location",
    fill="#FFFFFF",
    font=("Commissioner Regular", 14 * -1)
)

canvas.create_rectangle(
    512.0,
    203.0,
    513.0,
    229.0,
    fill="#000000",
    outline="")

canvas.create_rectangle(
    50.0,
    273.0,
    550.0,
    302.0,
    fill="#D9D9D9",
    outline="")

canvas.create_text(
    50.0,
    248.0,
    anchor="nw",
    text="Convert To",
    fill="#FFFFFF",
    font=("Commissioner Regular", 14 * -1)
)

canvas.create_rectangle(
    512.0,
    277.0,
    513.0,
    298.0,
    fill="#000000",
    outline="")

button_image_1 = PhotoImage(
    file=relative_to_assets("button_1.png"))
button_1 = Button(
    image=button_image_1,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: convert_file(file_path[0], value_inside.get(), dr[0]),
    relief="flat"
)
button_1.place(
    x=217.0,
    y=365.0,
    width=166.0,
    height=60.0
)


button_image_3 = PhotoImage(
    file=relative_to_assets("button_3.png"))
button_3 = Button(
    image=button_image_3,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: dr.__setitem__(0, save_file_directory()),
    relief="flat"
)
button_3.place(
    x=521.0,
    y=208.0,
    width=18.0,
    height=18.0
)

button_image_4 = PhotoImage(
    file=relative_to_assets("button_4.png"))
button_4 = Button(
    
    image=button_image_4,
    borderwidth=0,
    highlightthickness=0,
    command=lambda: (result := select_file(), file_path.__setitem__(0, result[0]), selected_list.__setitem__(0, result[1]), set_option_list(selected_list)),
    relief="flat"
)
button_4.place(
    x=520.0,
    y=133.0,
    width=20.0,
    height=20.0
)

input_directory_text = canvas.create_text(
    61.0,
    134.0,
    anchor="nw",
    text="C:/Example.png",
    fill="#FFFFFF",
    font=("Inter", 16 * -1)
)

def option_selected(event):
    # Clear text selection to prevent it from staying highlighted
    question_menu.selection_clear()
    # Remove focus from the combobox
    window.focus()


question_menu = ttk.Combobox(window, textvariable=value_inside, values=selected_list, state="readonly")

# Place the Combobox over the rectangle using create_window
canvas.create_window(300, 287.5, window=question_menu, width=500, height=30, anchor="center")

question_menu.bind("<<ComboboxSelected>>", option_selected)

style = ttk.Style()
style.configure("TCombobox",font=("Arial", 12),background="white",foreground="black")
style.map("TCombobox",fieldbackground=[("readonly", "lightgray")],foreground=[("readonly", "black")])



def set_option_list(selected_list):
    
    
     question_menu['values'] = selected_list[0]
    
    

def select_list(file_path):

    path = file_path
    name_ext = os.path.splitext(path)[1][1:]

    selected_list = []

    for lst_name in combined_list:
        if name_ext in lst_name:
            selected_list = lst_name
            break
    else:
       value_inside.set("No Supported Formats")
        
    
    

    return selected_list


    


def select_file():
    file_path = filedialog.askopenfilename(title="Select a File")
    if file_path:
        
        canvas.itemconfig(input_directory_text, text=f"{file_path}") 
    else:
        
        canvas.itemconfig(input_directory_text, text=f"No file selected") 

    file_list_select = select_list(file_path)

    
    return file_path, file_list_select
    
    

def to_docx(path, saveTo):
    document = Document()
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        file = f.read()
    file = re.sub(r'[^\x00-\x7F]+|\x0c',' ', file)
    document.add_paragraph(file)
    output = os.path.join(saveTo, os.path.splitext(os.path.basename(path))[0] + ".docx")
    return output, document




def convert_file(extention, convertTo, saveTo):
    print(extention) 
    print(convertTo)
    print(saveTo)
    #extention = path of input file
    #converTo = format to convert to
    #saveTo = save path
    
    

    if not extention:
        print("No file selected to convert!")
        return

    if not saveTo:
        print("No save directory specified!")
        return

    path = extention #full path.
    name_ext = os.path.basename(path)#name of file with extention. eg .txt .png
    name, name_of_ext = os.path.splitext(name_ext) #name of the file,/ name of the extention with . eg ".txt"
    

   


    

    

    if convertTo == "only audio":
        try:
            video = VideoFileClip(path)  
            audio_path = os.path.join(saveTo, f"{name[0]}.mp3")
            video.audio.write_audiofile(audio_path)
        except Exception as e:
            print(f"Error converting video to mp3: {e}")
    elif convertTo == "webm" or convertTo == "mp4": #add progress bar to the progress of the conversion. 
        try:
            input = ffmpeg.input(path)
            print(f"{name[0]}.{convertTo[0]}")
            print(f"{convertTo[0]}")
            output = ffmpeg.output(input, os.path.join(saveTo, f"{name[0]}.{convertTo}"), format=f"{convertTo}")
            ffmpeg.run(output)
        except Exception as e:
            print(f"Error converting video to webm: {e}")
    elif convertTo == "mp3":
        if name_ext == "mp3":
            AudioSegment.from_mp3(file_path).export(saveTo, format="wav")
        elif name_ext == "wav":
            AudioSegment.from_wav(file_path).export(saveTo, format="mp3")
    elif convertTo == "pdf":
            if name_of_ext == ".docx":
       
              out_put = os.path.join(f"{saveTo}" + f"/{name}" + f".{convertTo}")
              convert(path, out_put)
            elif name_of_ext == ".txt":
                output, document = to_docx(path, saveTo)
                document.save(output)
                out_put = os.path.join(f"{saveTo}" + f"/{name}" + f".{convertTo}")
                pdf_path = os.path.join(f"{saveTo}" + f"/{name}" + ".docx")
                convert(pdf_path, out_put)
                os.remove(pdf_path)
                
    
    elif convertTo == "docx":

        
        if name_of_ext == ".pdf":
            cv = Converter(path)
            save_with_name = os.path.join(saveTo, f"{name}.docx")
            cv.convert(save_with_name, start=0, end=None)
            cv.close()
            
        elif name_of_ext == ".txt":
              
            
            output, document = to_docx(path, saveTo)
            document.save(output)

    elif convertTo == "txt":
        if name_of_ext == ".docx":
            output = pypandoc.convert_file(path, "plain", outputfile= os.path.join(f"{saveTo}" + f"/{name}" + ".txt"))

        elif name_of_ext == ".pdf":
            cv = Converter(path)
            save_with_name = os.path.join(saveTo, f"{name}.docx")
            cv.convert(save_with_name, start=0, end=None)
            cv.close()
            new_path = os.path.join(f"{saveTo}" + f"/{name}" + ".docx")
            output = pypandoc.convert_file(new_path, "plain", outputfile= os.path.join(f"{saveTo}" + f"/{name}" + ".txt"))
            os.remove(new_path)




    else:
        print(f"Conversion to {convertTo} is not supported yet.")
        

        
def save_file_directory():
    dr = filedialog.askdirectory(title="select a file")
    if dr:
        
        canvas.itemconfig(save_directory_text, text=dr)   
    else:
        
        canvas.itemconfig(save_directory_text, text="No folder selected")

    return dr



text_var = tk.StringVar()
text_var.set("Select file")

text_var1 = tk.StringVar()
text_var1.set("Save file")




window.resizable(False, False)
window.mainloop()
