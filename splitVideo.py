import os
import tkinter
import warnings
def getSplitLength():
    global sl
    sl = -999
    def val(char):
        if str.isdigit(char) or char == "":
            return True
        else:
            return False
    top2 = tkinter.Tk()
    val2 = (top2.register(val))
    top2.title("SplitVideo")
    mainpanel3 = tkinter.PanedWindow(top2,orient=tkinter.HORIZONTAL)
    
    mainpanel3.pack(fill=tkinter.BOTH,expand = 1, side = tkinter.LEFT)
    label3 = tkinter.Label(mainpanel3, text="Enter split for file(s)"+" in seconds:     ",font=("Arial", 12))
    label3.pack(side = tkinter.LEFT)
    txt2 = tkinter.Entry(mainpanel3, validate='all',validatecommand=(val2, '%P'))
    txt2.config(width=5)
    txt2.pack(side = tkinter.LEFT)
    label4 = tkinter.Label(mainpanel3, text=" to ",font=("Arial", 12))
    label4.pack(side = tkinter.LEFT)
    txt3 = tkinter.Entry(mainpanel3, validate='all',validatecommand=(val2, '%P'))
    txt3.config(width=5)
    txt3.pack(side = tkinter.LEFT)
    def setSplitLength():
        global sl
        sl = str(txt2.get())+"-"+str(txt3.get())
        if(int(txt3.get())-int(txt2.get()) <= 0):
            warnings.warn("Split time cannot be 0 or negative")
            os._exit(0)
        
        print(sl)
        top2.destroy()
        return sl
    btn2 = tkinter.Button(mainpanel3, text ="GO", command=setSplitLength)
    btn2.pack(side = tkinter.BOTTOM)
    while(sl==-999):
        top2.update_idletasks()
        top2.update()
        #print(sl)

    return sl

def bye():
    os._exit(0)

def startSplit(splittimes):
    
    from moviepy.video.io.ffmpeg_tools import ffmpeg_extract_subclip
    print(splittimes)
    for i in splittimes:
        splits = splittimes[i].split("-")
        start = int(splits[0])
        end = int(splits[1])
        ffmpeg_extract_subclip(i, start, end, targetname="split_"+i.split("/")[-1])
    bye()

def oneFile():
    try:
        top1.destroy()
        from tkinter.filedialog import askopenfilename

        tkinter.Tk().withdraw() 
        filename = askopenfilename(filetypes=[("*.avi", "*.AVI")])
        print(filename)
        if(filename==""):
            os._exit(0)
        avifiles = []
        avifiles.append(filename)
        splittimes = {}
        
        x = getSplitLength()
        for i in avifiles: 
            splittimes[i] = x
        startSplit(splittimes)
    except FileNotFoundError:
        os._exit(0)


def multipleFile():
    try:
        top1.destroy()
        from tkinter.filedialog import askdirectory

        tkinter.Tk().withdraw() 
        folder = askdirectory()
        folder = folder+"/"
        print(folder)
        if(folder=="" or folder=="/"):
            os._exit(0)
        dirs = os.listdir(folder)
        #print(dirs)
        avifiles = []
        splittimes = {}
        for i in dirs:
            if(len(i)>3 and (i[-4:]==".avi" or i[-4:]==".AVI")):
                print(i)
                avifiles.append(i)
        print(avifiles)
        
        x = getSplitLength()
        for i in avifiles:
            splittimes[i] = x
            print(i)
        startSplit(splittimes)
    except FileNotFoundError:
        os._exit(0)



top1 = tkinter.Tk()
top1.title("SplitVideo")
mainpanel = tkinter.PanedWindow(orient=tkinter.VERTICAL)
mainpanel.pack(fill=tkinter.BOTH,expand = 1)
panel1 = tkinter.PanedWindow(mainpanel)
panel1.pack(fill=tkinter.BOTH, expand=1)
label6 = tkinter.Label(panel1, text="SplitVideo",justify=tkinter.CENTER,font=("Arial", 12))
label6.place(relx = 0.5, anchor=tkinter.CENTER)
label6.config(font=("Arial", 24))
panel1.add(label6)
btn1 = tkinter.Button(top1, text ="Split .avi File", command=oneFile)
btn1.config(width=50);


btn2 = tkinter.Button(top1, text ="Split .avi files in a Folder", command=multipleFile)
btn2.config(width=50);

btn3 = tkinter.Button(top1, text ="Exit", command=bye)
btn3.config(width=50);

btn3.pack(side = tkinter.BOTTOM)
btn2.pack(side = tkinter.BOTTOM)
btn1.pack(side = tkinter.BOTTOM)
top1.mainloop()
