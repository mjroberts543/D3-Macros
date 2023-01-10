from pynput import keyboard as kb
from pynput import mouse as ms
import time
import configparser
from ast import literal_eval
from pathlib import Path
import win32gui
import pygetwindow


try:
    d3=pygetwindow.getWindowsWithTitle("Diablo III")[0]._hWnd
    xres=pygetwindow.getWindowsWithTitle("Diablo III")[0].size[0]
    i_desktop_window_dc = win32gui.GetWindowDC(d3)
    resolution=(xres-1920)//2
except:
    pass
here = Path(__file__).parent.resolve()
builds = Path(__file__).parent.joinpath('builds').resolve()

config=configparser.ConfigParser()
config.read(str(here / 'config.ini'))


build=configparser.ConfigParser()  
sendkey=kb.Controller()

def presskey(keystr):
    if len(keystr)>1:
        exec("sendkey.press(kb.Key.{0})".format(keystr))
    else:
        sendkey.press(keystr)

def releasekey(keystr):
    if len(keystr)>1:
        exec("sendkey.release(kb.Key.{0})".format(keystr))
    else:
        sendkey.release(keystr)

def runskill(numb):
    global config
    global sendkey

    build=configparser.ConfigParser()
    build.read(str(here / 'build.ini'))
    if build[numb]['activation']=='off':
        return False
    
    color=literal_eval(build[numb]['color'])
    
    x=int(config[numb]['x'])+resolution
    
    y=int(config[numb]['y'])
    
    gbx=int(config[numb]['gbx'])+resolution
    
    gby=int(config[numb]['gby'])
    
    key=config[numb]['key']
    
    timewait=float(build[numb]['time'])
 
    if build[numb]['activation']=='cd':
        while 1:
            if get_pixel_color(x,y)==color:
                time.sleep(.15)
                presskey(key)
                releasekey(key)
                time.sleep(.5)
            time.sleep(.01)

    elif build[numb]['activation']=='gb':
        while 1:
            testcol=sum(get_pixel_color(gbx,gby))
            if get_pixel_color(x,y)==color and testcol<50:
                time.sleep(.15)
                presskey(key)
                releasekey(key)
                time.sleep(.5)
            time.sleep(.01)

    elif build[numb]['activation']=='timed':
        while 1:
            presskey(key)
            releasekey(key)
            time.sleep(timewait)

def get_pixel_color(i_x, i_y):
    global i_desktop_window_dc
    long_colour = win32gui.GetPixel(i_desktop_window_dc, i_x, i_y)
    i_colour = int(long_colour)
    
    return (i_colour & 0xff), ((i_colour >> 8) & 0xff), ((i_colour >> 16) & 0xff)



if __name__=="__main__":
    from ctypes import wintypes
    from elevate import elevate
    import win32api
    import multiprocessing
    from multiprocessing import Pool
    import tkinter
    from playsound import playsound
    import subprocess
    import ctypes
    
    #make print stop working
    elevate(show_console=False)

    res = Path(__file__).parent.joinpath('res').resolve()
    try:
        thon=len(config['other']['turbohud'])!=1
        togglekey=config['customhotkeys']['toggle']
        gearswapkey=config['customhotkeys']['gearswap']
        recordkey=config['customhotkeys']['record']
        gemtownkey=config['customhotkeys']['gemtown']
        bloodshardkey=config['customhotkeys']['bloodshards']
        salvagekey=config['customhotkeys']['salvage']
        upgraderareskey=config['customhotkeys']['upgraderares']
        inihotkeys=[togglekey,gearswapkey,recordkey,gemtownkey,bloodshardkey,salvagekey,upgraderareskey]
    except:
        pass
    hotkeytext="\n\n{0}:Toggle automatic skills.        \n{1}:Gear swap top left of inventory.\n{2}:Record skill colors.            \n{3} :Upgrade gems and tp.            \n{4} :Spend bloodshards.              \n{5} :Salvage all.                    \n{6} :Upgrade inventory of rares.     \n\n".format(*inihotkeys)
    recordtext=" To record skill colors, press the\nfollowing keys on the numpad while\nyour skills are not on cooldown,\nin order for skills 1 to 6.\n\n 0 :Skill will not cast automatically.  \n 1 :Skill casts when off cooldown.      \n 2: Skill casts near its effect's end.  \n 3 :Skill casts every (below) secs.     \n 9 :Skip recording the skill.           \n f3:Reset skill recording.              " 

    BlockInput = ctypes.windll.user32.BlockInput
    BlockInput.argtypes = [wintypes.BOOL]
    BlockInput.restype = wintypes.BOOL
    multiprocessing.freeze_support()
    if thon:
        try:
            subprocess.call("taskkill /F /IM  TurboHUD.exe",creationflags=0x08000000)
        except:
            pass
        
        try:
            subprocess.Popen(config['other']['turbohud'],stdout=subprocess.PIPE, creationflags=0x08000000)
        except:
            pass
    toggled=0   
    skillpool=Pool(6)
    mouse=ms.Controller()
    invpos=[[1529, 585], [1579, 585], [1630, 585], [1681, 585], [1731, 585], [1781, 585], [1832, 585], [1883, 585], [1529, 634], [1579, 634], [1630, 634], [1683, 632], [1733, 632], [1781, 634], [1834, 632], [1883, 634], [1529, 684], [1579, 684], [1630, 684], [1683, 682], [1733, 682], [1781, 684], [1834, 682], [1883, 684], [1529, 734], [1579, 734], [1630, 734], [1683, 732], [1733, 732], [1781, 734], [1834, 732], [1883, 734], [1529, 783], [1579, 783], [1630, 783], [1683, 781], [1733, 781], [1781, 783], [1834, 781], [1883, 783], [1531, 828], [1581, 828], [1632, 828], [1681, 832], [1731, 832], [1783, 828], [1832, 832], [1883, 832]]
    for xy in invpos:
        xy[0]=xy[0]+2*resolution
    
    if len(config['panelverification']['salvage'])==1:
        emptyinv=[(0,0,0) for xy in invpos]
    else:
        emptyinv=literal_eval(config['data']['emptyinv'])


    def savebuild():
        global builds
        global build
        global entry
        with open(str(builds)+"\\"+entry.get()+".ini",'w') as buildfile:
            build.write(buildfile)

    def loadbuild():
        global build
        global entry
        build=configparser.ConfigParser()
        build.read(str(builds)+"\\"+entry.get()+".ini")
        with open(str(here / 'build.ini'),'w') as buildfile:
            build.write(buildfile)
      
    i=1
    def readskills():
        global sendkey
        releasekey(recordkey)
        time.sleep(.5)
        global i
        i=1
        outputtext[0].set("\n \n \n \nRecording skill 1....")
        outputtext[2].set("Recording skill colors....\n")
        outputtext[1].set(recordtext)
        try:
            build.add_section(str(i))
        except:
            pass
        try:
            timerentrybig.set(float(build[str(i)]['time'])//1)
            timerentrysmall.set(float(build[str(i)]['time'])%1)
        except:
            build[str(i)]['time']="0"
            timerentrybig.set(float(build[str(i)]['time'])//1)
            timerentrysmall.set(float(build[str(i)]['time'])%1)

        def on_press(key):
            pass
                
        def on_release(key):
            global i
            global config
            global build

            if i<6:
                appendtext="\nRecording skill "+str(i+1)+"...."
            else:
                appendtext="\nRecording done."

            if str(key)=="<96>":
                build[str(i)]['activation']="off"
                build[str(i)]['color']=str(get_pixel_color(int(config[str(i)]['x']),int(config[str(i)]['y'])))
                outputtext[0].set("\n\nSkill "+str(i)+" casts only manually.\n"+appendtext)
                build[str(i)]['time']=str(timerentrybig.get()+timerentrysmall.get())
                i=i+1
            elif str(key)=="<97>":
                build[str(i)]['activation']="cd"
                build[str(i)]['color']=str(get_pixel_color(int(config[str(i)]['x']),int(config[str(i)]['y'])))
                outputtext[0].set("\n\nSkill "+str(i)+" casts when ready.\n"+appendtext)
                build[str(i)]['time']=str(timerentrybig.get()+timerentrysmall.get())
                i=i+1
            elif str(key)=="<98>":
                build[str(i)]['activation']="gb"
                build[str(i)]['color']=str(get_pixel_color(int(config[str(i)]['x']),int(config[str(i)]['y'])))
                outputtext[0].set("\n\nSkill "+str(i)+" casts as its effect ends.\n"+appendtext)
                build[str(i)]['time']=str(timerentrybig.get()+timerentrysmall.get())
                i=i+1
            elif str(key)=="<99>":
                build[str(i)]['activation']="timed"
                build[str(i)]['color']=str(get_pixel_color(int(config[str(i)]['x']),int(config[str(i)]['y'])))
                build[str(i)]['time']=str(timerentrybig.get()+timerentrysmall.get())
                outputtext[0].set("\n\nSkill "+str(i)+" casts every "+ str(float(build[str(i)]['time']))+ " seconds.\n"+appendtext)
                i=i+1
            elif str(key)=="<105>":
                outputtext[0].set("\n\nSkipping skill "+str(i)+".\n"+appendtext)
                i=i+1
            elif key==kb.Key.f1:
                toggle()
            elif key==kb.Key.f3:
                i=0
                outputtext[0].set("\n\nResetting to skill 1.\n\nRecording skill 1....")
                outputtext[1].set(recordtext)
                i=1
            
            if i==7:
                with open(str(here / 'build.ini'),'w') as buildfile:
                    build.write(buildfile)
                outputtext[2].set(" \n")

                return False
            try:
                build.add_section(str(i))
            except:
                pass
            try:
                timerentrybig.set(float(build[str(i)]['time'])//1)
                timerentrysmall.set(float(build[str(i)]['time'])%1)
            except:
                build[str(i)]['time']="0"
                timerentrybig.set(float(build[str(i)]['time'])//1)
                timerentrysmall.set(float(build[str(i)]['time'])%1)
        with kb.Listener(
        on_press=on_press,
        on_release=on_release) as listener:
            listener.join()
    
    def toggle():
        global sendkey
        global toggled
        global skillpool
        releasekey(togglekey)
        if toggled==1:
            skillpool.terminate()
            try:
                playsound(str(res / "off.mp3"),False)
            except:
                pass
            outputtext[1].set(hotkeytext)
            toggled=0
            outputtext[0].set("\n\nToggled off.\n\n " )

        elif toggled==0:
            
            skillpool=Pool(6)
            skillpool.imap_unordered(runskill,['1','2','3','4','5','6'])
            skillpool.close()
            try:
                playsound(str(res / "on.mp3"),False)
            except:
                pass
            toggled=1
            outputtext[1].set(hotkeytext)
            outputtext[0].set("\n\nToggled on.\n\n " )

    def nems():
        global toggled
        global sendkey
        global BlockInput
        releasekey(gearswapkey)
        if toggled==1:
            global mouse
            BlockInput(True)

            oldpos=mouse.position
            left=win32api.GetKeyState(0x01)
            right=win32api.GetKeyState(0x02)
            presskey("c")
            releasekey("c")
            mouse.release(ms.Button.left)
            mouse.release(ms.Button.right)
            
            mouse.position=(1430+2*resolution,600)
            mouse.click(ms.Button.right)
            time.sleep(.05)
            presskey("c")
            releasekey("c")
            
            mouse.position=oldpos
            time.sleep(.05)
            if left<0:
                mouse.press(ms.Button.left)
            if right<0:
                mouse.press(ms.Button.right)
            BlockInput(False)
            time.sleep(.4)


    def bloodshards():
        global sendkey
        global config
        releasekey(bloodshardkey)
        if len(config['panelverification']['bloodshards'])==1:
            config['panelverification']['bloodshards']=str(get_pixel_color(293, 62))
            with open(str(here)+"\\config.ini",'w') as configfile:
                config.write(configfile)

        if get_pixel_color(293, 62)==literal_eval(config['panelverification']['bloodshards']):
            global mouse
            mouse.click(ms.Button.right,24)

   
            


    def gemtown():
        global sendkey
        global config
        global toggled
        global mouse

        releasekey(gemtownkey)
        if len(config['panelverification']['gemtown'])==1:
            config['panelverification']['gemtown']=str(get_pixel_color(54, 838))

            BlockInput(True)
            mouse.position=(950,550)
            time.sleep(.2)
            
            i=5
            if toggled==1:
                toggle()   
            while i>0:
                mouse.position=(950,550)
                time.sleep(.2)
                if i==5:
                    config['data']['gemtown5']=str(get_pixel_color(320, 545))
                if i==4:
                    config['data']['gemtown4']=str(get_pixel_color(318, 547))
                if i==3:
                    config['data']['gemtown3']=str(get_pixel_color(301, 554))
                mouse.position=(250,550) 
                mouse.click(ms.Button.left,1)
                time.sleep(2)
                
                i=i-1

            presskey("t")
            releasekey("t")
            BlockInput(False)

            with open(str(here)+"\\config.ini",'w') as configfile:
                config.write(configfile)

        elif get_pixel_color(54, 838)==literal_eval(config['panelverification']['gemtown']):
            
            gemswait=0
            BlockInput(True)
            mouse.position=(950,550)
            time.sleep(.2)
            if get_pixel_color(320, 545)==literal_eval(config['data']['gemtown5']):
                gemswait=10
            elif get_pixel_color(318, 547)==literal_eval(config['data']['gemtown4']):
                gemswait=8
            elif get_pixel_color(301, 554)==literal_eval(config['data']['gemtown3']):
                gemswait=6
            if gemswait>0:
                mouse.position=(250,550) 
                i=0
                if toggled==1:
                    toggle()   
                while i<gemswait:
                    
                    mouse.position=(250,550) 
                    mouse.click(ms.Button.left,1)
                    
                    i=i+1
                    if gemswait-i==5:
                        presskey("t")
                        releasekey("t")
                    time.sleep(.05)
                    BlockInput(False)
                    time.sleep(.8)
                    BlockInput(True)
                    time.sleep(.05)
            BlockInput(False)

    def upgraderares():
        global sendkey
        global config
        releasekey(upgraderareskey)
        if len(config['panelverification']['upgraderares'])==1:
            config['panelverification']['upgraderares']=str(get_pixel_color(602, 305))
            with open(str(here)+"\\config.ini",'w') as configfile:
                config.write(configfile)

        if get_pixel_color(602, 305)==literal_eval(config['panelverification']['upgraderares']):
            global mouse
            BlockInput(True)
            currcols=[get_pixel_color(*xy) for xy in invpos]
            for i in range(len(invpos)):
                testcol=get_pixel_color(*invpos[i])
                
                if testcol!=emptyinv[i] and testcol==currcols[i]:
                    mouse.position=invpos[i]
                    time.sleep(.1)
                    mouse.click(ms.Button.right,1)
                    time.sleep(.1)
                    
                    mouse.position=(731,847)
                    mouse.click(ms.Button.left,1)
                    time.sleep(.1)
                    mouse.position=(253,836)
                    mouse.click(ms.Button.left,1)
                    time.sleep(.1)
                    mouse.position=(852,840)
                    mouse.click(ms.Button.left,1)
                    time.sleep(.1)
                    mouse.position=(582,840)
                    mouse.click(ms.Button.left,1)
                    time.sleep(.1)

                    
            mouse.position=(500,900)
            BlockInput(False)

    def salvageall():
        global sendkey
        global config
        global invpos
        global mouse
        releasekey(salvagekey)
        
        def clickyes():
            if len(config['panelverification']['clickyes'])==1:
                time.sleep(.5)
                last=[0,0,0,0,0]
                for i in range(5):
                    last[i]=get_pixel_color(984+resolution+i*3, 192)
                time.sleep(.5)
                diff=0
                for i in range(5):
                    if last[i]!=get_pixel_color(984+resolution+i*3, 192):
                        diff=diff+1
                if diff==0:
                    config['panelverification']['clickyes']=str(get_pixel_color(984+resolution, 192))
                    with open(str(here)+"\\config.ini",'w') as configfile:
                        config.write(configfile)
                    presskey("enter")
                    releasekey("enter")


            elif get_pixel_color(984+resolution, 192)==literal_eval(config['panelverification']['clickyes']):
                mouse.position=(852, 376)
                presskey("enter")
                releasekey("enter")

        if len(config['panelverification']['salvage'])==1:
            BlockInput(True)
            time.sleep(.01)
            
                    
            for i in range(4):
                mouse.position=(385-67*i, 295)
                mouse.click(ms.Button.left,1)
                time.sleep(.06)
                clickyes()
                
            for i in range(len(invpos)):
                j=(i+8*(i//8)+8*(i//24))%48
                mouse.position=invpos[j]
                mouse.click(ms.Button.left,1)
                mouse.position=(852, 376)
                time.sleep(.06)
                clickyes()
            mouse.position=(761, 642) 
            mouse.click(ms.Button.right)  
            time.sleep(1)
            for i in range(len(invpos)):
                emptyinv[i]=get_pixel_color(*invpos[i])
            config['data']['emptyinv']=str(emptyinv)
            config['panelverification']['salvage']=str(get_pixel_color(482, 470))
            with open(str(here)+"\\config.ini",'w') as configfile:
                config.write(configfile)

            BlockInput(False)

        elif get_pixel_color(482, 470)==literal_eval(config['panelverification']['salvage']):

            BlockInput(True)
            time.sleep(.01)
                    
            for i in range(4):
                mouse.position=(385-67*i, 295)
                mouse.click(ms.Button.left,1)
                time.sleep(.06)
                clickyes()
                
            for i in range(len(invpos)):
                j=(i+8*(i//8)+8*(i//24))%48
                testcol=get_pixel_color(*invpos[j])
                if testcol!=emptyinv[j]:
                    mouse.position=invpos[j]
                    mouse.click(ms.Button.left,1)
                    mouse.position=(852, 376)
                    time.sleep(.06)
                    clickyes()
            mouse.position=(761, 642) 
            mouse.click(ms.Button.right)  
            BlockInput(False)
        
    def closewindows():
        global sendkey
        releasekey("space")
        if len(config['panelverification']['closewindow'])==1:
            config['panelverification']['closewindow']=str(get_pixel_color(293, 62))
            with open(str(here)+"\\config.ini",'w') as configfile:
                config.write(configfile)

        if get_pixel_color(293, 62)==literal_eval(config['panelverification']['closewindow']):
            presskey("esc")
            releasekey("esc")
            presskey("space")
            releasekey("space")
            time.sleep(.5)

    ughangles=[]
    for i in range(len(inihotkeys)):
        if len(inihotkeys[i])==1:
            ughangles.append(inihotkeys[i])
        else:
            ughangles.append('<'+inihotkeys[i]+'>')
    h=kb.GlobalHotKeys({
        ughangles[0]: toggle,
        ughangles[1]: nems,
        ughangles[2]: readskills,
        ughangles[3]: gemtown,
        ughangles[4]: bloodshards,
        ughangles[5]: salvageall,
        ughangles[6]:upgraderares,
        '<space>': closewindows
        })
    h.start() 

    macrogui=tkinter.Tk()
    try:
        macrogui.iconbitmap(str(res / 'macros.ico'))
    except:
        pass
    macrogui.title("Mmmmacros")
    macrogui.resizable(0,0)
    savebutton=tkinter.Button(macrogui,text="Save",width=40,command=savebuild)
    loadbutton=tkinter.Button(macrogui,text="Load",width=40,command=loadbuild)
    entry = tkinter.Entry(macrogui,width=40)
    timerentrybig = tkinter.Scale(macrogui,resolution=1,orient=tkinter.HORIZONTAL, from_=0, to=120,sliderlength=10,length=242)
    timerentrysmall = tkinter.Scale(macrogui,resolution=.1,orient=tkinter.HORIZONTAL, from_=0, to=.9,sliderlength=10,length=100)
    savebutton.pack()
    entry.pack()
    loadbutton.pack()

    outputtext = tkinter.StringVar(),tkinter.StringVar(),tkinter.StringVar()
    label=[None,None,None]
    for i in [0,1,2]:
        label[i] = tkinter.Label(macrogui, text="", textvariable=outputtext[i], anchor="w", font="Courier 8")
        outputtext[1].set(hotkeytext)
        outputtext[0].set("\n\nWelcome to Mmmmacros!\n\n")
        outputtext[2].set("\n ")
    label[0].pack()
    label[1].pack()
    timerentrybig.pack()
    timerentrysmall.pack()
    label[2].pack()
    
    












    macrogui.mainloop()
    
    
    h.stop()


    if thon:
        try:
            subprocess.call("taskkill /F /IM  TurboHUD.exe",creationflags=0x08000000)
        except:
            pass
    try:
        subprocess.call("taskkill /F /IM  Macros.exe",creationflags=0x08000000)
    except:
        pass
    try:
        subprocess.call("taskkill /F /IM  Python.exe",creationflags=0x08000000)
    except:
        pass

