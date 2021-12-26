import threading
import os,sys
import subprocess
import shutil,winreg
from ffmpeg import probe
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import simpledialog

class mainDialog(simpledialog.Dialog):
    def buttonbox(self):
        global filepath
        self.geometry("550x300")
        self.attributes("-topmost",True)
        self.vcodecdict = {
            "copy":"copy",
            "libopenh264":"mp4",
            "h264":"mp4",
            "h264_nvenc":"mp4",
            "hevc_nvenc":"mp4",
            "flv1":"flv",
            "msvideo1":"avi",
            "vp9":"webm"}

        self.acodecdict = {
            "copy":"copy",
            "libmp3lame":"copy",
            "alac":"copy",
            "flac":"copy",
            "on2":"copy",
            "mp1":"copy",
            "mp2":"copy",
            "mp3":"copy",
            "opus":"webm",
            "pcm_s16le":"wav",
            "pcm_s24le":"wav",
            "vorbis":"ogg"}
        
        self.tanilist = ["k","m"]
        self.sizechange = tk.BooleanVar(self)
        self.sizechange.set(False)

        self.width = tk.IntVar(self)
        self.width.set(1980)

        self.height = tk.IntVar(self)
        self.height.set(1020)

        self.fpschange = tk.BooleanVar(self)
        self.fpschange.set(False)

        self.fps = tk.IntVar(self)
        self.fps.set(30)

        self.crfset = tk.BooleanVar(self)
        self.crfset.set(False)

        self.crf = tk.IntVar(self)
        self.crf.set(23)

        self.videochange = tk.BooleanVar(self)
        self.videochange.set(False)

        self.videob = tk.IntVar(self)
        self.videob.set(200)

        self.videou = tk.StringVar(self)
        self.videou.set("k")

        self.audiochange = tk.BooleanVar(self)
        self.audiochange.set(False)

        self.audiob = tk.IntVar(self)
        self.audiob.set(200)

        self.audiou = tk.StringVar(self)
        self.audiou.set("k")

        self.selectv = tk.StringVar(self)
        self.selectv.set("libopenh264")

        self.selecta = tk.StringVar(self)
        self.selecta.set("copy")

        self.openset = tk.BooleanVar(self)
        self.openset.set(False)
        box = tk.Frame(self)
        self.filepath = tk.StringVar(self.master)
        if self.master.title() == "":
            self.filepath.set("ファイルを選択してください")
        else:
            self.filepath.set(os.path.abspath(self.master.title()))
        self.title("動画変換")
    
        self.pathframe = tk.Frame(self)
        self.pathframe.pack(fill=tk.BOTH)
        self.searchbtn = ttk.Button(self.pathframe,text="参照",command=self.filechange)
        self.searchbtn.pack(fill=tk.X,side=tk.LEFT)
        self.pathlbl = tk.Label(self.pathframe,textvariable=self.filepath,font=("",15))
        self.pathlbl.pack(fill=tk.X,side=tk.LEFT)

        self.comboframe = tk.LabelFrame(self,text="動画コーデック\t音声コーデック")
        self.comboframe.pack(fill=tk.BOTH,expand=True)

        self.vcombo = ttk.Combobox(self.comboframe,state="readonly",values=list(self.vcodecdict.keys()),font=("",15),textvariable=self.selectv)
        self.vcombo.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)

        self.acombo = ttk.Combobox(self.comboframe,state="readonly",values=list(self.acodecdict.keys()),font=("",15),textvariable=self.selecta)
        self.acombo.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)

        self.sizeframe = tk.Frame(self)
        self.sizeframe.pack(fill=tk.BOTH,expand=True)

        self.wframe = tk.LabelFrame(self.sizeframe)
        self.wframe.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)

        self.wchange = tk.Checkbutton(self.wframe,variable=self.videochange,text="動画のビットレートを変更する(左:数値\t右:単位)")
        self.wchange.pack(fill=tk.X)

        self.wframet = tk.Frame(self.wframe)
        self.wframet.pack(fill=tk.BOTH,expand=True)
        self.wspin = tk.Spinbox(self.wframet,textvariable=self.videob,from_=1,to=100)
        self.wspin.pack(fill=tk.BOTH,side=tk.LEFT,expand=True)

        self.tanicom = ttk.Combobox(self.wframet,state="readonly",values=self.tanilist,font=("",15),textvariable=self.videou,width=2)
        self.tanicom.pack(fill=tk.Y,side=tk.LEFT)

        self.hframe = tk.LabelFrame(self.sizeframe)
        self.hframe.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)

        self.hchange = tk.Checkbutton(self.hframe,variable=self.audiochange,text="音声のビットレートを変更する(左:数値\t右:単位)")
        self.hchange.pack(fill=tk.X)

        self.hframet = tk.Frame(self.hframe)
        self.hframet.pack(fill=tk.BOTH,expand=True)
        self.hspin = tk.Spinbox(self.hframet,textvariable=self.audiob,from_=1,to=100)
        self.hspin.pack(fill=tk.BOTH,side=tk.LEFT,expand=True)

        self.tanicom = ttk.Combobox(self.hframet,state="readonly",values=self.tanilist,font=("",15),textvariable=self.audiou,width=2)
        self.tanicom.pack(fill=tk.Y,side=tk.LEFT)
        
        self.qualityframe = tk.Frame(self)
        self.qualityframe.pack(fill=tk.BOTH,expand=True)

        self.crfframe = tk.LabelFrame(self.qualityframe)
        self.crfframe.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)

        self.crfchange = tk.Checkbutton(self.crfframe,variable=self.crfset,text="crf固定する、高いほど品質が低くなります)")
        self.crfchange.pack(fill=tk.X)

        self.cspin = tk.Spinbox(self.crfframe,textvariable=self.crf,from_=1,to=51)
        self.cspin.pack(fill=tk.BOTH,side=tk.LEFT,expand=True)

        self.fpsframe = tk.LabelFrame(self.qualityframe)
        self.fpsframe.pack(fill=tk.BOTH,expand=True,side=tk.LEFT)

        self.fchange = tk.Checkbutton(self.fpsframe,variable=self.fpschange,text="動画のフレームレート(fps)を変更する         ")
        self.fchange.pack(fill=tk.X)

        self.fpspin = tk.Spinbox(self.fpsframe,textvariable=self.fps,from_=1,to=360)
        self.fpspin.pack(fill=tk.BOTH,side=tk.LEFT,expand=True)

        self.resolutionframe = tk.Frame(self)
        self.resolutionframe.pack(fill=tk.BOTH,expand=True)

        self.rlblframe = tk.Frame(self.resolutionframe)
        self.rlblframe.pack(fill=tk.BOTH,expand=True)        

        self.resolutionchange = tk.Checkbutton(self.rlblframe,variable=self.sizechange,text="解像度を変更する")
        self.resolutionchange.pack(fill=tk.X,side=tk.LEFT)

        self.resolutionlbl = tk.Label(self.rlblframe,font=("",15),text="左:横サイズ\t右:縦サイズ\t")
        self.resolutionlbl.pack(fill=tk.X,side=tk.RIGHT)

        self.spinframe = tk.Frame(self.resolutionframe)
        self.spinframe.pack(fill=tk.BOTH,expand=True)

        self.wspin = tk.Spinbox(self.spinframe,textvariable=self.width,from_=1,to=10000)
        self.wspin.pack(fill=tk.BOTH,side=tk.LEFT,expand=True)

        self.wspin = tk.Spinbox(self.spinframe,textvariable=self.height,from_=1,to=10000)
        self.wspin.pack(fill=tk.BOTH,side=tk.LEFT,expand=True)

        self.convertframe = tk.Frame(self)
        self.convertframe.pack(fill=tk.BOTH,expand=True)

        self.openchange = tk.Checkbutton(self.convertframe,variable=self.openset,text="処理終了後にファイルを開く")
        self.openchange.pack(side=tk.LEFT,fill=tk.BOTH)

        self.convertbtn = ttk.Button(self.convertframe,text="ファイルを変換する",command=self.convert)
        self.convertbtn.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)

        self.progress = ttk.Progressbar(self)
        self.progress.pack(fill=tk.BOTH,expand=True)

        self.protocol("WM_DELETE_WINDOW",self.on_closing)
    def on_closing(self):
        self.closeset = True
    def convertmain(self):
        try:
            self.closeset = False
            self.convertbtn.config(state=tk.DISABLED)
            if not self.is_int(self.width.get()):
                messagebox.showerror("エラー","横の解像度が無効です")
                self.convertbtn.config(state=tk.NORMAL)
                return
            if not self.is_int(self.height.get()):
                messagebox.showerror("エラー","縦の解像度が無効です")
                self.convertbtn.config(state=tk.NORMAL)
                return
            if not self.is_int(self.fps.get()):
                messagebox.showerror("エラー","fpsが無効です")
                self.convertbtn.config(state=tk.NORMAL)
                return
            if not self.is_int(self.crf.get()):
                messagebox.showerror("エラー","crf値が無効です")
                self.convertbtn.config(state=tk.NORMAL)
                return
            if not self.is_int(self.videob.get()):
                messagebox.showerror("エラー","動画のビットレートが無効です")
                self.convertbtn.config(state=tk.NORMAL)
                return
            if not self.is_int(self.audiob.get()):
                messagebox.showerror("エラー","音声のビットレートが無効です")
                self.convertbtn.config(state=tk.NORMAL)
                return
            if not int(self.crf.get()) >= 0:
                messagebox.showerror("エラー","crf値が無効です")
                self.convertbtn.config(state=tk.NORMAL)
                return
            if not int(self.fps.get()) >= 0:
                messagebox.showerror("エラー","fpsが無効です")
                self.convertbtn.config(state=tk.NORMAL)
                return
            if not int(self.videob.get()) >= 0:
                messagebox.showerror("エラー","動画のビットレートが無効です")
                self.convertbtn.config(state=tk.NORMAL)
                return
            if not int(self.audiob.get()) >= 0:
                messagebox.showerror("エラー","音声のビットレートが無効です")
                self.convertbtn.config(state=tk.NORMAL)
                return
            if not int(self.width.get()) >= 0:
                messagebox.showerror("エラー","横の解像度が無効です")
                self.convertbtn.config(state=tk.NORMAL)
                return
            if not int(self.height.get()) >= 0:
                messagebox.showerror("エラー","縦の解像度が無効です")
                self.convertbtn.config(state=tk.NORMAL)
                return
            if not os.path.exists(self.filepath.get()):
                messagebox.showerror("エラー","ファイルが存在しません")
                self.convertbtn.config(state=tk.NORMAL)
                return

            runcmd = [os.path.join(os.path.dirname(sys.executable),"ffmpeg.exe"),"-y","-i",self.filepath.get(),"-progress","-","-vcodec",self.selectv.get(),"-acodec",self.selecta.get()]
            if self.sizechange.get():
                runcmd.append("-s")
                runcmd.append(f"{str(self.width.get())}x{str(self.height.get())}")
            if self.fpschange.get():
                runcmd.append("-r")
                runcmd.append(str(self.fps.get()))
            if self.audiochange.get():
                runcmd.append("-b:a")
                runcmd.append(str(self.audiob.get()) + str(self.audiou.get()))
            if self.videochange.get():
                runcmd.append("-b:v")
                runcmd.append(str(self.videob.get()) + str(self.videou.get()))
            if self.crfset.get():
                runcmd.append("-crf")
                runcmd.append(str(self.crf.get()))
            
            outpath = self.filepath.get() + "_converted_" + os.path.splitext(os.path.basename(self.filepath.get()))[-1]
            if os.path.exists(outpath):
                if messagebox.askyesno("確認","出力先ファイルが存在します\n上書きしますか?"):
                    pass
                else:
                    self.convertbtn.config(state=tk.NORMAL)
                    return
            runcmd.append(outpath)
            process = fprogress(runcmd)
            for num in process.run():
                self.pid = str(num["pid"])
                pid.set(str(num["pid"]))
                self.progress.config(maximum=int(num["max"]*100))
                self.progress.config(value=int(num["now"]*100))
            if self.openset.get():
                self.explore(outpath)
            self.convertbtn.config(state=tk.NORMAL)
        except:
            import traceback
            if self.closeset:
                pass
            else:
                messagebox.showerror("エラー",traceback.format_exc())
                self.convertbtn.config(state=tk.NORMAL)
    def convert(self):
        self.convertth = threading.Thread(target=self.convertmain)
        self.convertth.setDaemon(True)
        self.convertth.start()
    def filechange(self):
        fTyp = [("動画ファイル","*.mp4"),("動画ファイル","*.mkv"),("動画ファイル","*.avi"),("動画ファイル","*.webm"),("すべてのファイル", "*")]
        if os.path.exists(self.filepath.get()):
            iDir = os.path.dirname(os.path.abspath(self.filepath.get()))
        else:
            iDir = os.path.abspath(os.path.dirname(__file__))
        file_name = filedialog.askopenfilename(filetypes=fTyp, initialdir=iDir)
        if file_name == "":
            pass
        else:
            self.filepath.set(file_name)
    def explore(self,path):
        FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')
        path = os.path.normpath(path)
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        if os.path.isdir(path):
            subprocess.run([FILEBROWSER_PATH, path],startupinfo=startupinfo)
        elif os.path.isfile(path):
            subprocess.run([FILEBROWSER_PATH, '/select,', path],startupinfo=startupinfo)
    def is_int(self,data):
        try:
            int(data)
            return True
        except:
            return False
class fprogress:
    def __init__(self, command, ffmpeg_loglevel="verbose"):
        self._command = command
        index_of_filepath = self._command.index("-i") + 1
        self._filepath = self._command[index_of_filepath]

        self._can_get_duration = True

        try:
            self._duration_secs = float(probe(self._filepath)["format"]["duration"])
        except Exception:
            self._can_get_duration = False

        self._ffmpeg_args = self._command + ["-loglevel", ffmpeg_loglevel]

        if self._can_get_duration:
            self._ffmpeg_args += ["-progress", "pipe:1", "-nostats"]

    def run(self):
        popen_args = [self._ffmpeg_args]
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        if self._can_get_duration:
            process = subprocess.Popen(
                self._ffmpeg_args,startupinfo=startupinfo, stdout=subprocess.PIPE,stderr=subprocess.DEVNULL,stdin=subprocess.DEVNULL
            )
            
            previous_seconds_processed = 0
        else:
            process = subprocess.Popen(self._ffmpeg_args,startupinfo=startupinfo,stdout=subprocess.PIPE,stderr=subprocess.DEVNULL,stdin=subprocess.DEVNULL)
        
        maxcount = self._duration_secs
        count = 0
        try:
            while process.poll() is None:
                if self._can_get_duration:
                    ffmpeg_output = process.stdout.readline().decode("cp932")
                    if "out_time_ms" in ffmpeg_output:
                        seconds_processed = int(ffmpeg_output.strip()[12:]) / 1_000_000
                        seconds_increase = seconds_processed - previous_seconds_processed
                        count += seconds_increase
                        previous_seconds_processed = seconds_processed
                    yield {"max":maxcount,"now":count,"pid":str(process.pid)}
        except Exception as e:
            process.kill()

root = tk.Tk()
root.overrideredirect(1)
pid = tk.StringVar(root)
root.wm_attributes("-transparentcolor", "snow")
root.attributes("-alpha",0)
ttk.Style().configure("TP.TFrame", background="snow")

def main(filepath):
    global root
    root.title(filepath)
    def showmain(root = root):
        mainDialog(root)
        try:
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            subprocess.run(["taskkill","/f","/t","/pid",str(pid.get())],startupinfo=startupinfo,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,stdin=subprocess.DEVNULL)
        except:
            pass
        root.destroy()
    root.after(1,showmain)
    root.mainloop()

if len(sys.argv) == 2:
    if sys.argv[1] == "--uninstall":
        if messagebox.askyesno("確認","レジストリキーを削除しますか?"):
            path = f'Software\\movie_converter_setting'
            extensions = [".mp4",".mkv",".avi",".webm"]
            try:
                winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER,path)
            except:
                pass
            for extension in extensions:
                path = f'Software\\Classes\\SystemFileAssociations\\{extension}\\Shell\\動画を変換する\\command'
                try:
                    winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER,path)
                except:
                    pass
            for extension in extensions:
                path = f'Software\\Classes\\SystemFileAssociations\\{extension}\\Shell\\動画を変換する'
                try:
                    winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER,path)
                except:
                    pass
        else:
            main(sys.argv[1])
    else:
        main(sys.argv[1])
else:
    path = f'Software\\movie_converter_setting'
    try:
        settingkey = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, access=winreg.KEY_WRITE)
        data, regtype = winreg.QueryValueEx(winreg.OpenKeyEx(winreg.HKEY_CURRENT_USER, path, access=winreg.KEY_READ), 'rightset')
    except:
        winreg.SetValueEx(settingkey, 'rightset', 0, winreg.REG_SZ,"0")
    extensions = [".mp4",".mkv",".avi",".webm"]
    
    if str(data) == "1":
        if messagebox.askyesno("確認","右クリックメニューから削除しますか?"):
            for extension in extensions:
                path = f'Software\\Classes\\SystemFileAssociations\\{extension}\\Shell\\動画を変換する\\command'
                try:
                    winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER,path)
                except:
                    pass
            for extension in extensions:
                path = f'Software\\Classes\\SystemFileAssociations\\{extension}\\Shell\\動画を変換する'
                try:
                    winreg.DeleteKeyEx(winreg.HKEY_CURRENT_USER,path)
                except:
                    pass
            winreg.SetValueEx(settingkey,'rightset', 0, winreg.REG_SZ,"0")
    else:
        if messagebox.askyesno("確認","右クリックメニューに追加しますか?"):
            pythonpath = os.path.join(os.path.dirname(os.path.abspath(__file__)),"pythonw.exe")
            mypath = os.path.abspath(__file__)
            exepath = f'{pythonpath} {mypath} "%V"'
            
            for extension in extensions:
                path = f'Software\\Classes\\SystemFileAssociations\\{extension}\\Shell\\動画を変換する\\command'
                key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, access=winreg.KEY_WRITE)
                winreg.SetValueEx(key, '', 0, winreg.REG_SZ,exepath)
                winreg.CloseKey(key)
                winreg.SetValueEx(settingkey,'rightset', 0, winreg.REG_SZ,"1")

    winreg.CloseKey(settingkey)
    main("")