import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from AutoCADManager import AutoCADManager
import time, threading
import pandas as pd
tempUtilityFolder=AutoCADManager.os.path.join(AutoCADManager.os.environ['TEMP'], "LegendGeneratorUtility")
class LegendGeneratorApp(AutoCADManager):
    def __init__(self, root, GapBetweenUtilityLines: float = -6.339, UtilityLineLength: float = 30.0):
        AutoCADManager.__init__(self)
        self.GapBetweenUtilityLines = GapBetweenUtilityLines
        self.UtilityLineLength = UtilityLineLength
        self.root = root
        self.root.title("Legend Generator")
        self.root.geometry("280x150")
        self.unknowLayers=[]
        self.alreadyRunning = False
        self.runningPermission = True
        # Select Data Source for layer
        self.layer_var = tk.BooleanVar()
        self.layer_checkbox = tk.Checkbutton(root, text="Select Data Source for layer", variable=self.layer_var, command=self.toggle_layer_button)
        self.layer_checkbox.grid(row=0, column=0, sticky='w', padx=10, pady=5)

        self.layer_button = tk.Button(root, text="Browse", command=self.select_layer_file, state=tk.DISABLED)
        self.layer_button.grid(row=0, column=1, padx=10, pady=5)

        # Select drawing file
        self.drawing_var = tk.BooleanVar()
        self.drawing_checkbox = tk.Checkbutton(root, text="Select drawing file", variable=self.drawing_var, command=self.toggle_drawing_button)
        self.drawing_checkbox.grid(row=1, column=0, sticky='w', padx=10, pady=5)

        self.drawing_button = tk.Button(root, text="Browse", command=self.select_drawing_file, state=tk.DISABLED)
        self.drawing_button.grid(row=1, column=1, padx=10, pady=5)
        
        self.script_button = tk.Button(root, text="Browse", command=self.select_script_folder, state=tk.DISABLED)
        self.script_button.grid(row=2, column=1, padx=10, pady=5)

        # Select folder for script output
        self.script_var = tk.BooleanVar()
        self.script_checkbox = tk.Checkbutton(root, text="Select output folder for script", variable=self.script_var, command=self.toggle_script_button)
        self.script_checkbox.grid(row=2, column=0, padx=10, pady=5)

        # Launch Button
        self.launch_button = tk.Button(root, text="Launch", command=self.startLauncherInThread)
        self.launch_button.grid(row=3, column=0, columnspan=2, pady=20)
        self.root.protocol("WM_DELETE_WINDOW", self.onClose)
    def onClose(self):
        self.runningPermission = False
        self.root.destroy()
    def toggle_layer_button(self):
        if self.layer_var.get():
            self.layer_button.config(state=tk.NORMAL)
        else:
            self.layer_button.config(state=tk.DISABLED)
    def toggle_drawing_button(self):
        if self.drawing_var.get():
            self.drawing_button.config(state=tk.NORMAL)
        else:
            self.drawing_button.config(state=tk.DISABLED)
    def toggle_script_button(self):
        if self.script_var.get():
            self.script_button.config(state=tk.NORMAL)
        else:
            self.script_button.config(state=tk.DISABLED)
    def select_layer_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.layerFilePath=file_path
    def select_drawing_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("DWG files", "*.dwg")])
        if file_path:
            self.TARGETDRAWINGFILE=file_path
    def select_script_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.SCRIPTOUTPUTPATH=folder_path
    @staticmethod
    def createTempFolderForUtility():
        global tempUtilityFolder
        if not AutoCADManager.os.path.exists(tempUtilityFolder):
            AutoCADManager.os.makedirs(tempUtilityFolder)
    def startLauncherInThread(self):
        if self.alreadyRunning == True:
            messagebox.showinfo("Legend Generator", "Legend Generator is already running.")
            return
        self.alreadyRunning = True
        threading.Thread(target=self.launch).start()
    def launch(self):
        global tempUtilityFolder
        LegendGeneratorApp.createTempFolderForUtility()
        exporter = self.exportLayers(tempUtilityFolder)
        if exporter[0]:
            layerFile = exporter[1]
            exportedLayersDataFrame = pd.read_csv(layerFile)
            AutoCADManager.os.remove(layerFile)
            Layers = list(exportedLayersDataFrame["Name"])
            KeySourceDataFrame = pd.read_excel(self.layerFilePath)
            SourceLayers = list(KeySourceDataFrame["LayerNames"])
            SourceKeys = list(KeySourceDataFrame["Keys"])
            for t in ("0","US_Viewport","US_Site Boundary", "US_Drawing Sheet", "US_Base Mapping", "Defpoints", "Z-Zz_20_10_90-T_Text", "Z-Zz_20_20-D_Dimensions"):
                if t in Layers:
                    Layers.remove(t)
            Lx = 0 #keep this same always
            Ly = 0 #add GapBetweenUtilityLines to go down evenly
            Tx = self.UtilityLineLength #keep this same always
            Ty = 0 # add GapBetweenUtilityLines to go down evenly
            for layer in Layers:
                if self.runningPermission == True:
                    if layer in SourceLayers:
                        while True:
                            try:
                                layerIndex = SourceLayers.index(layer)
                                keyForThisLayer = SourceKeys[layerIndex]
                                self.setCurrentLayer(layer)
                                time.sleep(0.5)
                                while True:
                                    self.straightLineByLength((Lx,Ly), self.UtilityLineLength)
                                    break
                                time.sleep(0.8)
                                while True:
                                    self.addText((Tx, Ty), keyForThisLayer)
                                    break
                                time.sleep(0.8)
                                Ly += self.GapBetweenUtilityLines
                                Ty += self.GapBetweenUtilityLines
                                break
                            except:
                                # print("Waiting for connection")
                                time.sleep(1)
                                continue
                    else:
                        while True:
                            try:
                                keyForThisLayer="Not found in the source file!"
                                self.setCurrentLayer(layer)
                                time.sleep(0.5)
                                while True:
                                    self.straightLineByLength((Lx,Ly), self.UtilityLineLength)
                                    break
                                time.sleep(0.8)
                                while True:
                                    self.addText((Tx, Ty), keyForThisLayer)
                                    break
                                time.sleep(0.8)
                                Ly += self.GapBetweenUtilityLines
                                Ty += self.GapBetweenUtilityLines
                                break
                            except:
                                # print("Waiting for connection")
                                time.sleep(1)
                                continue
                        self.unknowLayers.append(layer)
                else:
                    messagebox.showwarning("Legend Generator", "Process stopped!")
                    break
            if len(self.unknowLayers) < len(Layers):
                self.changeAllTextObjSize(2.5)
        if len(self.unknowLayers) > 0:
            messagebox.showinfo("Legend Generator", "Unkown layers found: " + ", ".join(self.unknowLayers))
        messagebox.showinfo("Legend Generator", "Legend has been generated successfully.")
        self.alreadyRunning = False

if __name__ == "__main__":
    root = tk.Tk()
    legendGeneratorApp = LegendGeneratorApp(root)
    root.mainloop()