import time
class AutoCADManager:
    import win32com.client as wc
    import os
    def __init__(self) -> None:
        self.TARGETDRAWINGFILE = None
        self.docObj = None
        self.SCRIPTOUTPUTPATH = None
        self.IfVersion2022 = False
    def autocad(self, command: str):
        if not command:
            return [False, "Please provide a command to run"]
        if self.TARGETDRAWINGFILE:
            if self.docObj:
                if self.SCRIPTOUTPUTPATH:
                    while True:
                        try:
                            __file = AutoCADManager.os.path.join(self.SCRIPTOUTPUTPATH, str(self.docObj.Name).replace(".dwg", ".lsp"))
                            break
                        except Exception:
                            continue
                    scriptOut = open(__file, "a")
                    scriptOut.write(command+"\n")
                    scriptOut.close()
                self.docObj.SendCommand(command)
            else:
                acad = AutoCADManager.wc.Dispatch(f"AutoCAD.Application")
                fullName = acad.FullName
                if fullName.find("2022") != -1:
                    self.IfVersion2022 = True
                acad.Visible = True
                self.docObj = acad.Documents.Open(self.TARGETDRAWINGFILE)
                if self.SCRIPTOUTPUTPATH:
                    while True:
                        try:
                            __file = AutoCADManager.os.path.join(self.SCRIPTOUTPUTPATH, str(self.docObj.Name).replace(".dwg", ".lsp"))
                            break
                        except Exception:
                            continue
                    scriptOut = open(__file, "a")
                    scriptOut.write(command+"\n")
                    scriptOut.close()
                return self.docObj.SendCommand(command)
        else:
            acad = AutoCADManager.wc.GetActiveObject("AutoCAD.Application")
            fullName = acad.FullName
            if fullName.find("2022") != -1:
                self.IfVersion2022 = True
            self.docObj = acad.ActiveDocument
            if self.SCRIPTOUTPUTPATH:
                    while True:
                        try:
                            __file = AutoCADManager.os.path.join(self.SCRIPTOUTPUTPATH, str(self.docObj.Name).replace(".dwg", ".lsp"))
                            break
                        except Exception:
                            continue
                    scriptOut = open(__file, "a")
                    scriptOut.write(command+"\n")
                    scriptOut.close()
            return self.docObj.SendCommand(command)
    def createLayer(self, layerName):
        if not layerName:
            return "You must provide a valid layer name to create."
        self.autocad(f'(command "LAYER" "NEW" "{layerName}" "") ')
        return True
    def renameLayer(self, oldLayerName, newLayerName):
        if not oldLayerName or not newLayerName:
            return "You must provide both old and new layer names to rename."
        self.autocad(f'(command "LAYER" "RENAME" "{oldLayerName}" "{newLayerName}" "") ')
        return True
    def lineTypeOfLayer(self, layerName, lineType):
        if not layerName or not lineType:
            return "You must provide both layer name and line type to set."
        self.autocad(f'(command "LAYER" "LTYPE" "{layerName}" "{lineType}" "") ')
        return True
    def colorOfLayer(self, layerName, indexColor=None, trueColor=None, colorBookAndColor=(None, None)):
        if not layerName:
            return "You must provide a valid layer name to set color."
        if indexColor:
            self.autocad(f'(command "LAYER" "COLOR" " "{indexColor}" "{layerName}" "") ')
        elif trueColor:
            self.autocad(f'(command "LAYER" "COLOR" "TRUECOLOR" "{trueColor}" "{layerName}" "") ')
        elif colorBookAndColor:
            if not len(colorBookAndColor) == 2:
                return "Provide a valid color book name and color index."
            colorBook, colorName = colorBookAndColor
            self.autocad(f'(command "LAYER" "COLOR" "COLORBOOK" "{colorBook}" "{colorName}" "{layerName}" "") ')
        else:
            return "Please provide either index color, true color or color book and color."
        return True
    def setCurrentLayer(self, layerName):
        if not layerName:
            return "You must provide a valid layer name to make it as current layer."
        self.autocad(f'(command "LAYER" "M" "{layerName}" "" ) ')
    def straightLineByLength(self, start_coordinate, lengthOfLine):
        try:
            x1=start_coordinate[0]
            y1=start_coordinate[1]
        except:
            return "Please provide both the x and y coordinates in list or tuple for start coordinate."
        if str(type(x1)) not in ["<class 'int'>", "<class 'float'>"]:
            return "Coordinate values must be a float or integer"
        if str(type(lengthOfLine)) not in ["<class 'int'>", "<class 'float'>"]:
            return "Lenght of line must be a float or an integer"
        x2=x1+lengthOfLine
        y2=y1
        self.autocad(f'(command "PLINE" "{x1},{y1}" "{x2},{y2}" "") ')
        return True
    def addText(self, coordinate, textValue, textHeight: float = 2.5, textAngle: float = 0):
        try:
            x1=coordinate[0]
            y1=coordinate[1]
        except:
            return "Please provide both the x and y coordinates in list or tuple for start coordinate."
        if str(type(x1)) not in ["<class 'int'>", "<class 'float'>"]:
            return "Coordinate values must be a float or integer"
        if not textValue:
            return "Please provide a valid text value to add."
        if self.IfVersion2022:
            self.autocad(f'(command "TEXT" "{x1},{y1}" "{textAngle}" "{textValue}") ')
        else:
            self.autocad(f'(command "TEXT" "{x1},{y1}" "{textHeight}" "{textAngle}" "{textValue}") ')
        return True
    def addMtext(self, coordinate, textValue):
        try:
            x1=coordinate[0]
            y1=coordinate[1]
        except:
            return "Please provide both the x and y coordinates in list or tuple for start coordinate."
        if str(type(x1)) not in ["<class 'int'>", "<class 'float'>"]:
            return "Coordinate values must be a float or integer"
        if not textValue:
            return "Please provide a valid text value to add."
        self.autocad(f'(command "MTEXT" "{x1},{y1}" "{textValue}" "") ')
        return True
    def createLayer(self, layerName):
        if not layerName:
            return "You must provide a valid layer name to create."
        self.autocad(f'(command "LAYER" "NEW" "{layerName}" "" ) ')
        return True
    def exportLayers(self, folderPath):
        filesBeforeExtraction = AutoCADManager.os.listdir(folderPath)
        outFolder=folderPath.replace("\\", "\\\\")
        self.autocad(f'(command "ADEPTREPORTLAYERS" "{outFolder}") ')
        filesAfterExtraction = AutoCADManager.os.listdir(folderPath)
        newFiles = list(filter(lambda x: x is not None, map(lambda y: y if y not in filesBeforeExtraction else None, filesAfterExtraction)))
        newFileIndex=0
        if len(newFiles) >0:
            for f in newFiles:
                if f.find(".") != -1:
                    if f.split(".")[-1] == "csv":
                        break
                newFileIndex +=1
        else:
            return [False, "Something went wrong"]
        csvFileName = newFiles[newFileIndex]
        return [True, AutoCADManager.os.path.join(folderPath,csvFileName)]
    def getLayerNames(self):
        listVariable = []
        while True:
            try:
                if not self.docObj:
                    if not self.TARGETDRAWINGFILE:
                        acad = AutoCADManager.wc.Dispatch("AutoCAD.Application")
                        self.docObj = acad.ActiveDocument
                for i in self.docObj.Layers:
                    if i.Name not in listVariable:
                        listVariable.append(i.Name)
                break
            except:
                continue
        return listVariable
    def changeAllTextObjSize(self, size: float):
        lispPath = AutoCADManager.os.path.join(AutoCADManager.os.environ["TEMP"], "templisp.lsp")
        lispcontent = f"""
(defun c:AutocadManagerLispChangeTextHeight (/ ss i ent height)
  (setq ss (ssget "X" '((0 . "TEXT")))) ; Select all text objects
  (if ss
    (progn
      (setq height {size}) ; Set the height to 250000
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (ssname ss i))
        (setq entData (entget ent))
        (setq entData (subst (cons 40 height) (assoc 40 entData) entData)) ; Change height
        (entmod entData)
        (setq i (1+ i))
      )
    )
  )
  (princ)
)
"""
        with open(lispPath, "w") as lispFile:
            lispFile.write(lispcontent)
            lispFile.close()
        if self.SCRIPTOUTPUTPATH:
            while True:
                try:
                    __file = AutoCADManager.os.path.join(self.SCRIPTOUTPUTPATH, str(self.docObj.Name).replace(".dwg", ".lsp"))
                    break
                except Exception:
                    continue
            scriptOut = open(__file, "a")
            scriptOut.write(lispcontent+"\n")
            scriptOut.close()
        self.autocad(f'(load "{lispPath.replace("\\", "\\\\")}")\n')
        time.sleep(1)
        self.autocad('AUTOCADMANAGERLISPCHANGETEXTHEIGHT ')
        return True
    