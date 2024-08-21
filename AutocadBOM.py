import xlrd


class AutocadBOM():
    def __init__(self, partNumber, filename):
        self.partNumber = partNumber.upper()
        self.filename = filename
        self.description = "Bill of Materials"
        
        self.partClass = "ELEC" if self.partNumber[5] == "9" else "MECH"
        self.partType = "A"
        
        
    def formatCSV(self):
        wb = xlrd.open_workbook(self.filename)
        ws = wb.sheet_by_index(0)
        
        partsList = []
        
        for row in range(1, ws.nrows):
            lineNumber = str(ws.cell_value(row, 0)).strip()
            partNumber = str(ws.cell_value(row, 3)).strip()
            referenced = (ws.cell_value(row, 4) != "")
            qty = str(ws.cell_value(row, 5)).strip()
            
            print(partNumber)
            
            if lineNumber.isnumeric() and not referenced and partNumber != "":
                partsList.append([self.partNumber, lineNumber, partNumber, qty])
        
        BOMImportPath = "\\\\Erp\\dbamfg\\IMPORT\\BOM1.csv"
        
        with open(BOMImportPath, "w") as f:
            f.write("Parent Part Code,	Line Number,	Component Part Code,	Quantity Required\n")
            f.write("DO NOT REARRANGE THE ORDER OF THE COLUMNS IN THE FILE\n")
            
            for part in partsList:
                f.write(",".join(part) + "\n")
            