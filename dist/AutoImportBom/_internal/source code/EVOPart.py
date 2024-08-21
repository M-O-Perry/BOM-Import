from FindVendor import findVendor
import tkinter as tk
from EVOUtil import createNewPart


class Part():
    def __init__(self, partNumber, description, partClass = "", partType = "", mfg = "", mfgNumber = "", vendor = "", vendorNumber = "", specs = ""):
        self.partNumber = partNumber.strip()
        self.description = description.strip()
        self.partClass = partClass.strip()
        self.partType = partType.strip()
        self.mfg = mfg.strip()
        self.mfgNumber = mfgNumber.strip()
        self.vendor = findVendor(vendor.strip()[:4])
        self.vendorNumber = vendorNumber.strip()
        
        
        
        self.specs = specs.strip()
        
        middle2Digits = partNumber[5:7]
        
        classTypes = {
            "1": "FINM",
            "2": "ASSL",
            "3": "MECH",
            "4": "MSCE",
            "5": "MSCE",
            "6": "MECH",
            "7": "MECH",
            "8": "MSCE",
            "9": "ELEC",
        }

        if self.mfg == "" and self.mfgNumber == "" and self.vendor == "" and self.vendorNumber == "":
            self.partType = "A"
        elif self.mfg != "" and self.mfgNumber != "" and self.vendor != "" and self.vendorNumber != "": 
            self.partType = "R"
            
        else:
                
            
            missingInfo = []
            if self.mfg == "":
                missingInfo.append("Manufacturer")
            if self.mfgNumber == "":
                missingInfo.append("Manufacturer Number")
            if self.vendor == "":
                missingInfo.append("Vendor")
            if self.vendorNumber == "":
                missingInfo.append("Vendor Number")
                
            raise_error_message = f"Error: Incomplete information for part creation, part: {partNumber}.\n Missing information: {", ".join(missingInfo)}"
            
            
            root = tk.Tk()
            root.withdraw()
            tk.messagebox.showerror("Error", raise_error_message)
            
            
            if self.vendor == "" and vendor != "":
                tk.messagebox.showerror("Error", "Vendor not found in vendor list.")
                
            
            
            root.destroy()
            
            self.partType = "R"
        
        if partClass == "":
            self.partClass = classTypes[middle2Digits[0]]
        
    def createNew(self):
        createNewPart(self.partNumber, self.description, self.partClass, self.partType, self.mfg, self.mfgNumber, self.vendor, self.vendorNumber, self.specs, closeINB=False)