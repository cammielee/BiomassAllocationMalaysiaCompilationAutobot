import os
from pathlib import Path

class directory(object):
    def __init__(self):
        # self.main_dir = os.getcwd() #! change this

        self.path = Path(r"C:\TMO_SharedFolder\Projects\BiomassAllocationMalaysiaCompilationAutobot")
        
        self.referenceDir = self.path /  "User"
        # -- Logging directory ---
        self.log_dir = os.path.join(self.path, "Logs")
        
        # --- report directory------
        self.report_dir = self.path / "User" /"Reports" 
   

        # ---raw directory ------
        self.raw_dir = self.path / "User" /"Allocation Malaysia"

        #-----Lookup tables, maintained by end users-------
        self.lookup_fileDir = self.path / "User" /"Lookup File" 
        # self.lookup_dir = os.path.join(self.user_dir, 'Lookup File')
        
    

print(os.getcwd())
