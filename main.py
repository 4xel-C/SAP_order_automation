import win32com.client
import subprocess
import time

# Initialize objects
application = None
connection = None
session = None

# Check if SAP is already open
try:
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
except:
    SapGuiAuto = None

if not SapGuiAuto:
    subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
    while not SapGuiAuto:
        try: 
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        except:
            SapGuiAuto = None
# Vérifier si l'objet application existe, sinon le créer
try:
    application = SapGuiAuto.GetScriptingEngine
except Exception as e:
    print(f"Erreur lors de l'accès à SAP GUI : {e}")

# Vérifier si l'objet connection existe, sinon le créer
if application is not None:
    try:
        connection = application.OpenConnection('010 SAP R/3 Production (PBC)', True)
    except Exception as e:
        print(f"Erreur lors de l'accès à la connexion : {e}")

# Vérifier si l'objet session existe, sinon le créer
if connection is not None:
    try:
        session = connection.Children(0)
    except Exception as e:
        print(f"Erreur lors de l'accès à la session : {e}")

# Access to order menu
session.findById("wnd[0]").Maximize()

# Enter the transaction code and execute
session.findById("wnd[0]/tbar[0]/okcd").Text = "MB21"
session.findById("wnd[0]").sendVKey(0)
session.findById("wnd[0]/usr/ctxtRM07M-BWART").text = "201"
session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = "PFRE"

session.findById("wnd[0]/usr/ctxtRM07M-WERKS").setFocus()
session.findById("wnd[0]/usr/ctxtRM07M-WERKS").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[7]").press()
session.findById("wnd[0]/usr/txtRKPF-WEMPF").text = "GFEEU_D1-368"

# Table line 1
session.findById("wnd[0]/usr/subBLOCK:SAPLKACB:1001/ctxtCOBL-KOSTL").text = "PF04121100"
session.findById("wnd[0]/usr/sub:SAPMM07R:0521/ctxtRESB-MATNR[0,7]").text = "test"
session.findById("wnd[0]/usr/sub:SAPMM07R:0521/txtRESB-ERFMG[0,26]").text = "1"
session.findById("wnd[0]/usr/sub:SAPMM07R:0521/ctxtRESB-LGORT[0,53]").text = "RE01"

# table line 2
session.findById("wnd[0]/usr/sub:SAPMM07R:0521/ctxtRESB-MATNR[1,7]").text = "test2"
session.findById("wnd[0]/usr/sub:SAPMM07R:0521/txtRESB-ERFMG[1,26]").text = "2"
session.findById("wnd[0]/usr/sub:SAPMM07R:0521/ctxtRESB-LGORT[1,53]").text = "RE01"

