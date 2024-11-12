import subprocess
import win32com.client


def create_connection(path: str):
    """
    Create connection on SAP, input a string: path of the local application, and return a session
    """
    # Initialize objects
    application = None
    connection = None
    session = None

    # Check if SAP is already open
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except Exception:
        SapGuiAuto = None

    if not SapGuiAuto:
        subprocess.Popen(path)
        print("Opening local SAP application...")
        while not SapGuiAuto:
            try:
                SapGuiAuto = win32com.client.GetObject("SAPGUI")
            except Exception:
                SapGuiAuto = None

    # Engine creation
    try:
        application = SapGuiAuto.GetScriptingEngine
    except Exception as e:
        print(f"Erreur lors de l'accès à SAP GUI : {e}")

    # Connection to SAP server
    if application is not None:
        try:
            connection = application.OpenConnection(
                "010 SAP R/3 Production (PBC)", True
            )
        except Exception as e:
            print(f"Erreur lors de l'accès à la connexion : {e}")

    # create session and communicate with API
    if connection is not None:
        try:
            session = connection.Children(0)
        except Exception as e:
            print(f"Erreur lors de l'accès à la session : {e}")

    return session

def order_product(session, cart):
    # # Access to order menu
    # session.findById("wnd[0]").Maximize()

    # Enter create reservation menu and inject information
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").Text = "MB21"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtRM07M-BWART").text = "201"
    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = "PFRE"

    # validate informations
    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").setFocus()
    session.findById("wnd[0]/usr/ctxtRM07M-WERKS").caretPosition = 4
    session.findById("wnd[0]/tbar[1]/btn[7]").press()
    
    session.findById("wnd[0]/usr/txtRKPF-WEMPF").text = "GFEEU_D1-368"
    session.findById("wnd[0]/usr/subBLOCK:SAPLKACB:1001/ctxtCOBL-KOSTL").text = "PF04121100"
    
    # Add each element of the cart in each line of SAP form
    for i, (item, qty) in enumerate(cart.items()):
        session.findById(f"wnd[0]/usr/sub:SAPMM07R:0521/ctxtRESB-MATNR[{i},7]").text = item
        session.findById(f"wnd[0]/usr/sub:SAPMM07R:0521/txtRESB-ERFMG[{i},26]").text = qty
        session.findById(f"wnd[0]/usr/sub:SAPMM07R:0521/ctxtRESB-LGORT[{i},53]").text = "RE01"

def confirm_transaction(session):
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]").Close()
    
def close_SAP(application, connection, session):
    session.findById("wnd[0]").Close()
    connection.Close()
    application.Quit()