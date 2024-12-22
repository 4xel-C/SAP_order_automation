from config import COST_CENTER, SERVER, USER 
import subprocess
import win32com.client

def create_connection(path: str):
    """
    Create connection on SAP, input a string: path of the local application, and return a session
    The fonction will first check is a session is open. If a connection is already made and a session open,
    it will return the session only if a session is on a page recognize by the rest of the code; 
    other wise, it will close the connection and start a fresh one.
    """
    # Initialize objects
    application = None
    connection = None
    session = None

    # Check if SAP is already open, if not, open SAP
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except Exception:
        SapGuiAuto = None
        subprocess.Popen(path)
        print("Opening local SAP application...")
        
        # Loop until SAP is open and SAPGUI is detected
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

    # Check if already connected and a session open
    if application.Children.Count > 0:
        connection = application.Children(0)
        session = connection.Children(0)
        
        # if a session is already open in an unknown page: close the sessions and start a connection again, otherwise, return the session
        if session.findById("wnd[0]").Text not in ["SAP Easy Access", "Create Reservation: Initial Screen", "Create Reservation: New Items"]:
            session.findById("wnd[0]").close()
            
            # if a pop-up appear
            if session.Children.Count > 1:
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        else:
            return session
            
    # Connection if not already connected
    try:
        connection = application.OpenConnection(
            SERVER, True
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

def order_product(session, cart: dict) -> None:
    """
    Check which SAP page is open or crawl through the pages to fill the ordering form.
    """
    # Enter create reservation menu and inject information
    session.findById("wnd[0]").maximize()

    # Check if on the right page before manipulation
    if session.findById("wnd[0]").Text == "SAP Easy Access":
        session.findById("wnd[0]/tbar[0]/okcd").Text = "MB21"
        session.findById("wnd[0]").sendVKey(0)
        
    # Check if on the right page before creating reservation
    if session.findById("wnd[0]").Text == "Create Reservation: Initial Screen":
        session.findById("wnd[0]/usr/ctxtRM07M-BWART").text = "201"
        session.findById("wnd[0]/usr/ctxtRM07M-WERKS").text = "PFRE"
        session.findById("wnd[0]/usr/ctxtRM07M-WERKS").setFocus()
        session.findById("wnd[0]/usr/ctxtRM07M-WERKS").caretPosition = 4
        session.findById("wnd[0]/tbar[1]/btn[7]").press()

    # Check if on the right page before creating the list of product to order
    if session.findById("wnd[0]").Text == "Create Reservation: New Items":
        session.findById("wnd[0]/usr/txtRKPF-WEMPF").text = USER
        session.findById("wnd[0]/usr/subBLOCK:SAPLKACB:1001/ctxtCOBL-KOSTL").text = COST_CENTER
        
        # Add each element of the cart in each line of SAP form
        for i, (item, qty) in enumerate(cart.items()):
            session.findById(f"wnd[0]/usr/sub:SAPMM07R:0521/ctxtRESB-MATNR[{i},7]").text = item
            session.findById(f"wnd[0]/usr/sub:SAPMM07R:0521/txtRESB-ERFMG[{i},26]").text = qty
            session.findById(f"wnd[0]/usr/sub:SAPMM07R:0521/ctxtRESB-LGORT[{i},53]").text = "RE01"

 
def confirm_transaction(session):
    # session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]").Close()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    try:
        # Obtenir l'objet SAP GUI
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine

        # Fermer toutes les sessions ouvertes
        for connection in application.Children:
            for session in connection.Children:
                session.findById("wnd[0]").Close()  # Ferme la session
                print("Session SAP fermée.")

        # Quitter l'application SAP
        application.Quit()
        print("Application SAP fermée.")

    except Exception as e:
        print(f"Erreur lors de la fermeture de l'application SAP : {e}")
    
    
