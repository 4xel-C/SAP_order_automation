# from items import Items, DESCRIPTION
# from sap_process import create_connection, order_product, confirm_transaction

# PATH_SAP = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
# PATH_STOCK = r"data/Stock-article-magasin-CRLD---2022.V2.xlsx"

# USE THIS IMPORT WITH THE NEW EXCEL FILE
from items_updated import Items, DESCRIPTION, CATEGORY    # To be used with the new xlsx file
from sap_process import create_connection, order_product, confirm_transaction
PATH_SAP = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
PATH_STOCK = r"data\Stock article magasin CRLD - updated.xlsx"   # To be used with the new file


if __name__ == "__main__":

    # start the application
    print("Loading DataFrame...")
    stock = Items(PATH_STOCK)
    print("Dataframe loaded!")

    # initialize Cart variable to store the user selection
    cart = dict()

    # Main Loop
    while True:
        print("\nChoose the categories/items using their corresponding number\n\nType 'cart' to see or remove your selected items\nType 'order' to validate your order")
        stock.display_categories()

        selection = input("Select category => ")

        # show cart if selected
        if selection.upper() == "CART":

            # restart the main loop if cart empty after informing the user
            if not cart:
                print("No item yet!\n")
                input("Press any  key to continue")
                print()
                continue

            # checking and removing submenu loop
            while True:
                if not cart:
                    selection = None
                    input("Cart  is now empty")
                    break

                # map the index of each value to the correponding code  {item_code: qty}
                index = dict()
                count = 0
                print()
                print("-----------Your CART----------")
                for code, qty in cart.items():
                    index[str(count)] = code
                    print(f"[{count}]---{stock.item_from_code(code)} x {qty} in cart")
                    count += 1
                
                # removing items from cart loop
                selection = input("\nSelect the item to remove or type any key to continue -> ")
                if selection in index:
                    cart[index[selection]] -= 1
                    if cart[index[selection]] == 0:
                        cart.pop(index[selection])
                else:
                    break
            continue

        # ordering process, if yes break from the main loop
        elif selection.upper() == "ORDER" or selection.upper() == "BUY":
            if not cart:
                print("Your cart is empty!")
                input()
                continue
            print("\nyou will order the following items from the stock:\n")
            for code, qty in cart.items():
                    print(f"--{stock.item_from_code(code)} x {qty} in cart")
            selection = input("\n Please confirm [y]/[n]")
            if selection == "y":
                break
            continue

        # select all the items from the selected categories
        try:
            category = (stock.categories[int(selection)])
        except (ValueError, IndexError):
            print("Wrong input\n")
            continue

        items = stock.select_category(category)
        stock.display_items(items)
        
        # Category submenu loop
        while True:
            selection = input("Select your Items -> ")
            if selection == "b" or selection == "":
                break

            # select code from item and add it to the cart
            try:
                item_code = items.iloc[int(selection), 1]
            except (ValueError, IndexError):
                print("Wrong input\n")
                continue
            
            # add the selected item in cart
            if item_code not in cart:
                cart[item_code] = 1
            else:
                cart[item_code] += 1

            print(f"----{stock.item_from_code(item_code)} added to your cart\n")
    
    # Enter here the code to sap_process
    try:
        session = create_connection(PATH_SAP)
    except FileNotFoundError:
        print()
        print(r"---/!\--- NO SAP APPLICATION  FOUND ---/!\---")
        print("Please  make sure SAP Logon is installed in your computer")
        print("Check that SAP path is provided into the variable 'PATH_SAP' from main.py")
        print("Order is cancelled.\n")
        
    # fill the SAP form
    order_product(session, cart)
    
    # confirm transaction and exit SAP program
    # confirm_transaction(session)