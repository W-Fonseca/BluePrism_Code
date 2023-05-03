sap 

   session.findById("wnd[0]/usr").FindByNameEx("BUTTON_HEADER_TOGGLE", 40).press()
        session.findById("wnd[0]/usr").FindByNameEx("BUTTON_ITEMDETAIL", 40).press()
    session.findById("wnd[0]/usr").FindByNameEx("GODYNPRO-ACTION", 34).Key = "A01"
    session.findById("wnd[0]/usr/").FindByNameEx("GODYNPRO-REFDOC", 34).Key = "R01"
    session.findById("wnd[0]/usr/").FindByNameEx("GODYNPRO-PO_NUMBER", 32).Text = po
    session.findById("wnd[0]/usr/").FindByNameEx("GODYNPRO-PO_ITEM", 31).Text = po_item
    session.findById("wnd[0]/usr/").FindByNameEx("GODYNPRO-PO_WERKS", 32).Text = ""
    session.findById("wnd[0]/usr").FindByNameEx("GODEFAULT_TV-BWART", 32).Text = "101"
