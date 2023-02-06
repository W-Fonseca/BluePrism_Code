Dim session As Object
Dim Application As Object
Dim Connection As Object
Dim SapGuiAuto As Object

SapGuiAuto = GetObject("SAPGUI")
Application = SapGuiAuto.GetScriptingEngine
Connection = Application.Children(0)
Session = Connection.Children(0)
