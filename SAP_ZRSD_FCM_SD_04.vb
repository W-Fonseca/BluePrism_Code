'======= Ler Toda a tabela ====='
Dim SapGuiAuto = GetObject("SAPGUI")
Dim Application = SapGuiAuto.GetScriptingEngine
Dim Connection = Application.Children(0)
Dim session = Connection.Children(0)
Dim Coluna = -1

For Each SapObject2 As Object In session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").GetAllNodeKeys()
Tabela.Rows.Add()
Next

For Each SapObject As Object In session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").GetColumnNames()
Coluna = Coluna + 1
Tabela.Columns.Add(session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").GetColumnTitleFromName(SapObject),GetType(string))
Dim Linha = -1
For Each SapObject2 As Object In session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").GetAllNodeKeys()
Linha = Linha + 1
Tabela.Rows(linha)(coluna) = session.findById("wnd[0]/usr/cntlTREE1/shellcont/shell/shellcont[1]/shell[1]").GetItemText(SapObject2, SapObject)
Next
Next
