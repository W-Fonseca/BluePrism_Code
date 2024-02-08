'Pagina Return All Properties in Page
'retorna tudo que está na aberto do sap

'Output Table = Collection (vazio)

Dim SapGuiAuto = GetObject("SAPGUI")
Dim App = SapGuiAuto.GetscriptingEngine
Dim Connection = App.Children(0)
Dim Session = Connection.Children(0)

Table.Columns.Add("ContainerType", GetType(string))
Table.Columns.Add("Id", GetType(string))
Table.Columns.Add("Name", GetType(string))
Table.Columns.Add("Parent", GetType(string))
Table.Columns.Add("Type", GetType(string))
Table.Columns.Add("TypeAsNumber", GetType(string))
Table.Columns.Add("AccLabelCollection", GetType(string))
Table.Columns.Add("AccText", GetType(string))
Table.Columns.Add("AccTextOnRequest", GetType(string))
Table.Columns.Add("AccTooltip", GetType(string))
Table.Columns.Add("Changeable", GetType(string))
Table.Columns.Add("DefaultTooltip", GetType(string))
Table.Columns.Add("Height", GetType(string))
Table.Columns.Add("IconName", GetType(string))
Table.Columns.Add("IsSymbolFont", GetType(string))
Table.Columns.Add("Left", GetType(string))
Table.Columns.Add("Modified", GetType(string))
Table.Columns.Add("ParentFrame", GetType(string))
Table.Columns.Add("ScreenLeft", GetType(string))
Table.Columns.Add("ScreenTop", GetType(string))
Table.Columns.Add("Text", GetType(string))
Table.Columns.Add("Tooltip", GetType(string))
Table.Columns.Add("Top", GetType(string))
Table.Columns.Add("Width", GetType(string))

For Each SapObject As Object In Session.Children
   Call AdicionarSapObjectNaTabela(SapObject, Table)
Next

'=======================================================
'Codigo na pagina principal
'=======================================================
Sub AdicionarSapObjectNaTabela(SapObject As Object, Table As DataTable)
on error resume next
Table.Rows.Add(
SapObject.ContainerType,	
SapObject.Id,	
SapObject.Name,	
SapObject.Parent,	
SapObject.Type,	
SapObject.TypeAsNumber,	
SapObject.AccLabelCollection,	
SapObject.AccText,	
SapObject.AccTextOnRequest,	
SapObject.AccTooltip,	
SapObject.Changeable,	
SapObject.DefaultTooltip,	
SapObject.Height,	
SapObject.IconName,	
SapObject.IsSymbolFont,	
SapObject.Left,	
SapObject.Modified,	
SapObject.ParentFrame,	
SapObject.ScreenLeft,	
SapObject.ScreenTop,	
SapObject.Text,	
SapObject.Tooltip,	
SapObject.Top,	
SapObject.Width)

    If SapObject.ContainerType = True Then
        For Each SapOb As Object In SapObject.Children
            AdicionarSapObjectNaTabela(SapOb, Table)
        Next
    End If
End Sub

