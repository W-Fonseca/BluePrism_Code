'precisa de um output collection com nome de Tabela
'Em inicialize foi adicionado em namespace imports: System.Collections.Generic

Dim session As Object
Dim Application As Object
Dim Connection As Object
Dim SapGuiAuto As Object

SapGuiAuto = GetObject("SAPGUI")
Application = SapGuiAuto.GetScriptingEngine
Connection = Application.Children(0)
session = Connection.Children(0)

dim contagem = 0
dim N_Aleatorio = 3
dim coluna = 0
dim texto = ""
dim texto2 = ""
dim linha = 0
for z= 0 to 1000
Tabela.rows.add()
next

On Error Resume Next
For i = 1 To 250 'colunas
texto = session.findById("wnd[0]/usr/sub/1[0,0]/sub/1/3[0,7]/lbl[" & i & ",8]").Text
contagem = 0
If texto <> "" Then
Tabela.Columns.Add(Trim(texto.Replace(".","_")),GetType(string))

'Cells(1, coluna).Value = texto
linha = 0
For x = 10 To 100 ' linhas

texto2 = session.findById("wnd[0]/usr/sub/1[0,0]/sub/1/3[0,7]/sub/1/3/" & x - 6 & "[0," & x & "]/lbl[" & i & "," & x & "]").Text
If texto2 <> "" Then



Tabela.rows(linha)(coluna) = texto2
'For Each dr As System.Data.DataRow In Tabela.Rows
	
'	dr(texto) = texto2
'Next
'Cells(x - 8, coluna).Value = texto2
linha = linha + 1
texto2 = ""
End If
Next
texto = ""
coluna = coluna + 1
End If
Next

Tabela = Tabela.DefaultView.ToTable(True)
