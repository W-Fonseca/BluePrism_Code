Dim workbookobjct = GetWorkbook(Handle,"")
Dim vbProj = workbookobjct.VBProject
Dim vbComp = vbProj.VBComponents.Add(1)
vbComp.Name = Module_Name
vbComp.CodeModule.AddFromString(Script)
