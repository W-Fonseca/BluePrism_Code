V_Input = Trim(V_Input)
V_Output = Decimal.Parse(V_Input, New System.Globalization.CultureInfo("pt-BR")).Tostring("F", New System.Globalization.CultureInfo("en-US"))
