Attribute VB_Name = "MÓDULO_IE_CEP"
Sub IE_CEP()

Range("B3:D3").ClearContents

Set IE = CreateObject("internetexplorer.application")

IE.Navigate "http://www.buscacep.correios.com.br/sistemas/buscacep/"
IE.Visible = True

Do While IE.busy And IE.ReadyState <> "READYSTATE_COMPLETE"
DoEvents
Loop

IE.Document.getElementsByTagName("input")(0).Value = Cells(3, 1).Value
IE.Document.getElementsByClassName("btn2 float-right")(0).Click
    
Do While IE.busy And IE.ReadyState <> "READYSTATE_COMPLETE"
DoEvents
Loop

Cells(3, 2) = IE.Document.getElementsByTagName("td")(0).innertext
Cells(3, 3) = IE.Document.getElementsByTagName("td")(1).innertext
Cells(3, 4) = IE.Document.getElementsByTagName("td")(2).innertext

IE.Quit

Range("A3:D3").WrapText = False

End Sub
