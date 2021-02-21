Attribute VB_Name = "MÓDULO_IE_CPF"
Sub IE_CPF()

Range("B3:C3").ClearContents

Set IE = CreateObject("internetexplorer.application")

IE.Navigate "http://www.situacao-cadastral.com/"
IE.Visible = True

Do While IE.busy And IE.ReadyState <> "READYSTATE_COMPLETE"
DoEvents
Loop

IE.Document.getElementById("doc").Value = Cells(3, 1).Value
IE.Document.getElementById("consultar").Click
    
Do While IE.busy And IE.ReadyState <> "READYSTATE_COMPLETE"
DoEvents
Loop

Cells(3, 2) = IE.Document.getElementsByClassName("dados nome")(0).innertext
Cells(3, 3) = IE.Document.getElementsByClassName("dados situacao")(0).innertext

IE.Quit

Range("A3:C3").WrapText = False

End Sub
