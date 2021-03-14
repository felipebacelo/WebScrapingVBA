![GitHub repo size](https://img.shields.io/github/repo-size/felipebacelo/WebScrapingVBA?style=for-the-badge)
![GitHub language count](https://img.shields.io/github/languages/count/felipebacelo/WebScrapingVBA?style=for-the-badge)
![GitHub forks](https://img.shields.io/github/forks/felipebacelo/WebScrapingVBA?style=for-the-badge)
![Bitbucket open pull requests](https://img.shields.io/bitbucket/pr-raw/felipebacelo/WebScrapingVBA?style=for-the-badge)
![Bitbucket open issues](https://img.shields.io/bitbucket/issues/felipebacelo/WebScrapingVBA?style=for-the-badge)

# WebScrapingVBA
Repositório com Simples Exemplos de WebScraping em VBA Excel

Este respositório foi desenvolvido com o objetivo de praticar alguns conceitos de WebScraping utilizando VBA Excel.

### Desenvolvimento

Desenvolvido em Microsoft VBA Excel.
***
### Requisitos

* Habilitar Macros
* Habilitar Guia de Desenvolvedor

### Referências às Bibliotecas

* Visual Basic For Applications
* Microsoft Excel 16.0 Object Library
* OLE Automation
* Microsoft Office 16.0 Object Library
* Microsoft Internet Controls
* Microsoft HTML Object Library

### Compatibilidade

Os exemplos deste repositório foram desenvolvidos no Excel 2019 (64 bits) e testados no Excel 2016 (64 bits). Sua compatibilidade é garantida para a versão 2016 e superiores. Sua utilização em versões anteriores pode ocasionar em não funcionamento dos mesmos.
***
### Exemplos de Códigos Utilizados

Código utilizado para realizar WebScraping de CEP:
```vba
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
```

Código utilizado para realizar WebScraping de CPF:
```vba
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
```
***
### Licenças

_MIT License_
_Copyright   ©   2021 Felipe Bacelo Rodrigues_
