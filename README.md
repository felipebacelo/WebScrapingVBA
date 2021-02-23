![GitHub repo size](https://img.shields.io/github/repo-size/felipebacelo/ApplicationVBA?style=for-the-badge)
![GitHub language count](https://img.shields.io/github/languages/count/felipebacelo/ApplicationVBA?style=for-the-badge)
![GitHub forks](https://img.shields.io/github/forks/felipebacelo/ApplicationVBA?style=for-the-badge)
![Bitbucket open pull requests](https://img.shields.io/bitbucket/pr-raw/felipebacelo/ApplicationVBA?style=for-the-badge)
![Bitbucket open issues](https://img.shields.io/bitbucket/issues/felipebacelo/ApplicationVBA?style=for-the-badge)

# WebScrapingVBA
Repositório com Simples Exemplos de WebScraping em VBA

A aplicação foi desenvolvida a partir do modelo de cadastro físico do programa Cidade Legal, seguindo o conteúdo padrão do mesmo, a partir do conceito CRUD (Create, Read, Update, Delete), que representa em acrônimo as quatro operações básicas utilizadas em bases de dados relacionais fornecidas aos utilizadores do sistema.

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

Esta aplicação foi desenvolvida no Excel 2019 (64 bits) e testado no Excel 2016 (64 bits). Sua compatibilidade é garantida para a versão 2016 e superiores. Sua utilização em versões anteriores pode ocasionar em não funcionamento do mesmo.
***
### Exemplos de Códigos Utilizados

Macro utilizada para conexão com banco de dados Microsoft Access SQL:
```vba
Option Explicit
Global BD As New ADODB.Connection

Sub ABRIRCONEXAO()

Dim CS As String
Dim ARQ As String
On Error Resume Next

ARQ = ThisWorkbook.Path & "\" & "BD SISTEMA DE CADASTRO.accdb;"

CS = "Provider=Microsoft.ACE.OLEDB.12.0;" _
& "Data Source=" & ARQ _
& "Persist Security Info=False;"

BD.Close
BD.Open CS

End Sub
```

Macro utilizada para edição dos registros salvos no banco de dados Microsoft Access SQL:
```vba
Sub EDITARREGISTROS(ID As Long, TODASCOLUNAS As String, REGISTRO() As String)

Dim SQL As String
Dim COLUNA() As String
Dim I As Integer
Dim STRINGFINAL As String
Dim RS As New ADODB.Recordset

COLUNA = Split(TODASCOLUNAS, ",")

For I = 1 To 81
    STRINGFINAL = STRINGFINAL & COLUNA(I - 1) & "=" & REGISTRO(I)
    If I < 81 Then STRINGFINAL = STRINGFINAL & ","
Next

STRINGFINAL = "SET " & STRINGFINAL
SQL = "Update CADASTROS " & STRINGFINAL
SQL = SQL & " WHERE ID LIKE " & ID

RS.Open SQL, BD

MsgBox "CADASTRO EDITADO COM SUCESSO!", vbInformation, "INFORMAÇÃO"

End Sub
```
***
### Licenças

_MIT License_
_Copyright   ©   2021 Felipe Bacelo Rodrigues_
