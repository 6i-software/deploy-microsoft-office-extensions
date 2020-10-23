Deploy Microsoft Office Extensions
==================================

## How to split a word file every x pages in VBA ?

```vba
'---------------------------------------------------------------------------------
' @copyright 6i (2020)
' @author 20100 <vb20100bv@gmail.com>
' Released under a MIT license.
'---------------------------------------------------------------------------------

Dim FolderOutput As String

'----------------------------------------------------------------------------------
' Give ability to make your function callable from the UI application Ribbon Office
'----------------------------------------------------------------------------------
Sub apply_CutterFileByPage_eventhandler(control As IRibbonControl)
    CutterFileByPage
End Sub

Function CountPagesInDocument()
    CountPagesInDocument = ActiveDocument.Range.Information(wdNumberOfPagesInDocument)
End Function

Sub DeleteLastPageBreak(wdDoc As Document)
  Dim i As Long
  For i = wdDoc.Paragraphs.Count To 1 Step -1
    If Asc(wdDoc.Paragraphs(i).Range.Text) = 12 Then
      wdDoc.Paragraphs(i).Range.Delete
      Exit For
    End If
    If Len(wdDoc.Paragraphs(i).Range.Text) > 1 Then
      Exit For
    End If
  Next i
End Sub

Sub CopyPagesContent(pageBegin As Integer, pageEnd As Integer, nbPages As Integer)
    Dim Rng As Range, RngTmp As Range, wdDoc As Document
  
    ' Select pages
    With ActiveDocument
        Set Rng = .Range.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pageBegin)
        Set RngTmp = .Range.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pageEnd)
        Set RngTmp = RngTmp.GoTo(What:=wdGoToBookmark, name:="\Page")
        Rng.End = RngTmp.End: Set RngTmp = Nothing
    End With
  
    ' Create output directory if necessary
    'FolderOutput = ActiveDocument.Path & "\" & Left(ActiveDocument.name, InStrRev(ActiveDocument.name, ".") - 1) & "_CUT_BY_" & nbPages & "_PAGES"
    FolderOutput = "c:\test\"
    If Dir(FolderOutput, vbDirectory) = "" Then
        MkDir FolderOutput
    End If
    
    ' Create word file
    Dim filename As String
    'filename = FolderOutput & "\" & Left(ActiveDocument.name, InStrRev(ActiveDocument.name, ".") - 1) & "_" & pageBegin & "_to_" & pageEnd & ".docx"
    filename = FolderOutput & "\test_" & pageBegin & "_to_" & pageEnd & ".docx"
    Set wdDoc = Documents.Add
    With wdDoc
        Rng.Copy
        .Range.Characters.Last.PasteAndFormat Type:=wdFormatOriginalFormatting
    End With
    
    ' Check if we must remove last page
    If wdDoc.Range.Information(wdNumberOfPagesInDocument) > ((pageEnd - pageBegin) + 1) Then
        DeleteLastPageBreak wdDoc
    End If
    
    ' Save word
    With wdDoc
        .SaveAs filename:=filename
        .Close
    End With
    Set Rng = Nothing
    Set RngTmp = Nothing
End Sub

Sub CutterFileByPage()
    Dim x As Integer
    Dim reste As Integer
    Dim selected As Range
    
    Dim nbPages As Integer
    Dim nbPagesRestantes As Integer
    nbPages = InputBox("Combien de pages par découpe voulez-vous ?", "Nombre de pages", 2)
    
    Application.ScreenUpdating = False
    For x = 1 To CountPagesInDocument
        reste = x Mod nbPages
        nbPagesRestantes = CountPagesInDocument - x
        Debug.Print "Page " & x & " | NbPageRestante " & nbPagesRestantes

        If reste = 0 Then
            Debug.Print "Copy pages (" & x - (nbPages - 1) & ";" & x & ")"
            CopyPagesContent x - (nbPages - 1), x, nbPages
        End If
        
        If nbPagesRestantes < nbPages - 1 Then
            Debug.Print "Copy pages restantes (" & x & ";" & CountPagesInDocument & ")"
            CopyPagesContent x, CountPagesInDocument, nbPages
            Exit For
        End If
    Next x
    Application.ScreenUpdating = True
    
    Shell "explorer.exe" & " " & FolderOutput, vbNormalFocus
End Sub
```