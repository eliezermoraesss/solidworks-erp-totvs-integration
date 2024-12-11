Dim swApp As Object

Dim Part As Object
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Dim retval As String
Dim resultado As VbMsgBoxResult

Sub main()

'frmdwgpdf.Show

Call SalvarPNG

If resultado = vbYes Then
Call SalvarPDF

End If

End Sub

Sub SalvarPDF()

Set swApp = _
Application.SldWorks

Set Part = swApp.ActiveDoc
nomeArquivo = Part.GetTitle

boolstatus = swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swPDFExportShadedDraftDPI, 300)
boolstatus = swApp.SetUserPreferenceIntegerValue(swUserPreferenceIntegerValue_e.swPDFExportOleDPI, 96)

If nomeArquivo Like "*[_]*" Then

    positionofSubstring = InStr(1, nomeArquivo, "_")
    nomeArquivo = Mid(nomeArquivo, 1, positionofSubstring - 1)

ElseIf nomeArquivo Like "[E]" & nomeArquivo Like "[A]" Then

    positionofSubstring = InStr(1, nomeArquivo, "-")
    nomeArquivo = Mid(nomeArquivo, 1, positionofSubstring - 2)

Else
    positionofSubstring = InStrRev(nomeArquivo, "-")
    nomeArquivo = Mid(nomeArquivo, 1, positionofSubstring - 2)

End If

If Len(nomeArquivo) = 8 Then

    primeiroTrecho = Left(nomeArquivo, 3)
    segundoTrecho = Mid(nomeArquivo, 4, 2)
    terceiroTrecho = Mid(nomeArquivo, 6, 3)
    
    longstatus = Part.SaveAs3("Y:\PDF-OFICIAL\" & "M-" & primeiroTrecho & "-0" & segundoTrecho & "-" & terceiroTrecho & ".PDF", 0, 0)
    
End If

If Len(nomeArquivo) = 10 Then

    primeiroTrecho = Left(nomeArquivo, 3)
    segundoTrecho = Mid(nomeArquivo, 5, 2)
    terceiroTrecho = Mid(nomeArquivo, 8, 3)
    
    longstatus = Part.SaveAs3("Y:\PDF-OFICIAL\" & "M-" & primeiroTrecho & "-0" & segundoTrecho & "-" & terceiroTrecho & ".PDF", 0, 0)
    
End If

If Len(nomeArquivo) = 13 Then

    longstatus = Part.SaveAs3("Y:\PDF-OFICIAL\" & nomeArquivo & ".PDF", 0, 0)

End If

End Sub

Sub SalvarPNG()

Dim nomeArquivoComparar As String

Set swApp = _
Application.SldWorks

Set Part = swApp.ActiveDoc
nomeArquivo = Part.GetTitle

If nomeArquivo Like "*[_]*" Then

    positionofSubstring = InStr(1, nomeArquivo, "_")
    nomeArquivo = Mid(nomeArquivo, 1, positionofSubstring - 1)

ElseIf nomeArquivo Like "[E]" & nomeArquivo Like "[A]" Then

    positionofSubstring = InStr(1, nomeArquivo, "-")
    nomeArquivo = Mid(nomeArquivo, 1, positionofSubstring - 2)

Else
    positionofSubstring = InStrRev(nomeArquivo, "-")
    nomeArquivo = Mid(nomeArquivo, 1, positionofSubstring - 2)

End If

If Len(nomeArquivo) = 8 Then

    primeiroTrecho = Left(nomeArquivo, 3)
    segundoTrecho = Mid(nomeArquivo, 4, 2)
    terceiroTrecho = Mid(nomeArquivo, 6, 3)
    pathFileName1 = "Y:\PDF-OFICIAL\" & "M-" & primeiroTrecho & "-0" & segundoTrecho & "-" & terceiroTrecho & ".tif"
    pathFileNameComp = "Y:\PDF-OFICIAL\" & "M-" & primeiroTrecho & "-0" & segundoTrecho & "-" & terceiroTrecho & ".PNG"
    
    nomeArquivoComparar = Dir(pathFileNameComp)
    nomeArquivoTemp = "M-" & primeiroTrecho & "-0" & segundoTrecho & "-" & terceiroTrecho & ".PNG"
    nomeArquivoTempComp = pathFileNameComp
    
If UCase(nomeArquivoComparar) <> nomeArquivoTemp Then
    
    longstatus = Part.SaveAs3(pathFileName1, 0, 0)

    positionofSubstring = InStr(1, pathFileName1, ".")
    nomeArquivoPng = Mid(pathFileName1, 1, positionofSubstring - 1)
    
    Name pathFileName1 As _
    nomeArquivoPng + ".PNG"
    
    Call SalvarPDF

Else

Call ExibirMsgBox

If resultado = vbYes Then

    Kill (nomeArquivoTempComp)
    
    longstatus = Part.SaveAs3(pathFileName1, 0, 0)
    
    positionofSubstring = InStr(1, pathFileName1, ".")
    nomeArquivoPng = Mid(pathFileName1, 1, positionofSubstring - 1)
    
    Name pathFileName1 As _
    nomeArquivoPng + ".PNG"

End If
End If
End If

If Len(nomeArquivo) = 10 Then '000.00.000'

    primeiroTrecho = Left(nomeArquivo, 3)
    segundoTrecho = Mid(nomeArquivo, 5, 2)
    terceiroTrecho = Mid(nomeArquivo, 8, 3)
    pathFileName1 = "Y:\PDF-OFICIAL\" & "M-" & primeiroTrecho & "-0" & segundoTrecho & "-" & terceiroTrecho & ".tif"
    pathFileNameComp = "Y:\PDF-OFICIAL\" & "M-" & primeiroTrecho & "-0" & segundoTrecho & "-" & terceiroTrecho & ".PNG"
    
    nomeArquivoComparar = Dir(pathFileNameComp)
    nomeArquivoTemp = "M-" & primeiroTrecho & "-0" & segundoTrecho & "-" & terceiroTrecho & ".PNG"
    nomeArquivoTempComp = pathFileNameComp
    
If UCase(nomeArquivoComparar) <> nomeArquivoTemp Then
    
    longstatus = Part.SaveAs3(pathFileName1, 0, 0)

    positionofSubstring = InStr(1, pathFileName1, ".")
    nomeArquivoPng = Mid(pathFileName1, 1, positionofSubstring - 1)
    
    Name pathFileName1 As _
    nomeArquivoPng + ".PNG"
    
    Call SalvarPDF

Else

Call ExibirMsgBox

If resultado = vbYes Then

    Kill (nomeArquivoTempComp)
    
    longstatus = Part.SaveAs3(pathFileName1, 0, 0)
    
    positionofSubstring = InStr(1, pathFileName1, ".")
    nomeArquivoPng = Mid(pathFileName1, 1, positionofSubstring - 1)
    
    Name pathFileName1 As _
    nomeArquivoPng + ".PNG"

End If
End If
End If

If Len(nomeArquivo) = 13 Then

nomeArquivoComparar = Dir("Y:\PDF-OFICIAL\" & nomeArquivo & ".PNG")
nomeArquivoTemp = nomeArquivo & ".PNG"
nomeArquivoTempComp = "Y:\PDF-OFICIAL\" & nomeArquivo & ".PNG"

If UCase(nomeArquivoComparar) <> nomeArquivoTemp Then

    pathFileName2 = "Y:\PDF-OFICIAL\" & nomeArquivo & ".tif"

    longstatus = Part.SaveAs3(pathFileName2, 0, 0)
    
    positionofSubstring = InStr(1, pathFileName2, ".")
    nomeArquivoPng = Mid(pathFileName2, 1, positionofSubstring - 1)
    
    Name pathFileName2 As _
    nomeArquivoPng + ".PNG"
    
    Call SalvarPDF
    
Else

Call ExibirMsgBox

If resultado = vbYes Then

Kill (nomeArquivoTempComp)

pathFileName2 = "Y:\PDF-OFICIAL\" & nomeArquivo & ".tif"

    longstatus = Part.SaveAs3(pathFileName2, 0, 0)
    
    positionofSubstring = InStr(1, pathFileName2, ".")
    nomeArquivoPng = Mid(pathFileName2, 1, positionofSubstring - 1)
    
    Name pathFileName2 As _
    nomeArquivoPng + ".PNG"

End If

End If

End If

End Sub

Public Sub ExibirMsgBox()
     textoCorpo = "Já existe uma revisão deste desenho." & vbNewLine & vbNewLine & "Deseja prosseguir?" & vbNewLine & vbNewLine & "CERTIFIQUE-SE QUE O DESENHO ATUAL ESTEJA ATUALIZADO!" & vbNewLine & vbNewLine & "Tenha um excelente dia! =)" & vbNewLine & "Engenharia ENAPLIC."
     resultado = MsgBox(textoCorpo, vbYesNo + vbQuestion, "ATENÇÃO!")
     'MsgBox resultado
End Sub
