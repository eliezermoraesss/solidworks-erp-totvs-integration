Option Explicit

' Constants for file paths and extensions
Private Const PDF_OFFICIAL_PATH As String = "Y:\PDF-OFICIAL\"
Private Const REVISIONS_FOLDER As String = "REVISOES_ANTERIORES"
Private Const FILE_EXT_TIF As String = ".tif"
Private Const FILE_EXT_PNG As String = ".PNG"
Private Const FILE_EXT_PDF As String = ".PDF"
Private Const PREFIX_M As String = "M-"
Private Const SUFIX_REV As String = "REV_"
Private Const EXIST_PNG_PDF_MSG As String = "Já existe uma revisão deste desenho." & vbNewLine & vbNewLine & _
             "Deseja prosseguir?" & vbNewLine & vbNewLine & _
             "CERTIFIQUE-SE QUE O DESENHO ATUAL ESTEJA ATUALIZADO!" & vbNewLine & vbNewLine & _
             "Tenha um excelente dia! =)" & vbNewLine & _
             "Engenharia ENAPLIC."
Private Const ASK_SAVE_BACKUP As String = "Deseja salvar uma revisão do PDF anterior?" & vbNewLine & vbNewLine & _
             "Clicando em SIM, o PDF anterior será salvo na pasta REVISOES_ANTERIORES em PDF-OFICIAL." & vbNewLine & vbNewLine & _
             "Tenha um excelente dia! =)" & vbNewLine & _
             "Engenharia ENAPLIC."

' Object declarations
Private swApp As Object
Private activePart As Object

' Variable declarations
Dim resultado As VbMsgBoxResult
Dim firstSaving As Boolean


' Main entry point
Sub main()
    Dim success As Boolean
    success = SaveAsPNG()
    
    If success And resultado = vbYes Then
        firstSaving = False
        Call SaveAsPDF
    ElseIf success Then
        firstSaving = True
        Call SaveAsPDF
    End If
End Sub

' Handles file name processing
Private Function ProcessFileName(originalName As String) As String
    If InStr(1, originalName, "_") > 0 Then
        ProcessFileName = Left(originalName, InStr(1, originalName, "_") - 1)
    ElseIf originalName Like "[E]*[A]*" Then
        ProcessFileName = Left(originalName, InStr(1, originalName, "-") - 2)
    Else
        ProcessFileName = Left(originalName, InStrRev(originalName, "-") - 2)
    End If
End Function

' Formats file name based on length
Private Function FormatFileName(baseFileName As String) As String
    Dim firstPart As String, secondPart As String, thirdPart As String
    
    Select Case Len(baseFileName)
        Case 8
            firstPart = Left(baseFileName, 3)
            secondPart = Mid(baseFileName, 4, 2)
            thirdPart = Mid(baseFileName, 6, 3)
            FormatFileName = PREFIX_M & firstPart & "-0" & secondPart & "-" & thirdPart
            
        Case 10
            firstPart = Left(baseFileName, 3)
            secondPart = Mid(baseFileName, 5, 2)
            thirdPart = Mid(baseFileName, 8, 3)
            FormatFileName = PREFIX_M & firstPart & "-0" & secondPart & "-" & thirdPart
            
        Case 13
            FormatFileName = baseFileName
            
        Case Else
            FormatFileName = ""
    End Select
End Function

' Initialize SolidWorks application
Private Sub InitializeSolidWorks()
    Set swApp = Application.SldWorks
    Set activePart = swApp.ActiveDoc
End Sub

' Creates revisions folder if it doesn't exist
Private Function EnsureRevisionsFolder() As String
    Dim revisionsPath As String
    revisionsPath = PDF_OFFICIAL_PATH & REVISIONS_FOLDER
    
    If Dir(revisionsPath, vbDirectory) = "" Then
        MkDir revisionsPath
    End If
    
    EnsureRevisionsFolder = revisionsPath
End Function

' Generates timestamp string for file names
Private Function GetTimestampString() As String
    GetTimestampString = Format(Now, "ddmmyyyy_hhnnss")
End Function

Private Function GetNextRevisionNumber(baseFileName As String, revisionsPath As String) As String
    Dim fileCount As Integer
    Dim file As String
    Dim pattern As String
    
    ' Create search pattern for similar files
    pattern = baseFileName & "*" & FILE_EXT_PDF
    
    ' Count files only in revisions folder
    fileCount = 0
    file = Dir(revisionsPath & "\" & pattern)
    
    Do While file <> ""
        fileCount = fileCount + 1
        file = Dir()
    Loop
    
    ' Next revision will be current count + 1
    GetNextRevisionNumber = SUFIX_REV & Format(fileCount + 1, "00")
End Function

' Backs up existing PDF file if it exists
Private Sub BackupExistingPDF(pdfPath As String)
    If Dir(pdfPath) <> "" Then
        On Error GoTo ErrorHandler
        
        Dim revisionsPath As String
        revisionsPath = EnsureRevisionsFolder()
        
        Dim fileName As String
        fileName = Mid(pdfPath, InStrRev(pdfPath, "\") + 1)
        fileName = Left(fileName, InStrRev(fileName, ".") - 1)
        
        Dim revisionNumber As String
        revisionNumber = GetNextRevisionNumber(fileName, revisionsPath)
        
        Dim newFileName As String
        newFileName = fileName & "_" & GetTimestampString() & "_" & revisionNumber & FILE_EXT_PDF
        
        Dim backupPath As String
        backupPath = revisionsPath & "\" & newFileName
        
        FileCopy pdfPath, backupPath
        
        ' Verify if backup was successful
        If Dir(backupPath) = "" Then
            MsgBox "Erro: Backup não foi criado corretamente!", vbCritical, "Erro"
        End If
        Exit Sub
        
ErrorHandler:
        MsgBox "Erro ao fazer backup do arquivo!" & vbNewLine & _
               "Erro: " & Err.Description, vbCritical, "Erro"
    End If
End Sub

' Save document as PNG
Private Function SaveAsPNG() As Boolean
    InitializeSolidWorks
    
    Dim baseFileName As String
    baseFileName = ProcessFileName(activePart.GetTitle)
    
    If Len(baseFileName) = 0 Then Exit Function
    
    Dim formattedName As String
    formattedName = FormatFileName(baseFileName)
    
    ' Generate file paths
    Dim tifPath As String, pngPath As String
    tifPath = PDF_OFFICIAL_PATH & formattedName & FILE_EXT_TIF
    pngPath = PDF_OFFICIAL_PATH & formattedName & FILE_EXT_PNG
    
    ' Check if file exists
    Dim existingFile As String
    existingFile = Dir(pngPath)
    
    If UCase(existingFile) <> formattedName & FILE_EXT_PNG Then
        SaveNewPNGFile tifPath
        SaveAsPNG = True
    Else
        ShowConfirmationDialog EXIST_PNG_PDF_MSG
        If resultado = vbYes Then
            Kill pngPath
            SaveNewPNGFile tifPath
            SaveAsPNG = True
        End If
    End If
End Function

' Save document as PDF
Private Sub SaveAsPDF()
    InitializeSolidWorks
    ConfigurePDFSettings
    
    Dim baseFileName As String
    baseFileName = ProcessFileName(activePart.GetTitle)
    
    If Len(baseFileName) = 0 Then
        MsgBox "Nome de arquivo inválido!", vbCritical, "Erro"
        Exit Sub
    End If
    
    Dim formattedName As String
    formattedName = FormatFileName(baseFileName)
    
    Dim pdfPath As String
    pdfPath = PDF_OFFICIAL_PATH & formattedName & FILE_EXT_PDF
    
    If Not firstSaving Then
        ShowConfirmationDialog ASK_SAVE_BACKUP
        If resultado = vbYes Then
            ' Backup existing PDF if it exists
            On Error Resume Next
            BackupExistingPDF pdfPath
            On Error GoTo 0
        End If
    End If
    
    ' Save new PDF
    activePart.SaveAs3 pdfPath, 0, 0
    
    ' Verify if PDF exists after save
    If Dir(pdfPath) <> "" Then
        MsgBox "PDF salvo com sucesso!", vbInformation, "Eureka® PDF"
    Else
        MsgBox "Falha ao salvar arquivo PDF!" & vbNewLine & _
               "Verifique as permissões da pasta e espaço em disco.", vbCritical, "Erro"
    End If
End Sub

' Configure PDF export settings
Private Sub ConfigurePDFSettings()
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swPDFExportShadedDraftDPI, 300
    swApp.SetUserPreferenceIntegerValue swUserPreferenceIntegerValue_e.swPDFExportOleDPI, 96
End Sub

' Save new PNG file
Private Sub SaveNewPNGFile(tifPath As String)
    activePart.SaveAs3 tifPath, 0, 0
    
    ' First check if TIF was saved
    If Dir(tifPath) = "" Then
        MsgBox "Falha ao salvar arquivo TIF!", vbCritical, "Erro"
        Exit Sub
    End If
    
    ' Attempt to rename to PNG
    Dim pngPath As String
    pngPath = Left(tifPath, InStr(1, tifPath, ".") - 1) & FILE_EXT_PNG
    
    On Error GoTo ErrorHandler
    Name tifPath As pngPath
    
    ' Verify if PNG exists after rename
    If Dir(pngPath) <> "" Then
        MsgBox "PNG salvo com sucesso!", vbInformation, "Eureka® PNG"
    Else
        MsgBox "Erro ao verificar arquivo PNG após salvamento!", vbCritical, "Erro"
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Erro ao renomear arquivo TIF para PNG!" & vbNewLine & _
           "Erro: " & Err.Description, vbCritical, "Erro"
End Sub

' Show confirmation dialog
Private Sub ShowConfirmationDialog(message As String)
    resultado = MsgBox(message, vbYesNo + vbQuestion, "ATENÇÃO!")
End Sub

