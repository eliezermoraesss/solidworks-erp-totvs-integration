Dim swApp As SldWorks.SldWorks
Dim swModDoc As SldWorks.IModelDoc2
Dim swTable As SldWorks.ITableAnnotation

Dim status As Boolean
Dim caminhoDoDesenho As String
Dim nomeDoDesenho As String
Dim bomPath As String

Option Explicit
Public Sub Main()

    On Error Resume Next
    
    Set swApp = Application.SldWorks
    Set swModDoc = swApp.ActiveDoc
    Dim swSM As ISelectionMgr
    Set swSM = swModDoc.SelectionManager
    Set swTable = swSM.GetSelectedObject6(1, -1)
    
    Dim swSpecTableCutListWeldment As IWeldmentCutListAnnotation
    Set swSpecTableCutListWeldment = swTable

    Dim swSpecTable As IBomTableAnnotation
    Set swSpecTable = swTable
    
    ' get the count of selected objects
    Dim Count As Long
    Count = swSM.GetSelectedObjectCount2(-1)
    ' if the user has made no selection then exit
    If Count = 0 Then
        Call MensagemBOM
    Exit Sub
    End If
    
    ' Obter o caminho e nome do arquivo do desenho aberto
    caminhoDoDesenho = swModDoc.GetPathName()
    
    If caminhoDoDesenho = "" Then
        Call MensagemSalvarDesenho
        Exit Sub
    End If
    
    nomeDoDesenho = ExtrairNomeDoArquivo(caminhoDoDesenho)
    
    Call DefinirVariavelAmbiente(nomeDoDesenho)
    
    bomPath = Environ("TEMP") & "\" & nomeDoDesenho & ".xlsx"
    
    DeleteFileIfExists bomPath
    
    If TypeOf swTable Is IWeldmentCutListAnnotation Then
        status = SaveAsExcelBOMListaCorteSoldagem(bomPath)
    Else
        ' Save the selected BOM table to Microsoft Excel, including hidden cells and images (path, cells, images)
        status = swSpecTable.SaveAsExcel(bomPath, True, False)
    End If

    ' Release the reference to the TableAnnotation object
    Set swTable = Nothing
    
    If status = True Then
        Call ExecutarScriptPython
    Else
        MsgBox "Falha ao salvar BOM no formato Excel.", vbCritical, "Erro"
    End If

End Sub
Function SaveAsExcelBOMListaCorteSoldagem(bomPath As String) As Boolean
    Dim swCutList As IWeldmentCutListAnnotation
    Set swCutList = swTable
    
    ' Crie um arquivo Excel usando o Excel.Application
    Dim xlApp As Object
    Dim xlWB As Object
    Dim xlSheet As Object

    Set xlApp = CreateObject("Excel.Application")
    Set xlWB = xlApp.Workbooks.Add
    Set xlSheet = xlWB.Sheets(1)
    
    ' Preencha o Excel manualmente
    Dim i As Long, j As Long
    For i = 0 To swCutList.RowCount - 1
        For j = 0 To swCutList.ColumnCount - 1
            xlSheet.Cells(i + 1, j + 1).Value = swCutList.Text(i, j)
        Next j
    Next i

    ' Salve e feche o arquivo
    xlWB.SaveAs bomPath
    xlWB.Close
    xlApp.Quit

    Set xlSheet = Nothing
    Set xlWB = Nothing
    Set xlApp = Nothing
    
    SaveAsExcelBOMListaCorteSoldagem = True
    
End Function
Function ExtrairNomeDoArquivo(caminhoDoArquivo As String) As String
    Dim caminho As String
    Dim nomeDoArquivo As String
    Dim posicaoDaUltimaBarra As Integer
    Dim posicaoDoPonto As Integer

    ' Defina a string do caminho do arquivo
    caminho = caminhoDoArquivo

    ' Encontre a posição da última barra e do ponto na string
    posicaoDaUltimaBarra = InStrRev(caminho, "\")
    posicaoDoPonto = InStrRev(caminho, ".")

    ' Extraia a substring desejada
    nomeDoArquivo = Mid(caminho, posicaoDaUltimaBarra + 1, posicaoDoPonto - posicaoDaUltimaBarra - 1)
    
    ExtrairNomeDoArquivo = nomeDoArquivo
End Function
Sub ExecutarScriptPython()
    Dim CaminhoArquivo As String
    
    ' Substitua o caminho abaixo pelo caminho do seu arquivo botao-salvar-bom-solidworks-totvs
    CaminhoArquivo = "\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\VBA\PYTHON\botao-salvar-bom-solidworks-totvs_CRUD_TEST.pyw"
    
    ' Use a função Shell para abrir o arquivo com a aplicação padrão
    Shell "explorer.exe """ & CaminhoArquivo & """", vbNormalFocus
End Sub
Sub DefinirVariavelAmbiente(codigoDesenho As String)

    ' Substitua "meuarquivo" pelo nome real do seu arquivo (sem a extensão)
    Dim nomeArquivo As String
    nomeArquivo = codigoDesenho

    ' Construa o comando para definir a variável de ambiente
    Dim comando As String
    comando = "setx CODIGO_DESENHO " & nomeArquivo

    ' Execute o comando no shell
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run comando, 0, True
    
End Sub
Sub DeleteFileIfExists(filePath As String)
    If Dir(filePath) <> "" Then
        ' File exists, so delete it
        Kill filePath
        filePath = ""
    End If
End Sub
Sub MensagemBOM()
    MsgBox "SELECIONE A TABELA DE BOM" & vbNewLine & vbNewLine & "Por gentileza, clique em qualquer parte da tabela de BOM e tente novamente!", vbInformation, "ATENÇÃO"
End Sub
Sub MensagemSalvarDesenho()
    MsgBox "SALVE O DESENHO!" & vbNewLine & vbNewLine & "Para continuar com o cadastro da BOM, por favor, salve este detalhamento!", vbInformation, "ATENÇÃO"
End Sub
