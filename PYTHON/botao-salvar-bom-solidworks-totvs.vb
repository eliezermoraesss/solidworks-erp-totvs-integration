'This example shows how to save a BOM table annotation to a Microsoft Excel document.

'---------------------------------------------------------------------------
' Preconditions:
' 1. Open a drawing of an assembly with a BOM.
' 2. Click the move-table icon in the upper-left corner
'    of the BOM table to open the table's PropertyManager page.
'
' Postconditions: Saves the selected BOM to c:\temp\BOMTable.xls.
'--------------------------------------------------------------------------
Dim swApp As SldWorks.SldWorks
Dim swModDoc As SldWorks.IModelDoc2
Dim swTable As SldWorks.ITableAnnotation
Dim status As Boolean
Dim caminhoDoDesenho As String
Dim nomeDoDesenho As String
Dim bomPath As String

Option Explicit

Public Sub Main()
    Set swApp = Application.SldWorks

    Set swModDoc = swApp.ActiveDoc
    Dim swSM As ISelectionMgr
    Set swSM = swModDoc.SelectionManager
    Set swTable = swSM.GetSelectedObject6(1, -1)
    ' swModDoc.ClearSelection2 (True)

    Dim swSpecTable As IBomTableAnnotation
    Set swSpecTable = swTable
    
    ' Set swSelectionManager = swModel.SelectionManager
    ' get the count of selected objects
    Dim Count As Long
    Count = swSM.GetSelectedObjectCount2(-1)
    'if the user has made no selection then exit
    If Count = 0 Then
    swApp.SendMsgToUser "Vocï¿½ nï¿½o selecionou nenhuma BOM!"
    Exit Sub
    End If
    
    ' Obter o caminho e nome do arquivo do desenho aberto
    caminhoDoDesenho = swModDoc.GetPathName()
    nomeDoDesenho = ExtrairNomeDoArquivo(caminhoDoDesenho)
    
    Call DefinirVariavelAmbiente(nomeDoDesenho)
    
    bomPath = Environ("TEMP") & "\" & nomeDoDesenho & ".xlsx"
    
    DeleteFileIfExists bomPath

    ' Save the selected BOM table to Microsoft Excel, including hidden cells and images
    status = swSpecTable.SaveAsExcel(bomPath, False, False)

    ' Release the reference to the TableAnnotation object
    Set swTable = Nothing
    
    While Dir(bomPath) = ""
    Wend
    
    Call ExecutarScriptPython

End Sub
Function ExtrairNomeDoArquivo(caminhoDoArquivo As String) As String
    Dim caminho As String
    Dim nomeDoArquivo As String
    Dim posicaoDaUltimaBarra As Integer
    Dim posicaoDoPonto As Integer

    ' Defina a string do caminho do arquivo
    caminho = caminhoDoArquivo

    ' Encontre a posiï¿½ï¿½o da ï¿½ltima barra e do ponto na string
    posicaoDaUltimaBarra = InStrRev(caminho, "\")
    posicaoDoPonto = InStrRev(caminho, ".")

    ' Extraia a substring desejada
    nomeDoArquivo = Mid(caminho, posicaoDaUltimaBarra + 1, posicaoDoPonto - posicaoDaUltimaBarra - 1)

    ' Imprima o nome do arquivo no Immediate Window
    ' Debug.Print "O nome do arquivo extraï¿½do ï¿½: " & nomeDoArquivo
    
    ExtrairNomeDoArquivo = nomeDoArquivo
End Function
Sub ExecutarScriptPython()
    Dim CaminhoArquivo As String
    
    ' Substitua o caminho abaixo pelo caminho do seu arquivo botao-salvar-bom-solidworks-totvs
    CaminhoArquivo = "\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\VBA\PYTHON\botao-salvar-bom-solidworks-totvs.pyw"
    
    ' Use a funï¿½ï¿½o Shell para abrir o arquivo com a aplicaï¿½ï¿½o padrï¿½o
    Shell "explorer.exe """ & CaminhoArquivo & """", vbNormalFocus
End Sub
Sub DefinirVariavelAmbiente(codigoDesenho As String)

    ' Substitua "meuarquivo" pelo nome real do seu arquivo (sem a extensï¿½o)
    Dim nomeArquivo As String
    nomeArquivo = codigoDesenho

    ' Construa o comando para definir a variï¿½vel de ambiente
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
    Else
        ' File does not exist
        ' MsgBox "File does not exist: " & filePath
    End If
End Sub
