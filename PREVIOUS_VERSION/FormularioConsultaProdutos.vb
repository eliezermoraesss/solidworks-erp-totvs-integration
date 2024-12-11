Dim swApp As Object
Sub main()
    Call AbrirArquivoExternoGenerico
Set swApp = Application.SldWorks
End Sub
Sub AbrirArquivoExternoGenerico()
    Dim objShell As Object
    Dim filePath As String

    ' Especifique o caminho completo do arquivo
    filePath = "\\192.175.175.4\f\INTEGRANTES\ELIEZER\PROJETO SOLIDWORKS TOTVS\VBA\PYTHON\python-test_013_PESQUISA_PYQT5.pyw" ' Substitua pelo caminho do seu arquivo

    ' Inicialize o objeto Shell
    Set objShell = CreateObject("Shell.Application")

    ' Abra o arquivo
    objShell.Open (filePath)
End Sub
