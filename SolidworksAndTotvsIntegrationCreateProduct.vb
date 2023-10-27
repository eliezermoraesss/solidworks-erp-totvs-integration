Dim swApp As Object
    Dim codigo As String
    Dim caixaTextoCodigo As String
    Dim descricaoProduto As String
    Dim descricaoProduto2 As String
    Dim tipo As String
    Dim unidade As String
    Dim editarCodigo As String
    Dim armazem() As String
    Dim grupo() As String
    Dim centroDeCusto() As String
    Dim radioRevisao As String
    Dim partes() As String
    Dim descricaoGrupo As String
    Dim produtoJaCadastrado As Boolean
    Dim baseDeDados As String
    Dim dataCadastro As String
    Dim variavelControleBanco As Long
Sub main()

baseDeDados = "PROTHEUS1233_HML" ' PROTHEUS12_R27 => PRODUCAO PROTHEUS1233_HML => TESTE
variavelControleBanco = 2 ' 2 => AMBIENTE DE PRODUCAO / 1 => AMBIENTE DE TESTES

Call testarConexaoAoSQLServer
Call buscarValoresNoFormularioPropriedadesPersonalizadasSW
Call validacaoDeDadosDoFormularioDeCadastro

If editarCodigo = "Unchecked" Then
    Call verificarSeProdutoJaEstaCadastrado(codigo)
    If produtoJaCadastrado = False Then
        Call consultarGrupoPeloCodigoRetornarDescricao(grupo(0))
        Call cadastrarProdutoNoTotvs(codigo)
    End If
Else
    Call verificarSeProdutoJaEstaCadastrado(caixaTextoCodigo)
    If produtoJaCadastrado = False Then
        Call consultarGrupoPeloCodigoRetornarDescricao(grupo(0))
        Call cadastrarProdutoNoTotvs(caixaTextoCodigo)
    End If
End If

Set swApp = Application.SldWorks
End Sub
Sub testarConexaoAoSQLServer()
    ' Defina as variáveis de conexão
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    
    ' Defina a string de conexão com o SQL Server
    Dim strCon As String
    strCon = "Provider=SQLOLEDB;Data Source=SVRERP;Initial Catalog=PROTHEUS12_R27;User ID=coognicao;Password=0705@Abc;"
    
    ' Tente abrir a conexão
    On Error Resume Next
    cn.Open strCon
    On Error GoTo 0 ' Restaura o tratamento de erros normal
    
    ' Verifique se a conexão foi aberta com sucesso
    If cn.State = 1 Then ' 1 indica que a conexão está aberta
        ' MsgBox "Conexão bem-sucedida ao banco de dados!"
       
        ' Feche a conexão quando terminar
        cn.Close
        Set cn = Nothing
    Else
        MsgBox "Falha na conexão ao banco de dados TOTVS.", vbCritical, "CADASTRO TOTVS"
    End If
    
End Sub
Sub cadastrarProdutoNoTotvs(codigoProduto As String)

    ' Defina as variáveis de conexão
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    
    ' Defina a string de conexão com o SQL Server
    Dim strCon As String
    strCon = "Provider=SQLOLEDB;Data Source=SVRERP;Initial Catalog=PROTHEUS12_R27;User ID=coognicao;Password=0705@Abc;"
    
    ' Tente abrir a conexão
    On Error Resume Next
    cn.Open strCon
    On Error GoTo 0 ' Restaura o tratamento de erros normal
    
    ' Verifique se a conexão foi aberta com sucesso
    If cn.State = 1 Then ' 1 indica que a conexão está aberta
    
        ' Defina o comando SQL COUNT para obter o número de registros na tabela
        Dim strSQLCount As String
        strSQLCount = "SELECT COUNT(*) FROM " & baseDeDados & ".dbo.SB1010"
        
        ' Execute o comando SQL COUNT e obtenha o resultado
        Dim rs As Object
        Set rs = cn.Execute(strSQLCount)
        
        ' Verifique se o resultado foi obtido com sucesso
        If Not rs.EOF Then
            ' Obtém o valor COUNT retornado (primeira coluna da primeira linha)
            Dim resultado As Long
            resultado = rs.Fields(0).Value
            
            ' Feche o recordset
            rs.Close
            
            ' Agora, você pode calcular o valor a ser inserido na consulta SQL
            Dim novoResultado As Long
            novoResultado = resultado + variavelControleBanco ' *******************************
            
            ' Defina o comando SQL INSERT INTO com o novoResultado
            Dim strSQL As String
                                   
            strSQL = "INSERT INTO " & baseDeDados & ".dbo.SB1010 " & _
                "(B1_COD, B1_DESC, B1_XDESC2, B1_TIPO, B1_UM, B1_LOCPAD, B1_GRUPO, B1_ZZNOGRP, B1_CC, B1_LOCALIZ, B1_GARANT, B1_REVATU, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) " & _
                "VALUES (N'" & codigoProduto & "', N'" & descricaoProduto & "', N'" & descricaoProduto2 & "', N'" & tipo & "', N'" & unidade & "', N'" & armazem(0) & "', N'" & grupo(0) & "', N'" & descricaoGrupo & "', N'" & centroDeCusto(0) & "', N'N' , N'2' , N'001', N' ', N'" & novoResultado & "', 0);"
            
            ' Execute o comando SQL
            cn.Execute strSQL
            
            MsgBox "Produto cadastrado com sucesso!" & vbNewLine & vbNewLine & codigoProduto & " - " & descricaoProduto, vbInformation, "CADASTRO TOTVS"
            
            ' Feche a conexão quando terminar
            cn.Close
            Set cn = Nothing
        End If
    Else
        MsgBox "Falha na conexão ao banco de dados TOTVS.", vbCritical, "CADASTRO TOTVS"
        
    End If
End Sub
Sub buscarValoresNoFormularioPropriedadesPersonalizadasSW()

    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swConfigMgr As SldWorks.ConfigurationManager
    Dim vConfName As Variant
    Dim vConfParam As Variant
    Dim vConfValue As Variant
    Dim i As Long
    Dim j As Long
    Dim bRet As Boolean
    Dim texto As String
    
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    Set swConfigMgr = swModel.ConfigurationManager
    Debug.Print "File = " + swModel.GetPathName
    vConfName = swModel.GetConfigurationNames
    For i = 0 To UBound(vConfName)
        bRet = swConfigMgr.GetConfigurationParams(vConfName(i), vConfParam, vConfValue)
        Debug.Assert bRet
        If Not IsEmpty(vConfParam) Then
            For j = 15 To UBound(vConfParam)
                Debug.Print "    " & vConfParam(j) & " = " & vConfValue(j)
                
                texto = CStr(vConfParam(j))
                partes = Split(texto, "@")
                          
                If InStr(texto, "@") > 0 Then
                partes = Split(texto, "@")
                
                Select Case partes(1)
                    Case "codigo"
                        codigo = CStr(vConfValue(j))
                        ' MsgBox "A segunda parte é 'codigo' e o valor é " & vConfValue(j)
                    Case "caixaTextoCodigo"
                        caixaTextoCodigo = CStr(vConfValue(j))
                        ' MsgBox "A segunda parte é 'codigo' e o valor é " & vConfValue(j)
                    Case "editarCodigo"
                        editarCodigo = CStr(vConfValue(j))
                        ' MsgBox "A segunda parte é 'codigo' e o valor é " & vConfValue(j)
                    Case "descricaoProduto"
                        descricaoProduto = CStr(vConfValue(j))
                        ' MsgBox "A segunda parte é 'descricaoProduto' e o valor é " & vConfValue(j)
                    Case "descricaoProduto2"
                        descricaoProduto2 = CStr(vConfValue(j))
                        ' MsgBox "A segunda parte é 'descricaoProduto' e o valor é " & vConfValue(j)
                    Case "tipo"
                        tipo = CStr(vConfValue(j))
                        ' MsgBox "A segunda parte é 'tipo' e o valor é " & vConfValue(j)
                    Case "unidade"
                        unidade = CStr(vConfValue(j))
                        ' MsgBox "A segunda parte é 'unidade' e o valor é " & vConfValue(j)
                    Case "armazem"
                        armazem = Split(CStr(vConfValue(j)), " ")
                        ' MsgBox "A segunda parte é 'armazem' e o valor é " & armazem(0)
                    Case "grupo"
                        grupo = Split(CStr(vConfValue(j)), " ")
                        ' MsgBox "A segunda parte é 'grupo' e o valor é " & grupo(0)
                    Case "centroDeCusto"
                        centroDeCusto = Split(CStr(vConfValue(j)), " ")
                        ' MsgBox "A segunda parte é 'centro de custo' e o valor é " & centroDeCusto(0)
                    Case Else
                        End Select
                        
                    End If
            Next j
        End If
    Next i
    Call FormatarDataAtual
End Sub
Sub validacaoDeDadosDoFormularioDeCadastro()

Call tratamentoDosCamposDescricao(descricaoProduto, descricaoProduto2)

   On Error GoTo ErrorHandler
    If Len(Trim(caixaTextoCodigo)) = 0 And editarCodigo = "Checked" Or Len(Trim(descricaoProduto)) = 0 Or tipo = "" Or unidade = "" Or armazem(0) = "" Or grupo(0) = "" Or centroDeCusto(0) = "" Then
        Err.Raise 9999
    End If
Exit Sub

ErrorHandler:
    MsgBox "Preencha TODOS os campos do formulário de cadastro!", vbExclamation, "CADASTRO TOTVS"
End Sub
Sub verificarSeProdutoJaEstaCadastrado(codigo As String)
    ' Defina as variáveis de conexão
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    
    ' Defina a string de conexão com o SQL Server
    Dim strCon As String
    strCon = "Provider=SQLOLEDB;Data Source=SVRERP;Initial Catalog=PROTHEUS12_R27;User ID=coognicao;Password=0705@Abc;"
    
    ' Tente abrir a conexão
    On Error Resume Next
    cn.Open strCon
    On Error GoTo 0 ' Restaura o tratamento de erros normal
    
    ' Verifique se a conexão foi aberta com sucesso
    If cn.State = 1 Then ' 1 indica que a conexão está aberta
    
        Dim strSQL As String
        strSQL = "SELECT B1_COD FROM " & baseDeDados & ".dbo.SB1010 s WHERE B1_COD = '" & codigo & "';"
        
        ' Execute o comando SQL COUNT e obtenha o resultado
        Dim rs As Object
        Set rs = cn.Execute(strSQL)
        
        ' Verifique se o resultado foi obtido com sucesso
        If Not rs.EOF Then
            ' Se o produto estiver cadastrado, faça algo aqui
            produtoJaCadastrado = True
            MsgBox "Já existe um produto cadastrado com o código " & codigo, vbExclamation, "CADASTRO TOTVS"
        Else
        produtoJaCadastrado = False
            ' Se o produto não estiver cadastrado, faça algo aqui
            ' MsgBox "O produto não está cadastrado."
        End If

        ' Feche o recordset
        rs.Close
        
        ' Feche a conexão quando terminar
        cn.Close
        Set cn = Nothing
    Else
        MsgBox "Falha na conexão ao banco de dados TOTVS.", vbCritical, "CADASTRO TOTVS"
    End If
End Sub
Sub consultarGrupoPeloCodigoRetornarDescricao(grupo As String)
    ' Defina as variáveis de conexão
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    
    ' Defina a string de conexão com o SQL Server
    Dim strCon As String
    strCon = "Provider=SQLOLEDB;Data Source=SVRERP;Initial Catalog=PROTHEUS12_R27;User ID=coognicao;Password=0705@Abc;"
    
    ' Tente abrir a conexão
    On Error Resume Next
    cn.Open strCon
    On Error GoTo 0 ' Restaura o tratamento de erros normal
    
    ' Verifique se a conexão foi aberta com sucesso
    If cn.State = 1 Then ' 1 indica que a conexão está aberta
    
        Dim strSQL As String
        strSQL = "SELECT BM_DESC FROM " & baseDeDados & ".dbo.SBM010 WHERE BM_GRUPO = '" & grupo & "';"
        
        ' Execute o comando SQL COUNT e obtenha o resultado
        Dim rs As Object
        Set rs = cn.Execute(strSQL)
        
        ' Verifique se o resultado foi obtido com sucesso
        If Not rs.EOF Then
        Dim resultado As String
            resultado = rs.Fields(0).Value
            descricaoGrupo = resultado
        Else
            ' MsgBox "Grupo não localizado."
        End If

        ' Feche o recordset
        rs.Close
        
        ' Feche a conexão quando terminar
        cn.Close
        Set cn = Nothing
    Else
        MsgBox "Falha na conexão ao banco de dados TOTVS.", vbCritical, "CADASTRO TOTVS"
    End If
End Sub

Sub tratamentoDosCamposDescricao(descricao1 As String, descricao2 As String)

    Dim tamanhoDescricao1 As Long
    Dim tamanhoDescricao2 As Long
    Dim quantidadeMaxDescricao1 As Long
    Dim quantidadeMaxDescricao2 As Long
    
    quantidadeMaxDescricao1 = 100 ' quantidade de caracteres do campo desc. 1
    quantidadeMaxDescricao2 = 60 ' quantidade de caracteres do campo desc. 2
	
	    ' MsgBox Len(descricaoProduto) & " - " & Len(descricaoProduto2)
    
    tamanhoDescricao1 = Len(descricao1)
    descricaoProduto = descricao1 & Space(quantidadeMaxDescricao1 - tamanhoDescricao1)
    
    tamanhoDescricao2 = Len(descricao2)
    descricaoProduto2 = descricao2 & Space(quantidadeMaxDescricao2 - tamanhoDescricao2)
	
	On Error GoTo ErrorHandler
		
	If Len(descricaoProduto) > 100 And Len(descricaoProduto2) < 60 Then
		Err.Raise 9999, Description:= "O campo DESCRIÇÃO 1 excedeu o limite de caracteres"
	Else If Len(descricaoProduto2) > 60 And Len(descricaoProduto2) < 100 Then
		Err.Raise 9999, Description:= "O campo DESCRIÇÃO 2 excedeu o limite de caracteres"
	Else If Len(descricaoProduto2) > 60 And Len(descricaoProduto2) > 100 Then
		Err.Raise 9999, Description:= "Os campos DESCRIÇÃO 1 e 2 excederam o limite de caracteres"
	End If
	
	On Error GoTo 0
	
	ErrorHandler:
		MsgBox Err.Description

End Sub
Sub FormatarDataAtual()
    Dim dataAtual As Date
    Dim dataString As String
    dataAtual = Date ' Obtém a data atual
    dataString = Format(dataAtual, "yyyymmdd") ' Formata a data como string no formato desejado
    
    dataCadastro = dataString
    ' Use dataString para salvar no banco de dados ou para outros fins
    ' Exemplo de uso para exibição em uma caixa de mensagem:
    ' MsgBox "Data formatada: " & dataCadastro
End Sub
