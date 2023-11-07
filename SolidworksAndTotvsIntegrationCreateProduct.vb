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
    Dim incrementoChavePrimariaBanco As Long
    Dim respostaSimOuNaoAlterarCadastro As VbMsgBoxResult
    
Sub main()

baseDeDados = "PROTHEUS12_R27" ' PROTHEUS12_R27 => PRODUCAO PROTHEUS1233_HML => TESTE
incrementoChavePrimariaBanco = 1

Call testarConexaoAoSQLServer
Call buscarValoresNoFormularioPropriedadesPersonalizadasSW
Call validacaoDeDadosDoFormularioDeCadastro

If editarCodigo = "Unchecked" Then
    Call verificarSeProdutoJaEstaCadastrado(codigo)
    If produtoJaCadastrado = False Then
        Call consultarGrupoPeloCodigoRetornarDescricao(grupo(0))
        Call cadastrarProdutoNoTotvs(codigo)
    Else
        Call exibirJanelaDePerguntaParaAlterarProduto(codigo)
        If respostaSimOuNaoAlterarCadastro = vbYes Then
            Call consultarGrupoPeloCodigoRetornarDescricao(grupo(0))
            Call alterarProdutoNoTotvs(codigo)
        Else
            End
    End If
    End If
Else
    Call verificarSeProdutoJaEstaCadastrado(caixaTextoCodigo)
    If produtoJaCadastrado = False Then
        Call consultarGrupoPeloCodigoRetornarDescricao(grupo(0))
        Call cadastrarProdutoNoTotvs(caixaTextoCodigo)
    Else
        Call exibirJanelaDePerguntaParaAlterarProduto(caixaTextoCodigo)
        If respostaSimOuNaoAlterarCadastro = vbYes Then
            Call consultarGrupoPeloCodigoRetornarDescricao(grupo(0))
            Call alterarProdutoNoTotvs(caixaTextoCodigo)
        Else
            End
    End If
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
       
        ' Feche a conexão quando terminar
        cn.Close
        Set cn = Nothing
    Else
        MsgBox "Falha na conexão ao banco de dados TOTVS.", vbCritical, "CADASTRO TOTVS"
    End If
    
End Sub
Sub cadastrarProdutoNoTotvs(codigoProduto As String)

Dim sqlPart1 As String
Dim sqlPart2 As String
Dim sqlPart3 As String
Dim sqlPart4 As String
Dim sqlPart5 As String
Dim sqlPart6 As String
Dim sqlPart7 As String

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

        Dim strSQLCount As String
        strSQLCount = "SELECT TOP 1 R_E_C_N_O_ FROM " & baseDeDados & ".dbo.SB1010 ORDER BY R_E_C_N_O_ DESC;"
        
        Dim rs As Object
        Set rs = cn.Execute(strSQLCount)
        
        ' Verifique se o resultado foi obtido com sucesso
        If Not rs.EOF Then
            Dim resultado As Long
            resultado = rs.Fields(0).Value
            
            ' Feche o recordset
            rs.Close
            
            Dim novoResultado As Long
            novoResultado = resultado + incrementoChavePrimariaBanco ' *******************************
            
            Dim strSQL As String
                                   
sqlPart1 = "INSERT INTO " & baseDeDados & ".dbo.SB1010 (B1_AFAMAD, B1_FILIAL, B1_COD, B1_DESC, B1_XDESC2, B1_CODITE, B1_TIPO, B1_UM, B1_LOCPAD, B1_GRUPO, B1_ZZLOCAL, B1_POSIPI, B1_ESPECIE, B1_EX_NCM, B1_EX_NBM, B1_PICM, B1_IPI, B1_ALIQISS, B1_CODISS, B1_TE, B1_TS, B1_PICMRET, B1_BITMAP, B1_SEGUM, B1_PICMENT, B1_IMPZFRC, B1_CONV, B1_TIPCONV, B1_ALTER, B1_QE, B1_PRV1, B1_EMIN, B1_CUSTD, B1_UCALSTD, B1_UCOM, B1_UPRC, B1_MCUSTD, B1_ESTFOR, B1_PESO, B1_ESTSEG, B1_FORPRZ, B1_PE, B1_TIPE, B1_LE, B1_LM, B1_CONTA, B1_TOLER, B1_CC, B1_ITEMCC, B1_PROC, B1_LOJPROC, B1_FAMILIA, B1_QB, B1_APROPRI, B1_TIPODEC, B1_ORIGEM, B1_CLASFIS, B1_UREV, B1_DATREF, B1_FANTASM, B1_RASTRO, B1_FORAEST, B1_COMIS, B1_DTREFP1, B1_MONO, B1_PERINV, B1_GRTRIB, B1_MRP, B1_NOTAMIN, B1_CONINI, B1_CONTSOC, B1_PRVALID, B1_CODBAR, B1_GRADE, B1_NUMCOP, B1_FORMLOT, B1_IRRF, B1_FPCOD, B1_CODGTIN, B1_DESC_P, B1_CONTRAT, B1_DESC_GI, B1_DESC_I, B1_LOCALIZ, B1_OPERPAD, B1_ANUENTE, B1_OPC, B1_CODOBS, B1_VLREFUS, B1_IMPORT, B1_FABRIC, B1_SITPROD, "
sqlPart2 = "B1_MODELO, B1_SETOR, B1_PRODPAI, B1_BALANCA, B1_TECLA, B1_DESPIMP, B1_TIPOCQ, B1_SOLICIT, B1_GRUPCOM, B1_QUADPRO, B1_BASE3, B1_DESBSE3, B1_AGREGCU, B1_NUMCQPR, B1_CONTCQP, B1_REVATU, B1_CODEMB, B1_INSS, B1_ESPECIF, B1_NALNCCA, B1_MAT_PRI, B1_NALSH, B1_REDINSS, B1_REDIRRF, B1_ALADI, B1_TAB_IPI, B1_GRUDES, B1_DATASUB, B1_REDPIS, B1_REDCOF, B1_PCSLL, B1_PCOFINS, B1_PPIS, B1_MTBF, B1_MTTR, B1_FLAGSUG, B1_CLASSVE, B1_MIDIA, B1_QTMIDIA, B1_QTDSER, B1_VLR_IPI, B1_ENVOBR, B1_SERIE, B1_FAIXAS, B1_NROPAG, B1_ISBN, B1_TITORIG, B1_LINGUA, B1_EDICAO, B1_OBSISBN, B1_CLVL, B1_ATIVO, B1_EMAX, B1_PESBRU, B1_TIPCAR, B1_FRACPER, B1_VLR_ICM, B1_INT_ICM, B1_CORPRI, B1_CORSEC, B1_NICONE, B1_ATRIB1, B1_ATRIB2, B1_ATRIB3, B1_REGSEQ, B1_VLRSELO, B1_CODNOR, B1_CPOTENC, B1_POTENCI, B1_REQUIS, B1_SELO, B1_LOTVEN, B1_OK, B1_USAFEFO, B1_QTDACUM, B1_QTDINIC, B1_CNATREC, B1_TNATREC, B1_AFASEMT, B1_AIMAMT, B1_TERUM, B1_AFUNDES, B1_CEST, B1_GRPCST, B1_IAT, B1_IPPT, B1_GRPNATR, B1_DTFIMNT, B1_DTCORTE, B1_FECP, B1_MARKUP, "
sqlPart3 = "B1_CODPROC, B1_LOTESBP, B1_QBP, B1_VALEPRE, B1_CODQAD, B1_AFABOV, B1_VIGENC, B1_VEREAN, B1_DIFCNAE, B1_ESCRIPI, B1_PMACNUT, B1_PMICNUT, B1_INTEG, B1_HREXPO, B1_CRICMS, B1_REFBAS, B1_MOPC, B1_USERLGI, B1_USERLGA, B1_UMOEC, B1_UVLRC, B1_PIS, B1_GCCUSTO, B1_CCCUSTO, B1_TALLA, B1_PARCEI, B1_GDODIF, B1_VLR_PIS, B1_TIPOBN, B1_TPREG, B1_MSBLQL, B1_VLCIF, B1_DCRE, B1_DCR, B1_DCRII, B1_TPPROD, B1_DCI, B1_COEFDCR, B1_CHASSI, B1_CLASSE, B1_FUSTF, B1_GRPTI, B1_PRDORI, B1_APOPRO, B1_PRODREC, B1_ALFECOP, B1_ALFECST, B1_CFEMA, B1_FECPBA, B1_MSEXP, B1_PAFMD5, B1_PRODSBP, B1_CODANT, B1_IDHIST, B1_CRDEST, B1_REGRISS, B1_FETHAB, B1_ESTRORI, B1_CALCFET, B1_PAUTFET, B1_CARGAE, B1_PRN944I, B1_ALFUMAC, B1_PRINCMG, B1_PR43080, B1_RICM65, B1_SELOEN, B1_TRIBMUN, B1_RPRODEP, B1_FRETISS, B1_AFETHAB, B1_DESBSE2, B1_BASE2, B1_VLR_COF, B1_PRFDSUL, B1_TIPVEC, B1_COLOR, B1_RETOPER, B1_COFINS, B1_CSLL, B1_CNAE, B1_ADMIN, B1_AFACS, B1_AJUDIF, B1_ALFECRN, B1_CFEM, B1_CFEMS, B1_MEPLES, B1_REGESIM, B1_RSATIVO, B1_TFETHAB, "
sqlPart4 = "B1_TPDP, B1_CRDPRES, B1_CRICMST, B1_FECOP, B1_CODLAN, B1_GARANT, B1_PERGART, B1_SITTRIB, B1_PORCPRL, B1_IMPNCM, B1_IVAAJU, B1_BASE, B1_ZZCODAN, B1_ZZNOGRP, B1_ZZOBS1, B1_XFORDEN, B1_ZZMEN1, B1_ZZLEGIS, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) VALUES(0.0, N'    ', N'" & codigoProduto & "', N'" & descricaoProduto & "', N'" & descricaoProduto2 & "', N'                           ', N'" & tipo & "', N'" & unidade & "', N'" & armazem(0) & "', N'" & grupo(0) & "', N'      ', N'          ', N'  ', N'   ', N'   ', 0.0, 0.0, 0.0, N'         ', N'   ', N'   ', 0.0, N'                    ', N'  ', 0.0, N' ', 0.0, N'M', N'               ', 0.0, 0.0, 0.0, 0.0, N'        ', N'        ', 0.0, N'1', N'   ', 0.0, 0.0, N'   ', 0.0, N' ', 0.0, 0.0, N'                    ', 0.0, N'" & centroDeCusto(0) & "', N'         ', N'      ', N'  ', N' ', 1.0, N' ', N'N', N' ', N'  ', N'" & dataCadastro & "', N'" & dataCadastro & "', N' ', N'N', N' ', 0.0, N'        ', N' ', 0.0, N'      ', N'S', 0.0, "
sqlPart5 = "N'        ', N' ', 0.0, N'               ', N' ', 0.0, N'   ', N' ', N'          ', N'               ', N'      ', N'N', N'      ', N'      ', N'N', N'  ', N'2', N'                                                                                ', N'      ', 0.0, N'N', N'                    ', N'  ', N'               ', N'  ', N'               ', N' ', N'   ', N'N', N'M', N'N', N'      ', N' ', N'              ', N'                                                            ', N'2', 0.0, 0.0, N'001', N'                              ', N'N', N'                                                                                ', N'       ', N'                    ', N'        ', 0.0, 0.0, N'   ', N'  ', N'   ', N'        ', 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, N'1', N'1', N'2', 0.0, N'1', 0.0, N'0', N'                    ', 0.0, 0.0, N'          ', N'                                                  ', N'                    ', N'   ', "
sqlPart6 = "N'                                        ', N'         ', N'S', 0.0, 0.0, N'      ', 0.0, 0.0, 0.0, N'      ', N'      ', N'               ', N'      ', N'      ', N'      ', N'      ', 0.0, N'   ', N'2', 0.0, N' ', N' ', 0.0, N'    ', N'1', 0.0, 0.0, N'   ', N'    ', 0.0, 0.0, N'  ', 0.0, N'         ', N'   ', N' ', N' ', N'  ', N'        ', N'        ', 0.0, 0.0, N'      ', 0.0, 0.0, N' ', N'                      ', 0.0, N'        ', N'  ', N'           ', N'3', 0.0, 0.0, N' ', N'        ', N'0', N' ', NULL, N' 0#  0@< 50A 80; ', N' 0#  0@< 50A 80; ', 0.0, 0.0, N'2', N'        ', N'         ', N'      ', N'      ', N' ', 0.0, N'  ', N' ', N'2', 0.0, N'          ', N'         ', 0.0, N'  ', N' ', 0.0, N'                         ', N'      ', N' ', N'    ', N'               ', N' ', N' ', 0.0, 0.0, 0.0, 0.0, N'        ', N'                                ', N'C', N'               ', N'                    ', 0.0, "
sqlPart7 = "N'  ', N'N', N'               ', N' ', 0.0, N' ', N'S', 0.0, 0.0, 0.0, N'2', N'      ', N'                    ', N' ', N' ', 0.0, N'                                                            ', N'              ', 0.0, 0.0, N'      ', N'          ', N'2', N'2', N'2', N'         ', N'          ', 0.0, N' ', 0.0, N' ', N' ', N' ', N' ', N' ', N' ', N' ', 0.0, N' ', N' ', N'      ', N'2', 0.0, N' ', N'  ', 0.0, N' ', N'              ', N'               ', N'" & descricaoGrupo & "', NULL, N' ', N'   ', N'                                                                                                                                                                                                                                                          ', N' ', N'" & novoResultado & "', 0);"
     
      strSQL = sqlPart1 & sqlPart2 & sqlPart3 & sqlPart4 & sqlPart5 & sqlPart6 & sqlPart7
            
            ' Execute o comando SQL
            cn.Execute strSQL
            
            MsgBox "Produto cadastrado com sucesso!" & vbNewLine & vbNewLine & codigoProduto & " - " & descricaoProduto & descricaoProduto2, vbInformation, "CADASTRO TOTVS"
            
            ' Feche a conexão quando terminar
            cn.Close
            Set cn = Nothing
        End If
    Else
        MsgBox "Falha na conexão ao banco de dados TOTVS.", vbCritical, "CADASTRO TOTVS"
        
    End If
End Sub
Sub alterarProdutoNoTotvs(codigoProduto As String)

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
     
            ' Defina o comando SQL UPDATE
            Dim strSQL As String
                                     
            strSQL = "UPDATE " & baseDeDados & ".dbo.SB1010 SET B1_DESC = N'" & descricaoProduto & "', B1_XDESC2 = N'" & descricaoProduto2 & "', B1_TIPO = N'" & tipo & "', B1_UM = N'" & unidade & "', B1_LOCPAD = N'" & armazem(0) & "', B1_GRUPO = N'" & grupo(0) & "', B1_ZZNOGRP = N'" & descricaoGrupo & "', B1_CC = N'" & centroDeCusto(0) & "' WHERE B1_COD = N'" & codigoProduto & "';"
            
            ' Execute o comando SQL
            cn.Execute strSQL
            
            MsgBox "Produto alterado com sucesso!" & vbNewLine & vbNewLine & codigoProduto & " - " & descricaoProduto & descricaoProduto2, vbInformation, "CADASTRO TOTVS"
            
            ' Feche a conexão quando terminar
            cn.Close
            Set cn = Nothing
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
                    Case "caixaTextoCodigo"
                        caixaTextoCodigo = CStr(vConfValue(j))
                    Case "editarCodigo"
                        editarCodigo = CStr(vConfValue(j))
                    Case "descricaoProduto"
                        descricaoProduto = UCase(CStr(vConfValue(j)))
                    Case "descricaoProduto2"
                        descricaoProduto2 = UCase(CStr(vConfValue(j)))
                    Case "tipo"
                        tipo = CStr(vConfValue(j))
                    Case "unidade"
                        unidade = CStr(vConfValue(j))
                    Case "armazem"
                        armazem = Split(CStr(vConfValue(j)), " ")
                    Case "grupo"
                        grupo = Split(CStr(vConfValue(j)), " ")
                    Case "centroDeCusto"
                        centroDeCusto = Split(CStr(vConfValue(j)), " ")
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
    MsgBox "Por favor preencher TODOS os campos obrigatórios (*) do formulário de cadastro!", vbExclamation, "CADASTRO TOTVS"
    End
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
            produtoJaCadastrado = True
        Else
        produtoJaCadastrado = False
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
    Dim descricao1Local As String
    Dim descricao2Local As String
    Dim descricaoTemCaracteresEspeciais As Boolean
    Dim descricaoUmMaisDescricaoComplementar As String
    
    quantidadeMaxDescricao1 = 100 ' quantidade de caracteres do campo desc. 1
    quantidadeMaxDescricao2 = 60 ' quantidade de caracteres do campo desc. 2
    
    descricaoUmMaisDescricaoComplementar = descricao1 & descricao2
    
    descricaoTemCaracteresEspeciais = HasSpecialCharacters(descricaoUmMaisDescricaoComplementar)
    
    If descricaoTemCaracteresEspeciais = True Then
        MsgBox "OPS!" & vbNewLine & vbNewLine & "O campo DESCRIÇÃO contém caracteres especiais." & vbNewLine & "Remova os caracteres especiais e tente novamente! =)" & vbNewLine & vbNewLine & "Caracteres especiais permitidos: . / - = ( )", vbExclamation, "CADASTRO TOTVS"
        End
    End If
    
On Error GoTo ErrorHandler

If Len(descricaoProduto) > 100 And Len(descricaoProduto2) <= 60 Then
        Err.Raise 9999, Description:="O campo DESCRIÇÃO excedeu o limite de caracteres."
    ElseIf Len(descricaoProduto2) > 60 And Len(descricaoProduto) <= 100 Then
        Err.Raise 9999, Description:="O campo DESCRIÇÃO COMPLEMENTAR excedeu o limite de caracteres."
    ElseIf Len(descricaoProduto2) > 60 And Len(descricaoProduto) > 100 Then
        Err.Raise 9999, Description:="Os campos DESCRIÇÃO e DESCRIÇÃO COMPLEMENTAR excederam o limite de caracteres."
    Else
        tamanhoDescricao1 = Len(descricao1)
        descricaoProduto = descricao1 & Space(quantidadeMaxDescricao1 - tamanhoDescricao1)
        
        tamanhoDescricao2 = Len(descricao2)
        descricaoProduto2 = descricao2 & Space(quantidadeMaxDescricao2 - tamanhoDescricao2)
    End If
    Exit Sub
    
ErrorHandler:
        MsgBox Err.Description, vbExclamation, "CADASTRO TOTVS"
        End
End Sub
Sub FormatarDataAtual()
    Dim dataAtual As Date
    Dim dataString As String
    dataAtual = Date ' Obtém a data atual
    dataString = Format(dataAtual, "yyyymmdd") ' Formata a data como string no formato desejado
    
    dataCadastro = dataString
End Sub

Sub exibirJanelaDePerguntaParaAlterarProduto(codigoMensagem As String)
Dim textoCorpo As String
textoCorpo = "Já existe um produto cadastrado com o código " & codigoMensagem & vbNewLine & vbNewLine & "Você TEM CERTEZA que deseja alterar os dados de cadastro deste produto?" & vbNewLine & vbNewLine & "SIM - Sobrescrever os dados do produto atual no TOTVS pelos dados atuais do formulário de cadastro." & vbNewLine & vbNewLine & "NÃO - Os dados de cadastro do produto não são alterados."

respostaSimOuNaoAlterarCadastro = MsgBox(textoCorpo, vbYesNo + vbExclamation, "CADASTRO TOTVS")

End Sub
Function HasSpecialCharacters(inputString As String) As Boolean
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")

    regEx.IgnoreCase = True
    regEx.Global = True
    regEx.Pattern = "[~!@#$Ø%^&*+\[\]{};:'"",<>?\\|çãÇÃáÁÀàíÍêÊóÓúÚõÕ]"

    HasSpecialCharacters = regEx.Test(inputString)
End Function
