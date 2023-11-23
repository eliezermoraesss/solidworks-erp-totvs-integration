SELECT BM_FILIAL, BM_GRUPO, BM_DESC, BM_PICPAD, BM_PROORI, BM_CODMAR, BM_STATUS, BM_GRUREL, BM_TIPGRU, BM_MARKUP, BM_PRECO, BM_MARGPRE, BM_LENREL, BM_TIPMOV, BM_CODGRT, BM_CLASGRU, BM_FORMUL, BM_TPSEGP, BM_DTUMOV, BM_HRUMOV, BM_CONC, BM_CORP, BM_EVENTO, BM_LAZER, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_, BM_USERLGI, BM_USERLGA
FROM PROTHEUS1233_HML.dbo.SBM010;

SELECT TOP 1 R_E_C_N_O_ FROM PROTHEUS1233_HML.dbo.SBM010 ORDER BY R_E_C_N_O_ DESC;

INSERT INTO PROTHEUS1233_HML.dbo.SBM010 (BM_FILIAL, BM_GRUPO, BM_DESC, BM_PICPAD, BM_PROORI, BM_CODMAR, BM_STATUS, BM_GRUREL, BM_TIPGRU, BM_MARKUP, BM_PRECO, BM_MARGPRE, BM_LENREL, BM_TIPMOV, BM_CODGRT, BM_CLASGRU, BM_FORMUL, BM_TPSEGP, BM_DTUMOV, BM_HRUMOV, BM_CONC, BM_CORP, BM_EVENTO, BM_LAZER, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_, BM_USERLGI, BM_USERLGA) VALUES(N'0101', N'TEST', N'TESTE', N'                              ', N'1', N'   ', N'1', N'                                        ', N'1 ', 0.0, N'   ', 0.0, 0.0, N' ', N'  ', N'1', N'      ', N' ', N'        ', N'     ', N' ', N'F', N'F', N'F', N' ', 278, 0, N' ', N' ');

SELECT BM_GRUPO, BM_DESC FROM PROTHEUS1233_HML.dbo.SBM010 ORDER BY BM_DESC ASC;

SELECT * FROM PROTHEUS12_R27.dbo.SBM010 WHERE BM_GRUPO = '550';

----------------------------------------------------------------------------------------------------------------------------------------

SELECT * FROM PROTHEUS12_R27.dbo.SG1010 WHERE R_E_C_N_O_=353566;

SELECT COUNT(*) FROM PROTHEUS12_R27.dbo.SB1010;

SELECT * FROM PROTHEUS12_R27.dbo.SG1010 WHERE G1_COD = 'M-048-020-284';
SELECT * FROM PROTHEUS1233_HML.dbo.SG1010 WHERE G1_COD = 'M-048-020-284';

SELECT * FROM PROTHEUS12_R27.dbo.SG1010 WHERE G1_COD LIKE 'E9999%';
SELECT * FROM PROTHEUS12_R27.dbo.SG1010 WHERE G1_COMP = 'M-059-026-011' ORDER BY G1_REVFIM ASC;
SELECT TOP 1 R_E_C_N_O_ FROM PROTHEUS12_R27.dbo.SG1010 ORDER BY R_E_C_N_O_ DESC;

-------------------------------------------------------------------------------------------------------------------------------------

INSERT INTO PROTHEUS12_R27.dbo.SG1010 (G1_FILIAL, G1_COD, G1_COMP, G1_TRT, G1_XUM, G1_QUANT, G1_PERDA, G1_INI, G1_FIM, G1_OBSERV, G1_FIXVAR, G1_GROPC, G1_OPC, G1_REVINI, G1_NIV, G1_NIVINV, G1_REVFIM, G1_OK, G1_POTENCI, G1_TIPVEC, G1_VECTOR, G1_VLCOMPE, G1_LOCCONS, G1_USAALT, G1_FANTASM, G1_LISTA, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) VALUES(N'0101', N'E9999-999-999  ', N'E9999-999-998  ', N'   ', N'PC', 16.0, 0.0, N'20230726', N'20491231', N'                                             ', N'V', N'   ', N'    ', N'', N'', N'', N'', N'    ', 0.0, N'      ', N'      ', N'N', N'  ', N'1', N' ', N'          ', N' ', 352930, 0);
INSERT INTO PROTHEUS12_R27.dbo.SG1010 (G1_FILIAL, G1_COD, G1_COMP, G1_TRT, G1_XUM, G1_QUANT, G1_PERDA, G1_INI, G1_FIM, G1_OBSERV, G1_FIXVAR, G1_GROPC, G1_OPC, G1_REVINI, G1_NIV, G1_NIVINV, G1_REVFIM, G1_OK, G1_POTENCI, G1_TIPVEC, G1_VECTOR, G1_VLCOMPE, G1_LOCCONS, G1_USAALT, G1_FANTASM, G1_LISTA, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) VALUES(N'0101', N'E9999-999-999  ', N'E9999-999-997  ', N'   ', N'PC', 16.0, 0.0, N'20230726', N'20491231', N'                                             ', N'V', N'   ', N'    ', N'', N'', N'', N'', N'    ', 0.0, N'      ', N'      ', N'N', N'  ', N'1', N' ', N'          ', N' ', 352931, 0);

INSERT INTO PROTHEUS12_R27.dbo.SG1010 (G1_FILIAL, G1_COD, G1_COMP, G1_TRT, G1_XUM, G1_QUANT, G1_PERDA, G1_INI, G1_FIM, G1_OBSERV, G1_FIXVAR, G1_GROPC, G1_OPC, G1_REVINI, G1_NIV, G1_NIVINV, G1_REVFIM, G1_OK, G1_POTENCI, G1_TIPVEC, G1_VECTOR, G1_VLCOMPE, G1_LOCCONS, G1_USAALT, G1_FANTASM, G1_LISTA, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) VALUES(N'0101', N'E9999-999-998  ', N'E9999-999-996  ', N'   ', N'PC', 16.0, 0.0, N'20230726', N'20491231', N'                                             ', N'V', N'   ', N'    ', N'', N'', N'', N'', N'    ', 0.0, N'      ', N'      ', N'N', N'  ', N'1', N' ', N'          ', N' ', 352854, 0);
INSERT INTO PROTHEUS12_R27.dbo.SG1010 (G1_FILIAL, G1_COD, G1_COMP, G1_TRT, G1_XUM, G1_QUANT, G1_PERDA, G1_INI, G1_FIM, G1_OBSERV, G1_FIXVAR, G1_GROPC, G1_OPC, G1_REVINI, G1_NIV, G1_NIVINV, G1_REVFIM, G1_OK, G1_POTENCI, G1_TIPVEC, G1_VECTOR, G1_VLCOMPE, G1_LOCCONS, G1_USAALT, G1_FANTASM, G1_LISTA, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) VALUES(N'0101', N'E9999-999-997  ', N'E9999-999-995  ', N'   ', N'PC', 16.0, 0.0, N'20230726', N'20491231', N'                                             ', N'V', N'   ', N'    ', N'', N'', N'', N'', N'    ', 0.0, N'      ', N'      ', N'N', N'  ', N'1', N' ', N'          ', N' ', 352855, 0);

INSERT INTO PROTHEUS12_R27.dbo.SG1010 (G1_FILIAL, G1_COD, G1_COMP, G1_TRT, G1_XUM, G1_QUANT, G1_PERDA, G1_INI, G1_FIM, G1_OBSERV, G1_FIXVAR, G1_GROPC, G1_OPC, G1_REVINI, G1_NIV, G1_NIVINV, G1_REVFIM, G1_OK, G1_POTENCI, G1_TIPVEC, G1_VECTOR, G1_VLCOMPE, G1_LOCCONS, G1_USAALT, G1_FANTASM, G1_LISTA, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) VALUES(N'0101', N'E9999-999-999  ', N'E9999-999-996  ', N'   ', N'PC', 1.0, 0.0, N'20231109', N'20491231', N'                                             ', N'V', N'   ', N'    ', N'001', N'01', N'99', N'001', N'    ', 0.0, N'      ', N'      ', N'N', N'  ', N'1', N' ', N'          ', N' ', 352932, 0);

-------------------------------------------------------------------------------------------------------------------------------------

SELECT * FROM PROTHEUS12_R27.dbo.SB1010 ORDER BY R_E_C_N_O_ DESC;

SELECT TOP 1 * FROM PROTHEUS12_R27.dbo.SB1010 ORDER BY R_E_C_N_O_ DESC;

SELECT B1_COD, B1_DESC, B1_XDESC2, B1_MSBLQL FROM PROTHEUS12_R27.dbo.SB1010 s WHERE B1_DESC LIKE 'CEMEN%';

SELECT * FROM PROTHEUS12_R27.dbo.SB1010 WHERE D_E_L_E_T_ = '*';

SELECT B1_COD, B1_DESC, B1_MSBLQL FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_COD LIKE 'E7047-001-1%' ORDER BY B1_COD ASC;

SELECT B1_COD, B1_DESC FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_COD LIKE 'C-007-101%' ORDER BY 1 DESC;

SELECT * FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_COD LIKE 'E9999%' ORDER BY B1_COD DESC;
SELECT B1_COD, B1_DESC, B1_REVATU FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_COD LIKE 'E9999%' ORDER BY B1_COD DESC;

SELECT * FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_DESC LIKE 'MOTOR%' AND B1_DESC LIKE '%CA%';

SELECT * FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_COD LIKE 'M-150-007%' AND B1_DESC LIKE '%500%' AND B1_DESC LIKE '%400%' AND B1_MSBLQL = '2';

SELECT * FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_COD LIKE 'M-150-010-97%';

SELECT B1_REVATU FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_COD = 'E9999-999-010';

SELECT * FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_COD = 'M-048-020-452';

SELECT * FROM PROTHEUS1233_HML.dbo.SB1010 WHERE B1_COD = 'M-048-020-452';

-------------------------------------------------------------------------------------------------------------------------------------

INSERT INTO PROTHEUS12_R27.dbo.SB1010 (B1_AFAMAD, B1_FILIAL, B1_COD, B1_DESC, B1_XDESC2, B1_CODITE, B1_TIPO, B1_UM, B1_LOCPAD, B1_GRUPO, B1_ZZLOCAL, B1_POSIPI, B1_ESPECIE, B1_EX_NCM, B1_EX_NBM, B1_PICM, B1_IPI, B1_ALIQISS, B1_CODISS, B1_TE, B1_TS, B1_PICMRET, B1_BITMAP, B1_SEGUM, B1_PICMENT, B1_IMPZFRC, B1_CONV, B1_TIPCONV, B1_ALTER, B1_QE, B1_PRV1, B1_EMIN, B1_CUSTD, B1_UCALSTD, B1_UCOM, B1_UPRC, B1_MCUSTD, B1_ESTFOR, B1_PESO, B1_ESTSEG, B1_FORPRZ, B1_PE, B1_TIPE, B1_LE, B1_LM, B1_CONTA, B1_TOLER, B1_CC, B1_ITEMCC, B1_PROC, B1_LOJPROC, B1_FAMILIA, B1_QB, B1_APROPRI, B1_TIPODEC, B1_ORIGEM, B1_CLASFIS, B1_UREV, B1_DATREF, B1_FANTASM, B1_RASTRO, B1_FORAEST, B1_COMIS, B1_DTREFP1, B1_MONO, B1_PERINV, B1_GRTRIB, B1_MRP, B1_NOTAMIN, B1_CONINI, B1_CONTSOC, B1_PRVALID, B1_CODBAR, B1_GRADE, B1_NUMCOP, B1_FORMLOT, B1_IRRF, B1_FPCOD, B1_CODGTIN, B1_DESC_P, B1_CONTRAT, B1_DESC_GI, B1_DESC_I, B1_LOCALIZ, B1_OPERPAD, B1_ANUENTE, B1_OPC, B1_CODOBS, B1_VLREFUS, B1_IMPORT, B1_FABRIC, B1_SITPROD, B1_MODELO, B1_SETOR, B1_PRODPAI, B1_BALANCA, B1_TECLA, B1_DESPIMP, B1_TIPOCQ, B1_SOLICIT, B1_GRUPCOM, B1_QUADPRO, B1_BASE3, B1_DESBSE3, B1_AGREGCU, B1_NUMCQPR, B1_CONTCQP, B1_REVATU, B1_CODEMB, B1_INSS, B1_ESPECIF, B1_NALNCCA, B1_MAT_PRI, B1_NALSH, B1_REDINSS, B1_REDIRRF, B1_ALADI, B1_TAB_IPI, B1_GRUDES, B1_DATASUB, B1_REDPIS, B1_REDCOF, B1_PCSLL, B1_PCOFINS, B1_PPIS, B1_MTBF, B1_MTTR, B1_FLAGSUG, B1_CLASSVE, B1_MIDIA, B1_QTMIDIA, B1_QTDSER, B1_VLR_IPI, B1_ENVOBR, B1_SERIE, B1_FAIXAS, B1_NROPAG, B1_ISBN, B1_TITORIG, B1_LINGUA, B1_EDICAO, B1_OBSISBN, B1_CLVL, B1_ATIVO, B1_EMAX, B1_PESBRU, B1_TIPCAR, B1_FRACPER, B1_VLR_ICM, B1_INT_ICM, B1_CORPRI, B1_CORSEC, B1_NICONE, B1_ATRIB1, B1_ATRIB2, B1_ATRIB3, B1_REGSEQ, B1_VLRSELO, B1_CODNOR, B1_CPOTENC, B1_POTENCI, B1_REQUIS, B1_SELO, B1_LOTVEN, B1_OK, B1_USAFEFO, B1_QTDACUM, B1_QTDINIC, B1_CNATREC, B1_TNATREC, B1_AFASEMT, B1_AIMAMT, B1_TERUM, B1_AFUNDES, B1_CEST, B1_GRPCST, B1_IAT, B1_IPPT, B1_GRPNATR, B1_DTFIMNT, B1_DTCORTE, B1_FECP, B1_MARKUP, B1_CODPROC, B1_LOTESBP, B1_QBP, B1_VALEPRE, B1_CODQAD, B1_AFABOV, B1_VIGENC, B1_VEREAN, B1_DIFCNAE, B1_ESCRIPI, B1_PMACNUT, B1_PMICNUT, B1_INTEG, B1_HREXPO, B1_CRICMS, B1_REFBAS, B1_MOPC, B1_USERLGI, B1_USERLGA, B1_UMOEC, B1_UVLRC, B1_PIS, B1_GCCUSTO, B1_CCCUSTO, B1_TALLA, B1_PARCEI, B1_GDODIF, B1_VLR_PIS, B1_TIPOBN, B1_TPREG, B1_MSBLQL, B1_VLCIF, B1_DCRE, B1_DCR, B1_DCRII, B1_TPPROD, B1_DCI, B1_COEFDCR, B1_CHASSI, B1_CLASSE, B1_FUSTF, B1_GRPTI, B1_PRDORI, B1_APOPRO, B1_PRODREC, B1_ALFECOP, B1_ALFECST, B1_CFEMA, B1_FECPBA, B1_MSEXP, B1_PAFMD5, B1_PRODSBP, B1_CODANT, B1_IDHIST, B1_CRDEST, B1_REGRISS, B1_FETHAB, B1_ESTRORI, B1_CALCFET, B1_PAUTFET, B1_CARGAE, B1_PRN944I, B1_ALFUMAC, B1_PRINCMG, B1_PR43080, B1_RICM65, B1_SELOEN, B1_TRIBMUN, B1_RPRODEP, B1_FRETISS, B1_AFETHAB, B1_DESBSE2, B1_BASE2, B1_VLR_COF, B1_PRFDSUL, B1_TIPVEC, B1_COLOR, B1_RETOPER, B1_COFINS, B1_CSLL, B1_CNAE, B1_ADMIN, B1_AFACS, B1_AJUDIF, B1_ALFECRN, B1_CFEM, B1_CFEMS, B1_MEPLES, B1_REGESIM, B1_RSATIVO, B1_TFETHAB, B1_TPDP, B1_CRDPRES, B1_CRICMST, B1_FECOP, B1_CODLAN, B1_GARANT, B1_PERGART, B1_SITTRIB, B1_PORCPRL, B1_IMPNCM, B1_IVAAJU, B1_BASE, B1_ZZCODAN, B1_ZZNOGRP, B1_ZZOBS1, B1_XFORDEN, B1_ZZMEN1, B1_ZZLEGIS, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_) VALUES(0.0, N'    ', N'M-090-002-001  ', N'REVESTIMENTO P/ CONJ. ASPA REVESTIDA L=342 - CONF. DESENHO M-039-033-585                            ', N'                                                            ', N'                           ', N'MP', N'PC', N'97', N'116 ', N'      ', N'          ', N'  ', N'   ', N'   ', 0.0, 0.0, 0.0, N'         ', N'   ', N'   ', 0.0, N'                    ', N'  ', 0.0, N' ', 0.0, N'M', N'               ', 0.0, 0.0, 0.0, 0.0, N'20230220', N'20230220', 215.0, N'1', N'   ', 0.0, 0.0, N'   ', 0.0, N' ', 0.0, 0.0, N'11301020001         ', 0.0, N'5.4.1.01 ', N'         ', N'      ', N'  ', N' ', 1.0, N' ', N'N', N' ', N'  ', N'20230220', N'20230214', N' ', N'N', N' ', 0.0, N'        ', N' ', 0.0, N'      ', N'S', 0.0, N'20230615', N' ', 0.0, N'               ', N' ', 0.0, N'   ', N' ', N'          ', N'               ', N'      ', N'N', N'      ', N'      ', N'N', N'  ', N'2', N'                                                                                ', N'      ', 0.0, N'N', N'                    ', N'  ', N'               ', N'  ', N'               ', N' ', N'   ', N'N', N'M', N'N', N'      ', N' ', N'              ', N'                                                            ', N'2', 0.0, 0.0, N'   ', N'                              ', N'N', N'                                                                                ', N'       ', N'                    ', N'        ', 0.0, 0.0, N'   ', N'  ', N'   ', N'        ', 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, N'1', N'1', N'2', 0.0, N'1', 0.0, N'0', N'                    ', 0.0, 0.0, N'          ', N'                                                  ', N'                    ', N'   ', N'                                        ', N'         ', N'S', 0.0, 0.0, N'      ', 0.0, 0.0, 0.0, N'      ', N'      ', N'               ', N'      ', N'      ', N'      ', N'      ', 0.0, N'   ', N'2', 0.0, N' ', N' ', 0.0, N'    ', N'1', 0.0, 0.0, N'   ', N'    ', 0.0, 0.0, N'  ', 0.0, N'         ', N'   ', N' ', N' ', N'  ', N'        ', N'        ', 0.0, 0.0, N'      ', 0.0, 0.0, N' ', N'                      ', 0.0, N'        ', N'  ', N'           ', N'3', 0.0, 0.0, N' ', N'        ', N'0', N' ', NULL, N' 0#  0@  10• 908 ', N' 0#  0@  502 202 ', 0.0, 0.0, N'2', N'        ', N'         ', N'      ', N'      ', N' ', 0.0, N'  ', N' ', N'2', 0.0, N'          ', N'         ', 0.0, N'  ', N' ', 0.0, N'                         ', N'      ', N' ', N'    ', N'               ', N' ', N' ', 0.0, 0.0, 0.0, 0.0, N'        ', N'                                ', N'C', N'               ', N'                    ', 0.0, N'  ', N'N', N'               ', N' ', 0.0, N' ', N'S', 0.0, 0.0, 0.0, N'2', N'      ', N'                    ', N' ', N' ', 0.0, N'                                                            ', N'              ', 0.0, 0.0, N'      ', N'          ', N'2', N'2', N'2', N'         ', N'          ', 0.0, N' ', 0.0, N' ', N' ', N' ', N' ', N' ', N' ', N' ', 0.0, N' ', N' ', N'      ', N'2', 0.0, N' ', N'  ', 0.0, N' ', N'              ', N'               ', N'TRATAMENTO SUPERFICIAL        ', NULL, N' ', N'   ', N'                                                                                                                                                                                                                                                          ', N' ', 59479, 0);
INSERT INTO PROTHEUS12_R27.dbo.SB1010
(B1_COD, B1_DESC, B1_XDESC2, B1_TIPO, B1_UM, B1_LOCPAD, B1_GRUPO, D_E_L_E_T_, R_E_C_N_O_, R_E_C_D_E_L_)
VALUES
(N'E9999-998-800', N'TESTE CADASTRO ITEM 20-10-2023', N'', N'MP', N'PC', N'03', N'202 ', N' ', 63597, 0);

-------------------------------------------------------------------------------------------------------------------------------------

UPDATE PROTHEUS12_R27.dbo.SB1010 SET B1_DESC = 'TESTE DESC', B1_XDESC2 = 'TESTE DESC2', B1_TIPO = 'MP', B1_UM = 'PC', B1_LOCPAD = '03', B1_GRUPO = '202', B1_ZZNOGRP = '', B1_CC = '5.3.1.01' WHERE B1_COD = 'E9999-999-005'; -- 1 DESBLOQUEADO 2 BLOQUEADO  

UPDATE PROTHEUS12_R27.dbo.SB1010 SET B1_MSBLQL = '2' WHERE B1_COD = 'M-090-002-010';

-------------------------------------------------------------------------------------------------------------------------------------

DELETE FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_COD LIKE 'E9999%';

DELETE FROM PROTHEUS12_R27.dbo.SG1010 WHERE G1_COD LIKE 'E9999%';

DELETE FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_COD = 'E7047-001-939';

DELETE FROM PROTHEUS12_R27.dbo.SB1010 WHERE B1_COD = 'C-007-101-328';

DELETE FROM PROTHEUS12_R27.dbo.SB1010 WHERE R_E_C_N_O_=63768;

DELETE FROM PROTHEUS12_R27.dbo.SG1010 WHERE R_E_C_N_O_=353566;

DELETE FROM PROTHEUS12_R27.dbo.SG1010 WHERE R_E_C_N_O_ IN (352852, 352853);
