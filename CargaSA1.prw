#Include 'protheus.ch'
#Include 'tbiconn.ch'
#Include 'fileio.ch'
/*/{Protheus.doc} User Function fAxCadas
    Funcao utilizada para importar arquivo CSV com cadastro de Clientes e gravar na Base de Dados atraves do ExecAuto
    @type  Function
    @author Scheron Martins
    @since 02/10/2023
    @version 1.0
    @param Nenhum
    @return Vazio (nil)
        @see : ExecAuto          https://tdn.totvs.com/pages/releaseview.action?pageId=285654185
		Função FCREATE           https://tdn.totvs.com/display/public/framework/FCreate
		Função CGETFILE          https://tdn.totvs.com/display/tec/cGetFile
		Função FTFUSE            https://tdn.totvs.com/display/tec/FT_FUse
		Função MSGSTOP           https://tdn.totvs.com/pages/releaseview.action?pageId=24346998
/*/

User Function CargaSA1()
    Local nCliente
	Local i

	Private aDados	:= {}
    Private aCabec  := {}
	Private ENTER 	 := CHR(13)+CHR(10)
	Private nHandle1 := 0
	Private nHandle2 := 0
	Private nOk		 := 0
	Private nTotal	 := 0 	
	Private nNok	 := 0 
	
    //Cria um Arquivo de LOG
    cDirD	 := "C:\temp\log\"
	cFileLOG := cDirD + "CargaSA1.log"
    nHandle1 := FCREATE(cFileLOG)
	if nHandle1 = -1
		Alert("Erro ao criar arquivo - ferror " + Str(Ferror()))
		Return
	Endif
    IniLOG()                                            // cabeçalho do arquivo LOG
	GravaLOG("Preparar para carregar Cabecalho e Dados")//Grava a data e horario no arquivo de log

    //Abre uma Janela Para o Usuario escolher o arquivo CSV
    cMascara    := "*.csv|*.csv"
    cTitulo     := "Selecione o Arquivo CSV com os cadastros de Clientes que deseja importar"
    nMascpadra  := 0
	cDir        := "C:\temp\"
    lAbrir      := .T.
    lArvore     := .T.
    cArquivo    := cGetFile(cMascara, cTitulo, nMascpadra, cDir, lAbrir,/*nOpcoes*/, lArvore)  //Abre janela para escolha de diretorio
    

    //Abre o arquivo selecionado anteriormente pelo usuario.
	nHandle2    := FT_FUSE(cArquivo)
    If nHandle2 == -1	// Se houver erro de abertura, mostra mensagem e grava texto no arquivo de Log
		cTexto := "Erro de abertura : FERROR " + str(ferror() )
		MsgStop( cTexto , 4)
		GravaLOG( cTexto )// mostra mensagem no arquivo de log
		FClose(nHandle2)
	EndIf


    // Carrega os dados do arquivo csv para um array de dados
	lCabec	:= .T.
	aCabec := {}
	aDados := {}
	nLinha	:= 1
	//nTotal	:= ft_flastrec() -1 // Total de Linhas com conteudo
	//nOk		:= 0 				// Linhas processadas
	//nNok	:= 0				// Linhas nao processadas 
	While !FT_FEOF() //Lê todo o arquivo enquanto não for o final dele		
		GravaLOG( "lendo linha no " + Str( nLinha++ )) //Grava a leitura linha a linha do arquivo CSV
		cLine := FT_FReadLn() //retorna a linha com o texto
		If lCabec == .T.
			lCabec:= .F.	// Le a primeira linha e desabilita cabelalho
			aAdd(aCabec, StrTokArr(cLine, ";")) //StrTokArr()Retorna um array, de acordo com os dados passados como parâmetro à função. 
		Else
			aAdd(aDados, StrTokArr(cLine, ";"))
		Endif
		FT_FSKIP()
	End
	GravaLOG("Conseguiu carregar Cabecalho e Dados") // Grava que conseguiu passar os dados para o ARRAy

	//Tranforma o array multidimencional em 1 vetor
	aHeader  := Array(  Len( aCabec[1] ) )
	For i := 1 to Len( aCabec [1])
		aHeader [i] := aCabec[1][i]
	Next
	
	//Tranforma o array multidimencional em 1 vetor
	For nCliente := 1 to Len( aDados)
		GravaLOG( "Gravando linha no " + Str( nCliente ) )
		aCols	:= aDados[nCliente]
	//Chama a função para executar o MSExecAuto
		xMATA030(aHeader, aCols)	
	Next

	If nOk	== 0 	// Não pocessou nada
			GravaLOG("Não processou nada")
	Else
		GravaLOG("Processado com sucesso!")
	Endif
	FimLOG()

Return

//Cabeçaho do arquivo de TXT log
Static Function IniLog()
	FWrite(nHandle1, REPLICATE("=",120) + ENTER,99)
	FWrite(nHandle1, "INICIO DO LOG"	+ ENTER,99)
	FWrite(nHandle1, REPLICATE("=",120) + ENTER,99)
Return

//Grava no arquivo log o horario e data e um texto passado como parametro
Static Function GravaLog( _cTEXTO )
	Local cTexto := ""
	cTexto += DtoC( Date() ) + " - "
	cTexto += Time() + " - "
	cTexto += _cTEXTO
	cTexto += ENTER

	FWrite(nHandle1, cTEXTO ,99)
Return( .T. )

Static Function FimLog()
	FWrite(nHandle1, REPLICATE("=",130) + ENTER,99)
	FWrite(nHandle1, "FIM DO LOG"		+ ENTER,99)
	FWrite(nHandle1, "Total de Linhas da planilha"  +str(nTotal)+ ENTER,99)
	FWrite(nHandle1, "Total linhas processadas" 	+str(nOk)	+ ENTER,99)
	FWrite(nHandle1, "Total linhas nao processadas" +str(nNok)	+ ENTER,99)
	FWrite(nHandle1, REPLICATE("=",130) + ENTER,99)
	
	FClose(nHandle1)

Return

//Função com MSExecAuto
Static Function xMATA030(aCabec, aDados)

	Local aSA1Auto := {}	// cabeçalho
	Local aAI0Auto := {}	// itens
	Local nOpcAuto := 3 	// 5.excluir - 3.inserir
	Local lRet := .T.

	Private lMsErroAuto := .F.

	//lRet := RpcSetEnv("T1","D MG 01","Admin")    //abertura de ambiente para rotinas automáticas, permitindo definir empresa e filial

	If lRet

		//----------------------------------
		// Dados do Cliente                 |
		//----------------------------------

		aAdd(aSA1Auto,{"A1_COD"    ,aCols[AScan(aHeader,{|x| Upper(AllTrim(x)) == "A1_COD"		})] ,Nil})
		aAdd(aSA1Auto,{"A1_LOJA"   ,aCols[AScan(aHeader,{|x| Upper(AllTrim(x)) == "A1_LOJA"		})] ,Nil})
		aAdd(aSA1Auto,{"A1_NOME"   ,aCols[AScan(aHeader,{|x| Upper(AllTrim(x)) == "A1_NOME"		})] ,Nil})
		aAdd(aSA1Auto,{"A1_NREDUZ" ,aCols[AScan(aHeader,{|x| Upper(AllTrim(x)) == "A1_NREDUZ"	})] ,Nil}) 
		aAdd(aSA1Auto,{"A1_TIPO"   ,aCols[AScan(aHeader,{|x| Upper(AllTrim(x)) == "A1_TIPO"		})] ,Nil})
		aAdd(aSA1Auto,{"A1_END"    ,aCols[AScan(aHeader,{|x| Upper(AllTrim(x)) == "A1_END"		})] ,Nil}) 
		aAdd(aSA1Auto,{"A1_BAIRRO" ,aCols[AScan(aHeader,{|x| Upper(AllTrim(x)) == "A1_BAIRRO"	})] ,Nil}) 
		aAdd(aSA1Auto,{"A1_EST"    ,aCols[AScan(aHeader,{|x| Upper(AllTrim(x)) == "A1_EST"		})] ,Nil})
		aAdd(aSA1Auto,{"A1_MUN"    ,aCols[AScan(aHeader,{|x| Upper(AllTrim(x)) == "A1_MUN"		})] ,Nil})

		//---------------------------------------------------------
		// Dados do Complemento do Cliente                         |
		//---------------------------------------------------------
		aAdd(aAI0Auto,{"AI0_SALDO" ,30 ,Nil})

		//------------------------------------
		// Chamada para cadastrar o cliente.  |
		//------------------------------------
		MSExecAuto({|a,b,c| MATA030(a,b,c)}, aSA1Auto, nOpcAuto, aAI0Auto)

		If lMsErroAuto 
			lRet := lMsErroAuto
			nNok++
			MostraErro()// não usar via JOB
		Else
			nOk++
			Conout("Cliente incluído com sucesso!") //opção 3
			//	Conout("Cliente excluido com sucesso!") //opção 5
		EndIf

	EndIf

Return (.T.)
