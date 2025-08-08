
Public Sub Main()

    ' Declaração das variáveis
    
    Dim qry_email As Object
    Dim Email As Mail
    Dim Corpo_Email As String
    Dim Registros_Processados As String
    Dim qrySQL As Object
    Dim qryUPDT_2 As Object
    Dim qryUPDT_3 As Object
    Dim codigoGDS As String
    Dim empresa As String

	' Pesquisa as vendas sem fornecedor

    Set qrySQL_2 = NewQuery
    qrySQL_2.Active = False
    qrySQL_2.Clear

    qrySQL_2.Add("Select a.datainclusao, a.localizadora, a.handle as handlepnr, b.handle as handleacc, " & _
				"Case " & _
        		"when charindex('<VendorCode>', convert(varchar(max), log.xmlreserva)) > 0 then " & _
            	"substring( " & _
                "convert(varchar(max), Log.xmlreserva), " & _
                "charindex('<VendorCode>', convert(varchar(max), log.xmlreserva), charindex('<Vehicle', convert(varchar(max), log.xmlreserva))) + len('<vendorcode>'), " & _
				"charindex('</VendorCode>', convert(varchar(max), log.xmlreserva), charindex('<Vehicle', convert(varchar(max), log.xmlreserva))) " & _
				"- charindex('<VendorCode>', convert(varchar(max), log.xmlreserva), charindex('<Vehicle', convert(varchar(max), log.xmlreserva))) " & _
				"- Len('<vendorcode>') " & _
            	") " & _
        		"Else '' end as vendorcode, " & _
    			"Case " & _
        		"when charindex('<locationCode>', convert(varchar(max), log.xmlreserva)) > 0 then " & _
            	"substring( " & _
                "convert(varchar(max), Log.xmlreserva), " & _
                "charindex('<locationcode>', convert(varchar(max), log.xmlreserva)) + len('<locationcode>'), " & _
                "charindex('</locationcode>', convert(varchar(max), log.xmlreserva)) - " & _
                "(charindex('<locationcode>', convert(varchar(max), log.xmlreserva)) + len('<locationcode>')) " & _
            	") " & _
    			"Else '' end as locationcode, " & _
    			"Case " & _
        		"when charindex('<locationname>', convert(varchar(max), log.xmlreserva)) > 0 then " & _
            	"substring(" & _
                "convert(varchar(max), Log.xmlreserva), " & _
                "charindex('<locationname>', convert(varchar(max), log.xmlreserva)) + len('<locationname>'), " & _
                "charindex('</locationname>', convert(varchar(max), log.xmlreserva)) - " & _
                "(charindex('<locationname>', convert(varchar(max), log.xmlreserva)) + len('<locationname>')) " & _
            	") " & _
        		"Else '' end as locationname " & _
				"from vm_pnrs a " & _
				"Left Join vm_pnraccountings b On a.Handle = b.pnr " & _
				"Left Join bb_logintegracoes Log On a.logintegracao = Log.Handle " & _
				"where a.tiporeserva In (1) " & _
				"And b.tipoacc = 2 " & _
				"And a.situacao In (3) And b.fornecedor In (0, Null)" & _
				"")

    qrySQL_2.Active = True


    ' UPDATE: procura fornecedores com os códigos gds correspondetes para inserir nas vendas

    Do While Not qrySQL_2.EOF

		 	' Concatena os campos vendorcode + locationcode ou locationame (caso exista)

	    	If qrySQL_2.FieldByName("locationname").AsString <> "" Then
				codigoGDS = qrySQL_2.FieldByName("vendorcode").AsString + qrySQL_2.FieldByName("locationname").AsString
			Else
				codigoGDS = qrySQL_2.FieldByName("vendorcode").AsString + qrySQL_2.FieldByName("locationcode").AsString
			End If

        Set qrySQL_3 = NewQuery
        qrySQL_3.Active = False
        qrySQL_3.Clear
        qrySQL_3.Add("SELECT CODIGOGDS, CONTRATO "+ _
                        "FROM BB_FORNECEDORCONTRATOCODIGOS  "+ _
                        "WHERE TIPORESERVA IN (1) AND SISTEMARESERVA IS NULL AND CODIGOGDS IN ('" & codigoGDS & "')")

        qrySQL_3.Active = True

        Set qryUPDT_2 = NewQuery
        qryUPDT_2.Active = False
        qryUPDT_2.Clear
        qryUPDT_2.Add("UPDATE VM_PNRACCOUNTINGS SET FORNECEDOR = '" + qrySQL_3.FieldByName("CONTRATO").AsString + "' WHERE HANDLE = " + qrySQL_2.FieldByName("HANDLEACC").AsString)

        qryUPDT_2.ExecSQL

        Set qryUPDT_3 = NewQuery
        qryUPDT_3.Active = False
        qryUPDT_3.Clear
        qryUPDT_3.Add("UPDATE A SET SITUACAO=1, CONCLUIDO='S', EXPORTADO='N', AGUARDANDOEMISSAO='N' FROM VM_PNRS A LEFT JOIN VM_PNRACCOUNTINGS B ON A.HANDLE = B.PNR WHERE B.FORNECEDOR NOT IN (0,NULL) AND A.HANDLE = " + qrySQL_2.FieldByName("handlepnr").AsString)

        qryUPDT_3.ExecSQL

        qrySQL_2.Next

    Loop

    ' Inicializa a consulta para buscar vendas sem fornecedor

    Set qry_email = NewQuery
    qry_email.Active = False
    qry_email.Clear

    qry_email.Add("Select a.datainclusao, a.localizadora, b.requisicao, a.handle as handlepnr, b.handle as handleacc, " & _
				"Case " & _
        		"when charindex('<VendorCode>', convert(varchar(max), log.xmlreserva)) > 0 then " & _
            	"substring( " & _
                "convert(varchar(max), Log.xmlreserva), " & _
                "charindex('<VendorCode>', convert(varchar(max), log.xmlreserva), charindex('<Vehicle', convert(varchar(max), log.xmlreserva))) + len('<vendorcode>'), " & _
				"charindex('</VendorCode>', convert(varchar(max), log.xmlreserva), charindex('<Vehicle', convert(varchar(max), log.xmlreserva))) " & _
				"- charindex('<VendorCode>', convert(varchar(max), log.xmlreserva), charindex('<Vehicle', convert(varchar(max), log.xmlreserva))) " & _
				"- Len('<vendorcode>') " & _
            	") " & _
        		"Else '' end as vendorcode, " & _
    			"Case " & _
        		"when charindex('<locationCode>', convert(varchar(max), log.xmlreserva)) > 0 then " & _
            	"substring( " & _
                "convert(varchar(max), Log.xmlreserva), " & _
                "charindex('<locationcode>', convert(varchar(max), log.xmlreserva)) + len('<locationcode>'), " & _
                "charindex('</locationcode>', convert(varchar(max), log.xmlreserva)) - " & _
                "(charindex('<locationcode>', convert(varchar(max), log.xmlreserva)) + len('<locationcode>')) " & _
            	") " & _
    			"Else '' end as locationcode, " & _
    			"Case " & _
        		"when charindex('<locationname>', convert(varchar(max), log.xmlreserva)) > 0 then " & _
            	"substring(" & _
                "convert(varchar(max), Log.xmlreserva), " & _
                "charindex('<locationname>', convert(varchar(max), log.xmlreserva)) + len('<locationname>'), " & _
                "charindex('</locationname>', convert(varchar(max), log.xmlreserva)) - " & _
                "(charindex('<locationname>', convert(varchar(max), log.xmlreserva)) + len('<locationname>')) " & _
            	") " & _
        		"Else '' end as locationname " & _
				"from vm_pnrs a " & _
				"Left Join vm_pnraccountings b On a.Handle = b.pnr " & _
				"Left Join bb_logintegracoes Log On a.logintegracao = Log.Handle " & _
				"where a.tiporeserva In (1) " & _
				"And b.tipoacc = 2 " & _
				"And a.situacao In (3) And b.fornecedor In (0, Null)" & _
				"")

    qry_email.Active = True

    ' Configura o e-mail

    Set Email = NewMail
    Email.SendTo = "diegomoreira@kontik.com.br"
    Email.Subject = "Portal Benner - Processado Erro - CARRO sem Fornecedor - SABRE - " & Format(Now, "DD/MM/YYYY")

    ' Inicializa o corpo do e-mail

    Corpo_Email = "Portal Benner - Processado Erro - CARRO sem Fornecedor - SABRE - " & Format(Now, "DD/MM/YYYY") & vbNewLine & vbNewLine
    Corpo_Email = Corpo_Email & "CÓDIGO DAS LOCADORAS PARA CADASTRAR:" & vbNewLine & vbNewLine

    ' Processa todos os registros encontrados

    Do While Not qry_email.EOF

		' concatena o campos vendorcode + locationcode ou locationame (caso exista)

    	If qry_email.FieldByName("locationname").AsString <> "" Then
			codigoGDS = qry_email.FieldByName("vendorcode").AsString + qry_email.FieldByName("locationname").AsString
		Else
			codigoGDS = qry_email.FieldByName("vendorcode").AsString + qry_email.FieldByName("locationcode").AsString
		End If

		' identifica a empresa pelo vendorCode

		If qry_email.FieldByName("vendorcode").AsString = "EP" Then
		    empresa = "EUROPCAR"
		ElseIf qry_email.FieldByName("vendorcode").AsString = "ET" Then
		    empresa = "ENTERPRISE"
		ElseIf qry_email.FieldByName("vendorcode").AsString = "LL" Then
		    empresa = "LOCALIZA"
		ElseIf qry_email.FieldByName("vendorcode").AsString = "MO" Then
		    empresa = "MOVIDA"
		ElseIf qry_email.FieldByName("vendorcode").AsString = "MV" Then
		    empresa = "MOVIDA"
		ElseIf qry_email.FieldByName("vendorcode").AsString = "ZD" Then
		    empresa = "BUDGET"
		ElseIf qry_email.FieldByName("vendorcode").AsString = "ZE" Then
		    empresa = "HERTZ"
		ElseIf qry_email.FieldByName("vendorcode").AsString = "ZI" Then
		    empresa = "AVIS"
		ElseIf qry_email.FieldByName("vendorcode").AsString = "ZL" Then
		    empresa = "NATIONAL"
		Else
			empresa = "INDEFINIDO"
		End If

        ' Verifica se o campo vendorCode não é nulo

        If Not IsNull(qry_email.FieldByName("vendorcode").AsString) Then

            ' Acumula as informações de cada registro

            Registros_Processados =	Format(qry_email.FieldByName("datainclusao").AsDateTime, "DD/MM/YYYY") & _
            						" - " & qry_email.FieldByName("localizadora").AsString & " - " & _
            						empresa & " - " & _
            						qry_email.FieldByName("locationcode").AsString & _
                                    " - SABRE - " & codigoGDS & vbNewLine

            Corpo_Email = Corpo_Email & Registros_Processados
        End If

        ' Avança para o próximo registro

        qry_email.Next

    Loop

    ' Finaliza o corpo do e-mail

    Corpo_Email = Corpo_Email & vbNewLine & "Cadastrar os códigos das locadoras acima no Benner, da seguinte forma:" & vbNewLine
    Corpo_Email = Corpo_Email & "SABRE - CÓDIGO GDS - NULO" & vbNewLine & vbNewLine
    Corpo_Email = Corpo_Email & "Após o cadastro as vendas no processado erro serão regularizadas." & vbNewLine
    Corpo_Email = Corpo_Email & "---------------x---------------x---------------x---------------x---------------" & vbNewLine
    Corpo_Email = Corpo_Email & vbNewLine & "Kontik Business Travel" & vbNewLine & "Equipe TI"

    ' Envia o e-mail

    Email.Text.Add Corpo_Email
    Email.Send

    ' Limpeza de objetos

    Set Email = Nothing
    Set qry_email = Nothing

    Set qrySQL = Nothing
    Set qryUPDT_2 = Nothing
    Set qryUPDT_3 = Nothing

End Sub
