Public Sub Main
	Dim qrySQL As Object
	Dim qryFOP As Object
	Dim updACC As Object
	Dim updPNR As Object

	Set qrySQL = NewQuery
	Set qryFOP = NewQuery

	qrySQL.Clear
	qrySQL.Active = False

	qryFOP.Clear
	qryFOP.Active = False

	' 1 - Buscar PNRs com erro de forma de pagamento
	qrySQL.Add("SELECT a.handle pnr, c.handle acc, a.datainclusao, a.localizadora, a.situacao, " & _
	           "b.nome cliente, a.cliente handle_cliente, c.passageironaocad pax " & _
	           "FROM vm_pnrs a " & _
	           "LEFT JOIN gn_pessoas b ON a.cliente = b.Handle " & _
	           "LEFT JOIN vm_pnraccountings c ON a.Handle = c.pnr " & _
	           "WHERE a.situacao IN (1,3) " & _
	           "AND c.tipoacc = 3 " & _
	           "AND a.cliente IN (35787,35115) " & _
	           "AND a.mensagem LIKE '%Este cliente não possui permissão para usar este tipo de pagamento e recebimento para este produto%' " & _
	           "AND a.datainclusao >= '01/01/2025' " & _
	           "AND a.tiporeserva IN (27) ")

	qrySQL.Active = True

	Do While Not qrySQL.EOF()

		' 2 - Buscar FOP recente do passageiro
		qryFOP.Clear
		qryFOP.Active = False

		qryFOP.Add("SELECT TOP 1 tipocartao, numerocartao, FORMAT(venccartao, 'MM/dd/yyyy') AS venc_cartao, titularcartaopg " & _
		           "FROM bb_pnraccountings " & _
		           "WHERE YEAR(datainclusao) = YEAR(GETDATE()) " & _
		           "AND formadepagamento = 1 " & _
		           "AND formarecebimento = 7 " & _
		           "AND passageironaocad = '" & Replace(qrySQL.FieldByName("pax").AsString, "'", "''") & "' " & _
		           "AND bb_cliente = " & qrySQL.FieldByName("handle_cliente").AsInteger & " " & _
		           "ORDER BY datainclusao DESC")

		qryFOP.Active = True

		' Se encontrou FOP válida
		If Not qryFOP.EOF Then

			' 3 - Atualizar ACC
			Set updACC = NewQuery

			On Error GoTo ERRO_ACC
			If Not InTransaction Then StartTransaction

			' 3.1 - Gera uma autorização aleatório para a forma de recebimento
			Dim autorizacao As String

			autorizacao = Right("000000" & CStr(Int(Rnd() * 1000000)), 6)

			updACC.Clear
			updACC.Add("UPDATE vm_pnraccountings SET " & _
			           "formapagamento = 2, " & _
			           "formarecebimento = 3, " & _
			           "tipoccpg = '" & Replace(qryFOP.FieldByName("tipocartao").AsString, "'", "''") & "', " & _
			           "numeroccpg = '" & Replace(qryFOP.FieldByName("numerocartao").AsString, "'", "''") & "', " & _
			           "vencccpg = '" & Replace(qryFOP.FieldByName("venc_cartao").AsString, "'", "''") & "', " & _
			           "titularccpg = '" & Replace(qryFOP.FieldByName("titularcartaopg").AsString, "'", "''") & "', " & _
			           "tipoccrc = '" & Replace(qryFOP.FieldByName("tipocartao").AsString, "'", "''") & "', " & _
			           "numeroccrc = '" & Replace(qryFOP.FieldByName("numerocartao").AsString, "'", "''") & "', " & _
			           "vencccrc = '" & Replace(qryFOP.FieldByName("venc_cartao").AsString, "'", "''") & "', " & _
			           "titularccrc = '" & Replace(qryFOP.FieldByName("titularcartaopg").AsString, "'", "''") & "', " & _
			           "autorizacaoccrc = '" & autorizacao & "' " & _
			           "WHERE handle = " & qrySQL.FieldByName("acc").AsInteger)

			updACC.ExecSQL
			Commit
			Set updACC = Nothing
			GoTo FIM_ACC

ERRO_ACC:
			If InTransaction Then Rollback
			MsgBox "Erro ao atualizar ACC do PNR " & qrySQL.FieldByName("pnr").AsString & ": " & Err.Description
			Set updACC = Nothing

FIM_ACC:

			' 4 - Atualizar o PNR para reprocessamento
			Set updPNR = NewQuery

			On Error GoTo ERRO_PNR
			If Not InTransaction Then StartTransaction

			updPNR.Clear
			updPNR.Add("UPDATE VM_PNRS SET " & _
			           "SITUACAO = 1, " & _
			           "CONCLUIDO = 'S', " & _
			           "EXPORTADO = 'N', " & _
			           "AGUARDANDOEMISSAO = 'N' " & _
			           "WHERE HANDLE = " & qrySQL.FieldByName("pnr").AsInteger)

			updPNR.ExecSQL
			Commit
			Set updPNR = Nothing
			GoTo FIM_PNR

ERRO_PNR:
			If InTransaction Then Rollback
			MsgBox "Erro ao atualizar PNR " & qrySQL.FieldByName("pnr").AsString & ": " & Err.Description
			Set updPNR = Nothing

FIM_PNR:

		End If

		qrySQL.Next
	Loop

End Sub
