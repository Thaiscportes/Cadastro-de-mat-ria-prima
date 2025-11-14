# Cadastro-de-mat-ria-prima
Option Explicit

' Função para gerar o lote interno sequencial
Private Function GerarLoteInterno() As String
    Dim ultimaLinha As Long, ultimoNumero As Long
    With Sheets("BaseDados")
        ultimaLinha = .Cells(.Rows.Count, 1).End(xlUp).Row
        If ultimaLinha < 2 Then
            ultimoNumero = 0
        Else
            On Error Resume Next
            ultimoNumero = CLng(.Cells(ultimaLinha, 1).Value)
            If Err.Number <> 0 Then ultimoNumero = 0
            On Error GoTo 0
        End If
    End With
    GerarLoteInterno = Format(ultimoNumero + 1, "00000")
End Function



' Registrar alterações no histórico
Private Sub RegistrarHistorico(ByVal tipo As String, ByVal lote As String, _
    ByVal campo As String, ByVal antigo As String, ByVal novo As String, ByVal usuario As String)
    Dim shtHist As Worksheet, novaLinha As Long
    Set shtHist = Sheets("HISTORICO")
    novaLinha = shtHist.Cells(shtHist.Rows.Count, 1).End(xlUp).Row + 1
    shtHist.Cells(novaLinha, 1).Value = tipo
    shtHist.Cells(novaLinha, 2).Value = lote
    shtHist.Cells(novaLinha, 3).Value = campo
    shtHist.Cells(novaLinha, 4).Value = antigo
    shtHist.Cells(novaLinha, 5).Value = novo
    shtHist.Cells(novaLinha, 6).Value = usuario
    shtHist.Cells(novaLinha, 7).Value = Now
End Sub



Private Sub UserForm_Initialize()
    cboTipo.List = Array("Comum", "Comum com restrição", "Teste", "Teste com restrição")
    cboFiltroTipo.List = Array("Todos", "Vencidos", "Vence hoje", "A vencer")
    cboFiltroTipo = "Todos"
    txtValidade.Value = ""
    txtLoteInterno = GerarLoteInterno()
    txtLoteInterno.Locked = True

    With lstCadastros
        .Clear
        .ColumnCount = 9
        .ColumnWidths = "60;60;100;100;80;80;150;120;90"
    End With
    CarregarCadastros "Todos"
End Sub

Private Sub CarregarCadastros(filtro As String)
    Dim shtBase As Worksheet, i As Long, ultimaLinha As Long
    Dim dtValidade As Date, alerta As String, incluir As Boolean
    Set shtBase = Sheets("BaseDados")
    lstCadastros.Clear
    ultimaLinha = shtBase.Cells(shtBase.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultimaLinha
        incluir = False
        If IsDate(shtBase.Cells(i, 6).Value) Then
            dtValidade = shtBase.Cells(i, 6).Value
        Else
            dtValidade = 0
        End If
        Select Case filtro
            Case "Todos": incluir = True
            Case "Vencidos": If dtValidade < Date Then incluir = True
            Case "Vence hoje": If dtValidade = Date Then incluir = True
            Case "A vencer": If dtValidade > Date Then incluir = True
        End Select
        If incluir Then
            alerta = ""
            If dtValidade < Date Then
                alerta = "? Vencido"
            ElseIf dtValidade = Date Then
                alerta = "? Vence hoje"
            ElseIf dtValidade <= Date + 60 Then
                alerta = "? Vence em breve"
            End If
            lstCadastros.AddItem shtBase.Cells(i, 1).Value
            lstCadastros.List(lstCadastros.ListCount - 1, 1) = shtBase.Cells(i, 2).Value
            lstCadastros.List(lstCadastros.ListCount - 1, 2) = shtBase.Cells(i, 3).Value
            lstCadastros.List(lstCadastros.ListCount - 1, 3) = shtBase.Cells(i, 4).Value
            lstCadastros.List(lstCadastros.ListCount - 1, 4) = shtBase.Cells(i, 5).Value
            lstCadastros.List(lstCadastros.ListCount - 1, 5) = shtBase.Cells(i, 6).Value
            lstCadastros.List(lstCadastros.ListCount - 1, 6) = shtBase.Cells(i, 7).Value
            lstCadastros.List(lstCadastros.ListCount - 1, 7) = shtBase.Cells(i, 8).Value
            lstCadastros.List(lstCadastros.ListCount - 1, 8) = alerta
            lstCadastros.List(lstCadastros.ListCount - 1, 7) = shtBase.Cells(i, 8).Value ' já com kg

        End If
    Next i
End Sub


  Private Sub btnSalvar_Click()
    Dim shtBase As Worksheet, novaLinha As Long, existente As Range
    Dim OutApp As Object, OutMail As Object
    Dim corpoEmail As String, caminhoPDF As String
    Dim entryID As String
    Dim erroEnvio As Boolean
    
    If txtMP = "" Or cboTipo = "" Then
        MsgBox "Preencha os campos obrigatórios.", vbExclamation
        Exit Sub
    End If
    
    Set shtBase = Sheets("BaseDados")
    Set existente = shtBase.Range("A:A").Find(txtLoteInterno.Value, , xlValues, xlWhole)
    If Not existente Is Nothing Then
        MsgBox "Este lote já existe.", vbCritical
        Exit Sub
    End If
    
    novaLinha = shtBase.Cells(shtBase.Rows.Count, 1).End(xlUp).Row + 1
    shtBase.Cells(novaLinha, 1).Value = txtLoteInterno
    shtBase.Cells(novaLinha, 2).Value = cboTipo
    shtBase.Cells(novaLinha, 3).Value = txtMP
    shtBase.Cells(novaLinha, 4).Value = txtLoteFornecedor
    shtBase.Cells(novaLinha, 5).Value = txtNotaFiscal
    shtBase.Cells(novaLinha, 6).Value = txtValidade
    shtBase.Cells(novaLinha, 7).Value = txtObs
    shtBase.Cells(novaLinha, 8).Value = Now
    shtBase.Cells(novaLinha, 9).Value = txtCaminhoPDF.Tag ' Caminho do PDF
    shtBase.Cells(novaLinha, 8).Value = txtQuantidade.Value ' já com kg

    
    ' Registrar no histórico
    RegistrarHistorico "Cadastro", txtLoteInterno, "-", "-", "Criado", Environ("username")
    
    ' Envio de e-mail
    erroEnvio = False
    On Error GoTo FalhaEnvio
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    caminhoPDF = txtCaminhoPDF.Tag
    
    corpoEmail = "Novo cadastro de matéria-prima realizado:" & vbCrLf & vbCrLf & _
    "Lote Interno: " & txtLoteInterno & vbCrLf & _
    "Tipo: " & cboTipo & vbCrLf & _
    "Matéria-prima: " & txtMP & vbCrLf & _
    "Lote Fornecedor: " & txtLoteFornecedor & vbCrLf & _
    "Nota Fiscal: " & txtNotaFiscal & vbCrLf & _
    "Validade: " & txtValidade & vbCrLf & _
    "Observações: " & txtObs & vbCrLf & _
    "Data/Hora do Cadastro: " & Now & vbCrLf & _
    "Quantidade: " & txtQuantidade.Value & vbCrLf

    With OutMail
        .To = "almoxarifado@dapp.com.br; lab1@dapp.com.br"
        .Subject = "Cadastro de MP: " & txtMP & " | Lote: " & txtLoteInterno
        .Body = corpoEmail
        
        If caminhoPDF <> "" And Dir(caminhoPDF) <> "" Then
            .Attachments.Add caminhoPDF
        End If
        
        .Send
    End With
    
    ' Tenta obter EntryID para histórico sem travar caso falhe
    On Error Resume Next
    entryID = OutMail.entryID
    If Err.Number = 0 Then
        shtBase.Cells(novaLinha, 10).Value = entryID
    End If
    On Error GoTo 0
    
    GoTo Sucesso
    
FalhaEnvio:
    erroEnvio = True
    Resume Proximo
    
Proximo:
    On Error GoTo 0

Sucesso:
    MsgBox "Cadastro salvo com sucesso!" & IIf(erroEnvio, vbCrLf & "Não foi possível confirmar o envio do e-mail. Caso tenha recebido, pode ignorar.", vbCrLf & "E-mail enviado com sucesso!"), vbInformation
    
    CarregarCadastros cboFiltroTipo.Value
    btnLimpar_Click
    
    ' Limpeza de objetos Outlook
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub



Private Sub btnEditar_Click()
    Dim shtBase As Worksheet, linha As Long
    Dim antigoTipo As String, antigoMP As String, antigoLoteFornecedor As String
    Dim antigoNota As String, antigoValidade As String, antigoObs As String
    Dim antigoCaminhoPDF As String
    Dim OutApp As Object, OutMail As Object
    Dim corpoEmail As String, alteracoes As String
    Dim caminhoPDF As String
    Dim entryID As String
    Dim ns As Object, origMail As Object, replyMail As Object

    If txtLoteInterno = "" Then
        MsgBox "Informe um lote para editar.", vbExclamation
        Exit Sub
    End If

    Set shtBase = Sheets("BaseDados")
    On Error Resume Next
    linha = shtBase.Range("A:A").Find(txtLoteInterno.Value, , xlValues, xlWhole).Row
    On Error GoTo 0

    If linha = 0 Then
        MsgBox "Lote não encontrado para edição.", vbExclamation
        Exit Sub
    End If

    antigoTipo = shtBase.Cells(linha, 2).Value
    antigoMP = shtBase.Cells(linha, 3).Value
    antigoLoteFornecedor = shtBase.Cells(linha, 4).Value
    antigoNota = shtBase.Cells(linha, 5).Value
    antigoValidade = shtBase.Cells(linha, 6).Value
    antigoObs = shtBase.Cells(linha, 7).Value
    antigoCaminhoPDF = shtBase.Cells(linha, 9).Value

    alteracoes = ""

    If antigoTipo <> cboTipo Then
        alteracoes = alteracoes & "- Tipo: " & antigoTipo & " -> " & cboTipo & vbCrLf
    End If
    If antigoMP <> txtMP Then
        alteracoes = alteracoes & "- MP: " & antigoMP & " -> " & txtMP & vbCrLf
    End If
    If antigoLoteFornecedor <> txtLoteFornecedor Then
        alteracoes = alteracoes & "- Lote Fornecedor: " & antigoLoteFornecedor & " -> " & txtLoteFornecedor & vbCrLf
    End If
    If antigoNota <> txtNotaFiscal Then
        alteracoes = alteracoes & "- Nota Fiscal: " & antigoNota & " -> " & txtNotaFiscal & vbCrLf
    End If
    If antigoValidade <> txtValidade Then
        alteracoes = alteracoes & "- Validade: " & antigoValidade & " -> " & txtValidade & vbCrLf
    End If
    If antigoObs <> txtObs Then
        alteracoes = alteracoes & "- Observações: " & antigoObs & " -> " & txtObs & vbCrLf
    End If
    If Trim(antigoCaminhoPDF) <> Trim(txtCaminhoPDF.Tag) Then
        alteracoes = alteracoes & "- PDF alterado" & vbCrLf
    End If

    If alteracoes = "" Then
        MsgBox "Nenhuma alteração detectada.", vbInformation
        Exit Sub
    End If

    shtBase.Cells(linha, 2).Value = cboTipo
    shtBase.Cells(linha, 3).Value = txtMP
    shtBase.Cells(linha, 4).Value = txtLoteFornecedor
    shtBase.Cells(linha, 5).Value = txtNotaFiscal
    shtBase.Cells(linha, 6).Value = txtValidade
    shtBase.Cells(linha, 7).Value = txtObs
    shtBase.Cells(linha, 9).Value = txtCaminhoPDF.Tag
    shtBase.Cells(linha, 8).Value = txtQuantidade.Value


    RegistrarHistorico "Edição", txtLoteInterno, "Vários campos", "-", "-", Environ("username")

    caminhoPDF = Trim(txtCaminhoPDF.Tag)
    entryID = shtBase.Cells(linha, 10).Value

    On Error GoTo ErroEmail
    Set OutApp = CreateObject("Outlook.Application")
    Set ns = OutApp.GetNamespace("MAPI")

    If entryID <> "" Then
        Set origMail = ns.GetItemFromID(entryID)
        Set replyMail = origMail.Reply

        corpoEmail = "O cadastro do lote interno " & txtLoteInterno & " foi alterado por " & Environ("username") & ":" & vbCrLf & vbCrLf & _
                     alteracoes & vbCrLf & _
                     "Data/Hora da edição: " & Now

        With replyMail
            .Body = corpoEmail & vbCrLf & .Body
            If caminhoPDF <> "" And Dir(caminhoPDF) <> "" Then
                .Attachments.Add caminhoPDF
            End If
            .Send
        End With

        Set replyMail = Nothing
        Set origMail = Nothing
    Else
        Set OutMail = OutApp.CreateItem(0)

        corpoEmail = "O cadastro do lote interno " & txtLoteInterno & " foi alterado por " & Environ("username") & ":" & vbCrLf & vbCrLf & _
                     alteracoes & vbCrLf & _
                     "Data/Hora da edição: " & Now


        With OutMail
             .To = "almoxarifado@dapp.com.br; lab1@dapp.com.br"
            .Subject = "Alteração no Lote: " & txtLoteInterno & " | MP: " & txtMP
            .Body = corpoEmail
            If caminhoPDF <> "" And Dir(caminhoPDF) <> "" Then
                .Attachments.Add caminhoPDF
            End If
            .Send
        End With
        Set OutMail = Nothing
    End If

    Set ns = Nothing
    Set OutApp = Nothing

    CarregarCadastros cboFiltroTipo.Value
    MsgBox "Cadastro atualizado e e-mail enviado com sucesso!"
    Exit Sub

ErroEmail:
    MsgBox "Cadastro atualizado, mas houve erro ao enviar o e-mail: " & Err.Description, vbExclamation
    CarregarCadastros cboFiltroTipo.Value
End Sub

Private Sub btnExcluir_Click()
    Dim shtBase As Worksheet, linha As Long, resposta As VbMsgBoxResult
    If txtLoteInterno = "" Then Exit Sub
    resposta = MsgBox("Deseja realmente excluir este lote?", vbYesNo + vbQuestion)
    If resposta = vbNo Then Exit Sub

    Set shtBase = Sheets("BaseDados")
    On Error Resume Next
    linha = shtBase.Range("A:A").Find(txtLoteInterno.Value, , xlValues, xlWhole).Row
    On Error GoTo 0

    If linha = 0 Then
        MsgBox "Lote não encontrado para exclusão.", vbExclamation
        Exit Sub
    End If

    shtBase.Rows(linha).Delete

    RegistrarHistorico "Exclusão", txtLoteInterno, "-", "-", "Excluído", Environ("username")
    CarregarCadastros cboFiltroTipo.Value
    btnLimpar_Click
    MsgBox "Lote excluído."
End Sub

Private Sub btnPesquisar_Click()
    Dim cel As Range, shtBase As Worksheet
    Set shtBase = Sheets("BaseDados")
    Set cel = shtBase.Range("A:A").Find(txtLoteInterno.Value, , xlValues, xlWhole)
    If cel Is Nothing Then
        MsgBox "Lote não encontrado.", vbExclamation
        Exit Sub
    End If
    txtLoteInterno = cel.Value
    cboTipo = cel.Offset(0, 1).Value
    txtMP = cel.Offset(0, 2).Value
    txtLoteFornecedor = cel.Offset(0, 3).Value
    txtNotaFiscal = cel.Offset(0, 4).Value
    txtValidade = cel.Offset(0, 5).Value
    txtObs = cel.Offset(0, 6).Value
    txtCaminhoPDF.Tag = cel.Offset(0, 8).Value
    txtQuantidade = cel.Offset(0, 7).Value ' já vem com "kg"


    If txtCaminhoPDF.Tag <> "" Then
        txtCaminhoPDF.Value = "PDF ADICIONADO - CLIQUE EM ABRIR PDF"
    Else
        txtCaminhoPDF.Value = ""
    End If
End Sub

Private Sub btnLimpar_Click()
    txtMP = ""
    cboTipo = ""
    txtLoteInterno = GerarLoteInterno()
    txtLoteFornecedor = ""
    txtNotaFiscal = ""
    txtValidade = ""
    txtObs = ""
    txtCaminhoPDF.Value = ""
    txtCaminhoPDF.Tag = ""
    txtQuantidade = ""

End Sub

Private Sub cboFiltroTipo_Change()
    CarregarCadastros cboFiltroTipo.Value
End Sub

Private Sub btnFechar_Click()
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("Deseja realmente fechar? Seus dados serão salvos automaticamente.", vbYesNo + vbQuestion, "Fechar o sistema")
    If resposta = vbYes Then
        ThisWorkbook.Save
        Application.Quit
    End If
End Sub



Private Sub btnAnexarPDF_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecione o PDF para anexar"
        .Filters.Clear
        .Filters.Add "Arquivos PDF", "*.pdf"
        .AllowMultiSelect = False
        If .Show = -1 Then
            txtCaminhoPDF.Tag = .SelectedItems(1)
            txtCaminhoPDF.Value = "PDF ADICIONADO - CLIQUE EM ABRIR PDF"
            MsgBox "PDF selecionado para este cadastro!", vbInformation
        End If
    End With
End Sub

Private Sub btnAbrirPDF_Click()
    Dim caminho As String
    caminho = Trim(txtCaminhoPDF.Tag)
    If caminho = "" Then
        MsgBox "Nenhum PDF vinculado a este cadastro.", vbExclamation
        Exit Sub
    End If
    If Dir(caminho) = "" Then
        MsgBox "Arquivo PDF não encontrado no caminho:" & vbCrLf & caminho, vbCritical
        Exit Sub
    End If
    ThisWorkbook.FollowHyperlink caminho
End Sub

Private Sub lstCadastros_Click()
    Dim linha As Long
    Dim sht As Worksheet
    Set sht = Sheets("BaseDados")
    If lstCadastros.ListIndex = -1 Then Exit Sub
    linha = lstCadastros.ListIndex + 2
    With sht
        txtLoteInterno.Value = .Cells(linha, 1).Value
        cboTipo.Value = .Cells(linha, 2).Value
        txtMP.Value = .Cells(linha, 3).Value
        txtLoteFornecedor.Value = .Cells(linha, 4).Value
        txtNotaFiscal.Value = .Cells(linha, 5).Value
        txtValidade.Value = .Cells(linha, 6).Value
        txtObs.Value = .Cells(linha, 7).Value
        txtCaminhoPDF.Tag = .Cells(linha, 9).Value
        txtQuantidade = .Cells(linha, 8).Value

        If txtCaminhoPDF.Tag <> "" Then
            txtCaminhoPDF.Value = "PDF ADICIONADO - CLIQUE EM ABRIR PDF"
        Else
            txtCaminhoPDF.Value = ""
        End If
    End With
End Sub


Private Sub txtQuantidade_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim valor As Double
    If IsNumeric(Replace(txtQuantidade.Value, " kg", "")) Then
        valor = CDbl(Replace(txtQuantidade.Value, " kg", ""))
        txtQuantidade.Value = Format(valor, "#,##0.00") & " kg"
    End If
End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Application.Visible = True ' Torna o Excel visível novamente
    ThisWorkbook.Close SaveChanges:=False ' Altere para True se quiser salvar ao fechar
End Sub

