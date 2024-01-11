Sub Email()
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim planilhaAtual As String
    Dim caminhoSalvar As String
    Dim nomeArquivo As String
    
    ' Defina o nome da planilha atual
    planilhaAtual = ActiveWorkbook.FullName
    
    ' Defina o caminho para salvar a planilha
    caminhoSalvar = "\\caminho para salvar"
    
    ' Crie um nome de arquivo baseado na data atual
    nomeArquivo = "Nome do arquivo para salvar" & Format(Date, "DD.MM.YYYY") & ".xlsx"
        'não esquecer colocar o tipo de arquivo correto

    ' Combine o caminho e o nome do arquivo
    caminhoSalvar = caminhoSalvar & nomeArquivo
    
    ' Verifique se o arquivo já existe e exclua-o
    If Dir(caminhoSalvar) <> "" Then
        Kill caminhoSalvar ' Exclui o arquivo existente
    End If
    
    ' Salve a planilha no novo caminho
    ActiveWorkbook.SaveCopyAs caminhoSalvar
    
    ' Crie uma instância do aplicativo Outlook
    Set outlookApp = CreateObject("Outlook.Application")
        'no meu caso utilizei o outlook, porem pode usar o email de sua preferencia
    
    ' Crie um novo item de e-mail
    Set outlookMail = outlookApp.CreateItem(0)
    
    ' Configure os detalhes do e-mail
    With outlookMail
        'assunto concatenado com a data atual do sistema
        .Subject = "assunto " & Date
        'mensagem com quebras de linhas
        .Body = "mensagem " & vbCrLf & vbCrLf & _
        .To = "xxxxxxxxxxxx@gmail.com"
        
        ' Adicione a planilha como anexo
        .Attachments.Add caminhoSalvar

        'da input a todos os itens de CCO, caso queira colocar como CC, use o comando .CC
        .BCC = "yyyyyyyyyyyyy@gmail.com"
        
        .Send

    End With
    
    ' Limpe as referências
    Set outlookMail = Nothing
    Set outlookApp = Nothing
End Sub