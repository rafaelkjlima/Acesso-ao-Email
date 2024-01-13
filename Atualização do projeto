Attribute VB_Name = "Modulo1"
Sub Email()
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim planilhaAtual As String
    Dim caminhoSalvar As String
    Dim nomeArquivo As String
    
    ' Defina o nome da planilha atual
    planilhaAtual = ActiveWorkbook.FullName
    
    ' Defina o caminho para salvar a planilha
    caminhoSalvar = "\\servidormicrolins\F\DINAMICA\CONTROLE DE FALTAS\2024\INTERATIVO\"
    
    ' Crie um nome de arquivo baseado na data atual
    nomeArquivo = "Contato de falta " & Format(Date, "DDMMYYYY") & ".xlsm"
    
    ' Combine o caminho e o nome do arquivo
    caminhoSalvar = caminhoSalvar & nomeArquivo
    
    ' Verifique se o arquivo ja existe e o apague
    If Dir(caminhoSalvar) <> "" Then
        Kill caminhoSalvar ' Exclui o arquivo existente
    End If
    
    ' Salve a planilha no novo caminho
    ActiveWorkbook.SaveCopyAs caminhoSalvar
    
    ' Crie uma instância do aplicativo Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    
    ' Crie um novo item de e-mail
    Set outlookMail = outlookApp.CreateItem(0)
    
    ' Configure os detalhes do e-mail
    With outlookMail
        .Subject = "Contato de falta INTERATIVO " & Date
        .Body = "Prezada Coordenadora," & vbCrLf & vbCrLf & _
                "Espero que esta mensagem a encontre bem." & vbCrLf & vbCrLf & _
                "Estamos entrando em contato para informar sobre o relatório de contato de falta." & vbCrLf & vbCrLf & _
                "Cordialmente Agradeço," & vbCrLf & _
                "Setor Pedagógico"
        
        .To = "ped.fazendariogrande@microlins.com.br"
        
        ' Adicione a planilha como anexo
        .Attachments.Add caminhoSalvar
        
        .BCC = "fazendariogrande@microlins.com.br;rafael.microlinsfrg@gmail.com"
        
        .Send

    End With
    
    ' Limpe as referências
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    ActiveWorkbook.Close SaveChanges:=True
End Sub

