Rem Auth : Excelnoob.com
Const Token = "12345678910" 'isi dengan Token Bot yg dimiliki
Const sURL = "https://api.telegram.org/bot" & Token

Sub KirimPesan(Pesan As String, ChatID As String)
    Dim oHTTP As Object
    Dim Respon As String
    Dim PostURL As String
    
    PostURL = sURL & "/sendMessage?chat_id=" & ChatID & "&text=" & Pesan
    
    Set oHTTP = CreateObject("MSXML2.XMLHTTP")
    With oHTTP
        .Open "POST", PostURL, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .Send
        Respon = .ResponseText
    End With
    Debug.Print Respon
End Sub
