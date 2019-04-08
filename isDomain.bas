' EBNF syntax of Domain Names http://tools.ietf.org/html/rfc1035
' <fqdn>        ::= <subdomain> <domain> "." <tld>
' <domain>      ::= <letter> [ <ldh-str> ] <let-dig>
' <tld>         ::= <letter> {2,6}
' <subdomain>   ::= <epsilon> | <label> "." <subdomain>
' <label>       ::= <letter> [ [ <ldh-str> ] <let-dig> ]
' <ldh-str>     ::= <let-dig-hyp> | <let-dig-hyp> <ldh-str>
' <let-dig-hyp> ::= <let-dig> | "-"
' <let-dig>     ::= <letter> | <digit>
' <letter>      ::= [a-z]
' <digit>       ::= [0-9]

Public Function isDomain(ByVal dom As String) As Boolean

    Dim i As Integer, l As Integer, s As Byte
    
    l = Len(dom)
    isDomain = False
    If l > 63 Then Exit Function
    
    s = 1
    For i = l To 1 Step -1
        Select Case Asc(Mid(dom, i))

        Case 97 To 122:                         'char [a-z]
            If s = 10 Or s = 11 Then
                s = 10
            ElseIf s >= 13 And s <= 15 Then
                s = 13
            ElseIf s = 7 Then
                Exit Function
            Else
                s = s + 1
            End If

        Case 48 To 57:                          'digit [0-9]
            If s >= 8 And s <= 11 Then
                s = 9
            ElseIf s >= 12 And s <= 15 Then
                s = 14
            Else
                Exit Function
            End If

        Case 45:                                'dash [-]
            If s = 9 Or s = 10 Then
                s = 11
            ElseIf s = 13 Or s = 14 Then
                s = 15
            Else
                Exit Function
            End If

        Case 46:                                'dot [.]
            If s >= 3 And s <= 7 Then
                s = 8
            ElseIf s = 10 Or s = 13 Then
                s = 12
            Else
                Exit Function
            End If

        Case Else
            Exit Function

        End Select
    Next i

    isDomain = (s = 10 Or s = 13)

End Function
