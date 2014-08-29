Attribute VB_Name = "Module1"
Option Explicit

Type URL
    Scheme As String
    Host As String
    Port As Long
    URI As String
    Query As String
End Type
Public retURL As URL
    

' returns as type URL from a string
Function ExtractUrl(ByVal strUrl As String) As URL
    Dim intPos As Integer
    Dim intPos1 As Integer
    Dim intPos2 As Integer

    intPos1 = InStr(strUrl, "://")
    
    If intPos1 > 0 Then
        retURL.Scheme = Mid(strUrl, 1, intPos1 - 1)
        strUrl = Mid(strUrl, intPos1 + 3)
    End If

    intPos1 = InStr(strUrl, ":")
    intPos2 = InStr(strUrl, "/")
    
    If intPos1 > 0 And intPos1 < intPos2 Then
        retURL.Host = Mid(strUrl, 1, intPos1 - 1)
        If (IsNumeric(Mid(strUrl, intPos1 + 1, intPos2 - intPos1 - 1))) Then
                retURL.Port = CInt(Mid(strUrl, intPos1 + 1, intPos2 - intPos1 - 1))
        End If
    ElseIf intPos2 > 0 Then
        retURL.Host = Mid(strUrl, 1, intPos2 - 1)
    Else
        retURL.Host = strUrl
        retURL.URI = "/"
        
        ExtractUrl = retURL
        Exit Function
    End If
    
    strUrl = Mid(strUrl, intPos2)
    
    ' find a question mark ?
    intPos = InStr(strUrl, " ")
    intPos1 = InStr(intPos, " ")
    
    If intPos1 > 0 Then
        retURL.URI = Mid(strUrl, 1, intPos1 - 1)
        retURL.Query = Mid(strUrl, intPos1 + 1)
    Else
        retURL.URI = strUrl
    End If
    
    ExtractUrl = retURL

End Function

Function Extract(ByRef row() As String, ByRef searchfor As String, ByRef output As String, ByVal start As Long) As Long
Dim i As Long, j As Long    ' indice pour la recherche
Dim ni As Long              ' ligne en cours
Dim ntot As Long            ' nb de lignes dans string1 (séparateur = vbLf)
Dim n As Long               ' longueur de la ligne en cours
'
ntot = UBound(row)
output = ""
For ni = start To ntot
i = InStr(row(ni), searchfor)
If i = 0 And output <> "" Then
Exit For
ElseIf i > 0 Then
n = Len(row(ni))
j = i + Len(searchfor)
While (Mid$(row(ni), j, 1) = " ") And (j < n)
j = j + 1
Wend
If j < n Then
If output <> "" Then
output = output & vbCrLf
End If
output = output & Mid$(row(ni), j, 1)
While (Mid$(row(ni), j + 1, 1) <> vbLf) And (j < n)
j = j + 1
output = output & Mid$(row(ni), j, 1)
Wend
End If
End If
Next
Extract = ni
End Function
