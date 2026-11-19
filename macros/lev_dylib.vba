Option Explicit

#If Mac Then
    Private Declare PtrSafe Function levd Lib "~/Library/Group Containers/UBF8T346G9.Office/lib/liblev.dylib" _
        (ByVal str1 As String, ByVal str2 As String) As Long
#End If

' 1. Gibt die Levenshtein-Distanz als Zahl zurück (0 = identisch)
Public Function LevDist(ByVal text1 As String, ByVal text2 As String) As Long
    LevDist = levd(text1, text2)
End Function

' 2. Test-Makro
Sub TesteStringVergleich()
    Dim dist As Long
    dist = LevDist("Hallo Welt", "Hallo Welt")
    
    If dist = 0 Then
        MsgBox "Die Strings sind identisch! (Distanz: 0)", vbInformation, "C-dylib"
    Else
        MsgBox "Die Strings unterscheiden sich. Distanz: " & dist, vbExclamation, "C-dylib"
    End If
End Sub


