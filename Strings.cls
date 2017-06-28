VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Methods"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function DLLChangeLog() As String
   DLLChangeLog = " Ver 1.0.0" & vbCrLf & _
                  "   Added:" & vbCrLf & _
                  "       DLLVersion()" & vbCrLf & _
                  "       DLLAuthor()" & vbCrLf & _
                  "       Encrypt()" & vbCrLf & _
                  "       Decrypt()" & vbCrLf & _
                  "       toHex()" & vbCrLf & _
                  "       toOct()" & vbCrLf & _
                  "       toBinary()" & vbCrLf & _
                  "       DecToBin()" & vbCrLf & _
                  " Ver 1.0.6" & vbCrLf & _
                  "   Added:" & vbCrLf & _
                  "       ShowDLLInfo()" & vbCrLf & _
                  "       DLLChangeLog()" & vbCrLf & _
                  "   Modified:" & vbCrLf & _
                  "       DLLVersion()"
End Function

Public Function DLLVersion() As String
    DLLVersion = App.Major & "." & App.Minor & "." & App.Revision
End Function

Public Function DLLAuthor() As String
    DLLAuthor = "Anthoni Wiese"
End Function

Public Sub ShowDLLInfo()
    MsgBox "Strings DLL" & vbCrLf & "By: " & DLLAuthor & vbCrLf & "Ver: " & DLLVersion
End Sub

Public Sub ShowDLLChangelog()
    MsgBox DLLChangeLog()
End Sub

Public Function Encrypt(s As String, encmethod As String, Optional variable As String = "X") As String
    Dim n As Object
    Set n = CreateObject("MSScriptControl.ScriptControl.1")
    n.Language = "VBScript"
    
    
    For i = 1 To Len(s)
        Encrypt = Encrypt & Chr(n.Eval(Replace$(encmethod, variable, Asc(Mid$(s, i, 1)))))
    Next
End Function

Public Function Decrypt(s As String, decmethod As String, Optional variable As String = "X") As String
    Decrypt = Encrypt(s, decmethod, variable)
End Function

Public Function toHex(s As String, Optional padLength As Integer = 2) As String
    For i = 1 To Len(s)
        toHex = toHex & String(padLength - Len(CStr(Asc(Mid$(s, i, 1)))), "0") & Hex(Asc(Mid$(s, i, 1)))
    Next
End Function

Public Function toOct(s As String, Optional padLength As Integer = 3) As String
    For i = 1 To Len(s)
        toOct = toOct & String(padLength - Len(CStr(Asc(Mid$(s, i, 1)))), "0") & Oct(Asc(Mid$(s, i, 1)))
    Next
End Function

Public Function toBinary(s As String) As String
    For i = 1 To Len(s)
        toBinary = toBinary & DecToBin(Asc(Mid$(s, i, 1)))
    Next
End Function

Public Function DecToBin(n As Long, Optional padLength As Integer = 8) As String
    While n > 0
        DecToBin = DecToBin & CStr(n Mod 2)
        n = n / 2
    Wend
    DecToBin = StrReverse(DecToBin)
    DecToBin = String$(padLength - Len(DecToBin), "0")
End Function
