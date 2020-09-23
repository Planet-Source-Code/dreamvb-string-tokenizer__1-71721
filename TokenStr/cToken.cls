VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mStr As String
Private mToken As String
Private mPatten As String
Private sPos As Long

Public Property Get Patten() As String
    Patten = mPatten
End Property

Public Property Let Patten(ByVal vNewPatten As String)
    mPatten = vNewPatten
End Property

Public Property Get Source() As String
    Source = mStr
End Property

Public Property Let Source(ByVal vNewSrc As String)
    mStr = vNewSrc
End Property

Public Property Get Token() As String
    Token = mToken
End Property

Public Sub GetToken()

    mToken = vbNullString
    
    'Exit if more than source
    If (sPos >= Len(Source)) Then
        mToken = Chr(0)
    End If
    
    'Skip over white spaces
    While Mid(Source, sPos, 1) = " " Or _
        Mid(Source, sPos, 1) = vbCr Or _
        Mid(Source, sPos, 1) = vbLf Or _
        Mid(Source, sPos, 1) = vbTab
        sPos = (sPos + 1)
    Wend
    
    'That the tokens that match the patten
    While Mid(Source, sPos, 1) Like Patten
        mToken = mToken + Mid(Source, sPos, 1)
        sPos = (sPos + 1)
    Wend

End Sub

Public Sub Init()
    sPos = 1
    mToken = vbNullString
End Sub

