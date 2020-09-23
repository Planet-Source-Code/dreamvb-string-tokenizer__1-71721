VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "String Tokenizer"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "End"
      Height          =   420
      Left            =   4170
      TabIndex        =   7
      Top             =   3180
      Width           =   1605
   End
   Begin VB.ListBox LstToks 
      Height          =   2205
      Left            =   150
      TabIndex        =   6
      Top             =   2700
      Width           =   3915
   End
   Begin VB.TextBox txtPatten 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   150
      TabIndex        =   4
      Text            =   "*[a-zA-Z0-9=+/]*"
      Top             =   2010
      Width           =   4860
   End
   Begin VB.TextBox Text1 
      Height          =   1290
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Frmmain.frx":0000
      Top             =   375
      Width           =   5640
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Tokenize"
      Height          =   420
      Left            =   4170
      TabIndex        =   0
      Top             =   2685
      Width           =   1605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tokens:"
      Height          =   195
      Left            =   165
      TabIndex        =   5
      Top             =   2445
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Patten Match"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   1770
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "String to Token"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   105
      Width           =   1095
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sTok As New cToken

Private Sub cmdExit_Click()
    Set sTok = Nothing
    Unload frmmain
End Sub

Private Sub cmdOk_Click()
    LstToks.Clear
    
    With sTok
        Call .Init
        'Source to scan
        .Source = Text1.Text
        'Patten to find
        .Patten = txtPatten.Text
        'Loop tho tokens
        Do Until (.Token = Chr(0))
            'get token
            Call .GetToken
            'If not end token add to listbox
            If (.Token <> Chr(0)) Then
                LstToks.AddItem .Token
            End If
        Loop
    End With
    
End Sub

Private Sub Command1_Click()
    Unload frmmain
End Sub
