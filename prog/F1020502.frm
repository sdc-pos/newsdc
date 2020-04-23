VERSION 5.00
Begin VB.Form F1020502 
   BackColor       =   &H00FFFFFF&
   Caption         =   "íIî‘ï å¬ëïî†ÅEÉâÉìÉNè∆âÔ"
   ClientHeight    =   6315
   ClientLeft      =   2130
   ClientTop       =   2835
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "ÇlÇr ÉSÉVÉbÉN"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   11280
   StartUpPosition =   2  'âÊñ ÇÃíÜâõ
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'µÃå≈íË
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'µÃå≈íË
      Index           =   2
      Left            =   4560
      MaxLength       =   2
      TabIndex        =   2
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'µÃå≈íË
      Index           =   3
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   3
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'µÃå≈íË
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   2052
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'µÃå≈íË
      Index           =   4
      Left            =   6360
      MaxLength       =   2
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'µÃå≈íË
      Index           =   5
      Left            =   7080
      MaxLength       =   2
      TabIndex        =   5
      Top             =   240
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   3900
      Index           =   0
      ItemData        =   "F1020502.frx":0000
      Left            =   2400
      List            =   "F1020502.frx":0002
      TabIndex        =   7
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   6
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command 
      Caption         =   "ëOâÊñ "
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   10320
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   9480
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   8640
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   7800
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6480
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   5640
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4800
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3960
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2640
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "ç≈Å@êV"
      BeginProperty Font 
         Name            =   "ÇlÇr ÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "Å`"
      Height          =   375
      Index           =   2
      Left            =   6720
      TabIndex        =   25
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "òA"
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   24
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "ëqå…"
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "óÒ"
      Height          =   375
      Index           =   6
      Left            =   3960
      TabIndex        =   22
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'íÜâõëµÇ¶
      BackColor       =   &H00FFFFFF&
      Caption         =   "Å`"
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   21
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "å¬ëïî†áÇ"
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   20
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "F1020502"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxSoko_No% = 0
Private Const ptxSoko_Name% = 1
Private Const ptxstRetu% = 2
Private Const ptxenRetu% = 3
Private Const ptxstRen% = 4
Private Const ptxenRen% = 5
Private Const ptxPacking_No% = 6

Private Const Text_Max% = 6

Private Const plstPacking_No% = 0

Private Sub Command_Click(Index As Integer)

Dim yn      As Integer
Dim sts     As Integer
    
    Select Case Index
        
        Case 0                  'ç≈êV
            
            If List_Disp() Then
                F1011002.Hide
            End If
            
                    
        Case 11                 'èIóπ
    
            F1011002.Hide
    
    End Select
End Sub

Private Sub Form_Activate()

    If List_Disp() Then
        F1011002.Hide
    End If

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   ÇjÇÖÇô ÇcÇèÇóÇé ëOèàóù
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command(KeyCode - vbKeyF1).Value = True
    End Select


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If

End Sub

Private Sub Text_GotFocus(Index As Integer)
    
    If Text(Index).TabStop = True Then
        Text(Index) = Trim(Text(Index).Text)
        Text(Index).SelStart = 0
        Text(Index).SelLength = Len(Text(Index).Text)
    End If


End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim i   As Integer
Dim sts As Integer

    If KeyCode <> vbKeyReturn Then Exit Sub
            
    
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i


End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   âÊñ çÄñ⁄ÉçÉbÉNÅiÉCÉxÉìÉgéÊìæïsâ¬Åj
'----------------------------------------------------------------------------
Dim i As Integer

    F1011002.MousePointer = vbHourglass

    Call Ctrl_Lock(F1011002)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   âÊñ çÄñ⁄ÉçÉbÉNâèúÅiÉCÉxÉìÉgéÊìæâ¬Åj
'----------------------------------------------------------------------------
Dim i As Integer

    Call Ctrl_UnLock(F1011002)


    F1011002.MousePointer = vbDefault

End Sub

Private Function List_Disp() As Integer

Dim com         As Integer
Dim Edit        As String
Dim sts         As Integer
    
Dim Save_Soko   As String * 2
Dim Save_Retu   As String * 2
Dim Save_Ren    As String * 2

Dim Skip_Flg    As Boolean
    

    List_Disp = True

    List1(plstPacking_No).Clear

    Call UniCode_Conv(K0_TPACKING.Soko_No, Text(ptxSoko_No).Text)
    Call UniCode_Conv(K0_TPACKING.Retu, Text(ptxstRetu).Text)
    Call UniCode_Conv(K0_TPACKING.Ren, Text(ptxstRen).Text)
    Call UniCode_Conv(K0_TPACKING.PACKING_NO, Text(ptxPacking_No))
    
    com = BtOpGetGreaterEqual
    
    Do
    
    
        sts = BTRV(com, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
        Select Case sts
            Case BtNoErr
                
                If (StrConv(TPACKINGREC.Soko_No, vbUnicode) <> Text(ptxSoko_No).Text) Or _
                    (StrConv(TPACKINGREC.Retu, vbUnicode) > Text(ptxenRetu).Text) Then
                    
                    Exit Do
                    
                End If
                
                
            Case BtErrEOF
                                
                Exit Do
                
            Case Else
                    
                Call File_Error(sts, com, "íIï å¬ëïî†É}ÉXÉ^")
                F1011002.Hide
            
            
        End Select
    
    
        Skip_Flg = False
        
        If (StrConv(TPACKINGREC.Ren, vbUnicode) < Text(ptxstRen).Text) Or _
            (StrConv(TPACKINGREC.Ren, vbUnicode) > Text(ptxenRen).Text) Then
            Skip_Flg = True
        End If
        
        If Len(Trim(Text(ptxPacking_No).Text)) = 0 Then
        Else
            If Trim(StrConv(TPACKINGREC.PACKING_NO, vbUnicode)) <> Trim(Text(ptxPacking_No).Text) Then
                Skip_Flg = True
            End If
        End If
    
        If Not Skip_Flg Then
            If Save_Soko <> StrConv(TPACKINGREC.Soko_No, vbUnicode) Or _
                Save_Retu <> StrConv(TPACKINGREC.Retu, vbUnicode) Or _
                Save_Ren <> StrConv(TPACKINGREC.Ren, vbUnicode) Then
                Edit = StrConv(TPACKINGREC.Retu, vbUnicode) & "-" & StrConv(TPACKINGREC.Ren, vbUnicode) & "     "
            Else
                Edit = "          "
            End If
    
            Edit = Edit & StrConv(TPACKINGREC.PACKING_NO, vbUnicode) & "      "
            Edit = Edit & StrConv(TPACKINGREC.RANK, vbUnicode)
    
            List1(plstPacking_No).AddItem Edit
       
       
            Save_Soko = StrConv(TPACKINGREC.Soko_No, vbUnicode)
            Save_Retu = StrConv(TPACKINGREC.Retu, vbUnicode)
            Save_Ren = StrConv(TPACKINGREC.Ren, vbUnicode)
       
        End If

       
        com = BtOpGetNext
    Loop
    
    List_Disp = False

End Function
