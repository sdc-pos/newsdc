VERSION 5.00
Begin VB.Form Conv_PI000701 
   Caption         =   "���i���V�X�e���@�O�؍폜����"
   ClientHeight    =   6300
   ClientLeft      =   1920
   ClientTop       =   2430
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   11310
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   7560
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   6000
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   3660
      Index           =   0
      ItemData        =   "Conv_PI000701.frx":0000
      Left            =   240
      List            =   "Conv_PI000701.frx":0002
      TabIndex        =   5
      Top             =   1800
      Width           =   8895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   240
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   17
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   9
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�X �V"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "�d���P��"
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   23
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "�d����"
      Height          =   255
      Index           =   3
      Left            =   6600
      TabIndex        =   22
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "�O�ؐ�"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   21
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "�O�ؓ�"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   20
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblBikou 
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   4560
      TabIndex        =   19
      Top             =   1560
      Width           =   135
   End
   Begin VB.Label Label 
      Caption         =   "�i��"
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   18
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "Conv_PI000701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'�e�L�X�g�p�Y��
Private Const ptxHIN_GAI% = 0               '�i��
Private Const ptxNYUKA_DT% = 1              '�O�ؓ�
Private Const ptxNYUKA_QTY% = 2             '�O�ؐ�

Private Const ptxSHIIRE_CODE% = 3           '�d����
Private Const ptxSHIIRE_TANKA% = 4          '�d���P��
'���X�g�a�n�w�p�Y��
Private Const plstNYU% = 0



Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    Conv_PI000701.MousePointer = vbHourglass

    Call Ctrl_Lock(Conv_PI000701)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(Conv_PI000701)


    Conv_PI000701.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
Dim com As Integer
Dim sts As Integer

    
    Error_Check_Proc = True
    
        
    Select Case Mode
    
        Case ptxHIN_GAI
    
    
            Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
            Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
            Call UniCode_Conv(K0_P_NYU.HIN_GAI, Text1(ptxHIN_GAI).Text)
    
            sts = BTRV(BtOpGetEqual, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
            Select Case sts
                Case BtNoErr
                    
                    Text1(ptxNYUKA_DT).Text = Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                                Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                                Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 7, 2)
                    Text1(ptxNYUKA_QTY).Text = Format(CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)), "#,##0")
                
                
                Case BtErrKeyNotFound
                
                    Text1(ptxNYUKA_DT).Text = ""
                    Text1(ptxNYUKA_QTY).Text = ""
                
                    MsgBox "���͂������ڂ̓G���[�ł��B"
                    Text1(ptxHIN_GAI).SetFocus
                    Exit Function
                
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "���ޑO���ް�")
                    Exit Function
            End Select
    
    
    
    
    
    End Select
        
        
        
        
    Error_Check_Proc = False
    

End Function
Private Function Item_Disp_Proc(CODE As String) As Integer
'----------------------------------------------------------------------------
'                   ��ʕ\��
'----------------------------------------------------------------------------
Dim sts As Integer
Dim i   As Integer

    Item_Disp_Proc = True
    
    Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_NYU.HIN_GAI, Text1(ptxHIN_GAI).Text)

    sts = BTRV(BtOpGetEqual, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    Select Case sts
        Case BtNoErr
            
            Text1(ptxNYUKA_DT).Text = Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 7, 2)
            Text1(ptxNYUKA_QTY).Text = Format(CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)), "#,##0")
        
        
            Text1(ptxSHIIRE_CODE).Text = Trim(StrConv(P_NYUREC.SHIIRE_CODE, vbUnicode))
            If IsNumeric(StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode)) Then
                Text1(ptxSHIIRE_TANKA).Text = Format(CDbl((StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode))), "#,##0.00")
            Else
                Text1(ptxSHIIRE_TANKA).Text = ""
            End If
        
        Case BtErrKeyNotFound
        
            Text1(ptxNYUKA_DT).Text = ""
            Text1(ptxNYUKA_QTY).Text = ""
            Text1(ptxSHIIRE_CODE).Text = ""
            Text1(ptxSHIIRE_TANKA).Text = ""
        
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Text1(ptxHIN_GAI).SetFocus
            Exit Function
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "���ޑO���ް�")
            Exit Function
    End Select

    Item_Disp_Proc = False

End Function
Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   ���X�g�{�b�N�X�\��
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim wkQty       As String
Dim wkTanka       As String

    List_Disp_Proc = True
    
    List1(plstNYU).Clear
    
    '���ޑO���ް��ǂݍ���
    
    com = BtOpGetFirst
    
    
    Do
    
        sts = BTRV(com, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
            
        Select Case sts
            Case BtNoErr
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "���ޑO���ް�")
                Exit Function
        
        End Select
        
        If IsNumeric(StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode)) Then
            wkTanka = Format(CDbl(StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode)), "#,##0.00")
            If Len(wkTanka) < 13 Then
                wkTanka = Space(13 - Len(wkTanka)) & wkTanka
            End If
        Else
            wkTanka = Space(13)
        End If
        
        
        
        wkQty = Format(CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)), "#,##0")
        If Len(wkQty) < 8 Then
            wkQty = Space(8 - Len(wkQty)) & wkQty
        End If
        
        List1(plstNYU).AddItem StrConv(P_NYUREC.HIN_GAI, vbUnicode) & "   " & _
                                    Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                    Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                    Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 7, 2) & "  " & _
                                    wkQty & "      " & _
                                    StrConv(P_NYUREC.SHIIRE_CODE, vbUnicode) & " " & _
                                    wkTanka
        
        com = BtOpGetNext
    Loop
        
    DoEvents

    List_Disp_Proc = False

End Function

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �O���ް��X�V
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim i       As Integer

    Update_Proc = True
    
    Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_NYU.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_P_NYU.NYUKA_DT, Format(Text1(ptxNYUKA_DT).Text, "YYYYMMDD"))

    sts = BTRV(BtOpGetEqual + BtSNoWait, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    
    Do
        Select Case sts
            Case BtNoErr
                
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Update_Proc = False
                    Exit Do
                End If
            
            Case BtErrKeyNotFound
                
                For i = ptxHIN_GAI To ptxSHIIRE_TANKA
                    Text1(i).Text = ""
                Next i
                
                Update_Proc = False
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ޑO���ް�")
                Exit Function
        End Select
    Loop
    
    If IsNumeric(Text1(ptxSHIIRE_TANKA).Text) Then
        Call UniCode_Conv(P_NYUREC.SHIIRE_TANKA, Format(CDbl(Text1(ptxSHIIRE_TANKA).Text), "00000000.00"))
    Else
        Call UniCode_Conv(P_NYUREC.SHIIRE_TANKA, "00000000.00")
    End If
    
    Call UniCode_Conv(P_NYUREC.SHIIRE_CODE, Text1(ptxSHIIRE_CODE).Text)
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpUpdate, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                For i = ptxHIN_GAI To ptxNYUKA_QTY
                    Text1(i).Text = ""
                Next i
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_NYU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
                    Update_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "���ޑO���ް�")
                Exit Function
        End Select
    Loop

    For i = ptxHIN_GAI To ptxSHIIRE_TANKA
        Text1(i).Text = ""
    Next i




    If List_Disp_Proc() Then
        Exit Function
    End If
    
    If List1(plstNYU).ListCount > 0 Then
        List1(plstNYU).SetFocus
        List1(plstNYU).ListIndex = 0
    Else
        Text1(ptxHIN_GAI).SetFocus
    End If


    Update_Proc = False


End Function
Private Function Delete_Proc() As Integer
'----------------------------------------------------------------------------
'                   �O���ް��폜
'----------------------------------------------------------------------------
Dim sts     As Integer
Dim ans     As Integer
Dim i       As Integer

    Delete_Proc = True
    
    Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_NYU.HIN_GAI, Text1(ptxHIN_GAI).Text)
    Call UniCode_Conv(K0_P_NYU.NYUKA_DT, Format(Text1(ptxNYUKA_DT).Text, "YYYYMMDD"))

    sts = BTRV(BtOpGetEqual + BtSNoWait, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    
    Do
        Select Case sts
            Case BtNoErr
                
                Exit Do
            
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_NYUKA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    Delete_Proc = False
                    Exit Do
                End If
            
            Case BtErrKeyNotFound
                
                For i = ptxHIN_GAI To ptxSHIIRE_TANKA
                    Text1(i).Text = ""
                Next i
                
                Delete_Proc = False
                Exit Function
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ޑO���ް�")
                Exit Function
        End Select
    Loop
    
    
    Do
        
        DoEvents
        
        sts = BTRV(BtOpDelete, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                For i = ptxHIN_GAI To ptxNYUKA_QTY
                    Text1(i).Text = ""
                Next i
                ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<P_NYU.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                If ans = vbCancel Then
                    sts = BTRV(BtOpUnlock, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
                    Delete_Proc = False
                    Exit Do
                End If
            Case Else
                Call File_Error(sts, BtOpDelete, "���ޑO���ް�")
                Exit Function
        End Select
    Loop

    For i = ptxHIN_GAI To ptxSHIIRE_TANKA
        Text1(i).Text = ""
    Next i




    If List_Disp_Proc() Then
        Exit Function
    End If
    
    If List1(plstNYU).ListCount > 0 Then
        List1(plstNYU).SetFocus
        List1(plstNYU).ListIndex = 0
    Else
        Text1(ptxHIN_GAI).SetFocus
    End If


    Delete_Proc = False


End Function

Private Sub Command1_Click(Index As Integer)

Dim ans As Integer
Dim i   As Integer

    Select Case Index
        Case P_CMD_Upd                      '�X�V
            ans = MsgBox("�X�V���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Update_Proc() Then
                    Unload Me
                End If
            End If
                    
        
        
        Case P_CMD_DEL                      '�폜
            ans = MsgBox("�폜���܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Delete_Proc() Then
                    Unload Me
                End If
            End If
        Case P_CMD_DSP                      '����/�\��
        Case P_CMD_OUT                      '�ް��o��
        Case P_CMD_PRT                      '���
        
        Case P_CMD_End                      '�I��
            Unload Me
    End Select

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �j���� �c������ �O����
'----------------------------------------------------------------------------
    Select Case KeyCode
        Case vbKeyF1 To vbKeyF12
            Command1(KeyCode - vbKeyF1).Value = True
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c       As String * 128
Dim i       As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If

                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                '���ޑO�؂n�o�d�m
    If P_NYU_Open(BtOpenNomal) Then
        Unload Me
    End If
   
    
    If List_Disp_Proc() Then
        Unload Me
    End If
    
    Show
    
    If List1(plstNYU).ListCount > 0 Then
        List1(plstNYU).ListIndex = 0
        List1(plstNYU).SetFocus
    End If
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
    
                                '���ޑO�؂n�o�d�m
    sts = BTRV(BtOpClose, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "���ޑO���ް�")
        End If
    End If
    sts = BTRV(BtOpReset, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set Conv_PI000701 = Nothing

    End
End Sub

Private Sub List1_DblClick(Index As Integer)

Dim W_KEY   As String
Dim sts     As Integer


    W_KEY = Left(List1(Index).List(List1(Index).ListIndex), 20)

    Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
    Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
    Call UniCode_Conv(K0_P_NYU.HIN_GAI, W_KEY)
    Call UniCode_Conv(K0_P_NYU.NYUKA_DT, Format(Mid(List1(Index).List(List1(Index).ListIndex), 24, 10), "YYYYMMDD"))


    sts = BTRV(BtOpGetEqual, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
    Select Case sts
        Case BtNoErr
            
            Text1(ptxHIN_GAI).Text = W_KEY
            
            
            Text1(ptxNYUKA_DT).Text = Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                        Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                        Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 7, 2)
            Text1(ptxNYUKA_QTY).Text = Format(CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)), "#,##0")
        
            Text1(ptxSHIIRE_CODE).Text = Trim(StrConv(P_NYUREC.SHIIRE_CODE, vbUnicode))
            If IsNumeric(StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode)) Then
                Text1(ptxSHIIRE_TANKA).Text = Format(CDbl((StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode))), "#,##0.00")
            Else
                Text1(ptxSHIIRE_TANKA).Text = ""
            End If
                    
        
        
        Case BtErrKeyNotFound
            Text1(ptxHIN_GAI).Text = ""
        
        
            Text1(ptxNYUKA_DT).Text = ""
            Text1(ptxNYUKA_QTY).Text = ""
        
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Text1(ptxHIN_GAI).SetFocus
            Exit Sub
        
        Case Else
            Call File_Error(sts, BtOpGetEqual, "���ޑO���ް�")
            Exit Sub
    End Select

End Sub

Private Sub List1_GotFocus(Index As Integer)
    
    If List1(plstNYU).ListCount > 0 And _
       List1(plstNYU).ListIndex < 0 Then
        List1(plstNYU).ListIndex = 0
    End If

End Sub

Private Sub List1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim W_KEY   As String
Dim sts     As Integer
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If

    If Shift = vbShiftMask Then
        Call Tab_Ctrl(Shift)        '�ړ�
    Else
        
        W_KEY = Left(List1(Index).List(List1(Index).ListIndex), 20)
        
        Call UniCode_Conv(K0_P_NYU.JGYOBU, SHIZAI)
        Call UniCode_Conv(K0_P_NYU.NAIGAI, NAIGAI_NAI)
        Call UniCode_Conv(K0_P_NYU.HIN_GAI, W_KEY)
        Call UniCode_Conv(K0_P_NYU.NYUKA_DT, Format(Mid(List1(Index).List(List1(Index).ListIndex), 24, 10), "YYYYMMDD"))

        sts = BTRV(BtOpGetEqual, P_NYU_POS, P_NYUREC, Len(P_NYUREC), K0_P_NYU, Len(K0_P_NYU), 0)
        Select Case sts
            Case BtNoErr
                Text1(ptxHIN_GAI).Text = W_KEY
                
                Text1(ptxNYUKA_DT).Text = Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & _
                                            Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & _
                                            Mid(StrConv(P_NYUREC.NYUKA_DT, vbUnicode), 7, 2)
                Text1(ptxNYUKA_QTY).Text = Format(CLng(StrConv(P_NYUREC.NYUKA_QTY, vbUnicode)), "#,##0")
            
            
                Text1(ptxSHIIRE_CODE).Text = Trim(StrConv(P_NYUREC.SHIIRE_CODE, vbUnicode))
                If IsNumeric(StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode)) Then
                    Text1(ptxSHIIRE_TANKA).Text = Format(CDbl((StrConv(P_NYUREC.SHIIRE_TANKA, vbUnicode))), "#,##0.00")
                Else
                    Text1(ptxSHIIRE_TANKA).Text = ""
                End If
            
            
            Case BtErrKeyNotFound
                Text1(ptxHIN_GAI).Text = ""
            
                Text1(ptxNYUKA_DT).Text = ""
                Text1(ptxNYUKA_QTY).Text = ""
            
                MsgBox "���͂������ڂ̓G���[�ł��B"
                Text1(ptxHIN_GAI).SetFocus
                Exit Sub
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "���ޑO���ް�")
                Exit Sub
        End Select
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
        
    If Error_Check_Proc(Index) Then     '�G���[�`�F�b�N
        Exit Sub
    End If
        
        
    Call Tab_Ctrl(Shift)        '�ړ�

End Sub

