VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form PI999981 
   Caption         =   "���i���w�}�[�ꊇ���s"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10185
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
   ScaleHeight     =   6270
   ScaleWidth      =   10185
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   13
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   1
      Left            =   7560
      TabIndex        =   12
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '�E����
      Height          =   375
      Index           =   0
      Left            =   2940
      TabIndex        =   9
      Top             =   5160
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   3900
      Left            =   4830
      TabIndex        =   6
      Top             =   1200
      Width           =   3585
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Index           =   0
      Left            =   1680
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   5
      Top             =   240
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Ǎ���"
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
      Left            =   630
      TabIndex        =   4
      Top             =   720
      Width           =   1380
   End
   Begin VB.ListBox List1 
      Height          =   3900
      Left            =   630
      TabIndex        =   3
      Top             =   1200
      Width           =   3165
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�I��"
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
      Left            =   7455
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
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
      Left            =   6090
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "�m�f����"
      Height          =   255
      Index           =   4
      Left            =   6615
      TabIndex        =   11
      Top             =   5760
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "�n�j����"
      Height          =   255
      Index           =   3
      Left            =   6615
      TabIndex        =   10
      Top             =   5280
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "�Ǎ��݌���"
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   5280
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "�������"
      Height          =   255
      Index           =   1
      Left            =   4830
      TabIndex        =   7
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "�d������"
      Height          =   255
      Index           =   0
      Left            =   525
      TabIndex        =   2
      Top             =   360
      Width           =   960
   End
End
Attribute VB_Name = "PI999981"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�e�L�X�g�p�Y��
Private Const ptxJGYOBU% = 0            '���ƕ�
Private Const ptxNAIGAI% = 1            '�����O


Private Const ptxS_YMD% = 2             '�J�n�@���t�͈�
Private Const ptxE_YMD% = 3             '�I���@���t�͈�

Private Const ptxCOUNT% = 4             '�Ώی�

Private Const ptxSEL_CLASS% = 5         '�I���@�׽
Private Const ptxSEL_BOX% = 6           '�I���@�׽


Private Const ptxS_SOKO_No% = 7         '�J�n�@�q�ɇ�
Private Const ptxS_Retu% = 8            '�J�n�@��
Private Const ptxS_Ren% = 9             '�J�n�@�A
Private Const ptxS_Dan% = 10            '�J�n�@�i

Private Const ptxe_SOKO_No% = 11        '�I���@�q�ɇ�
Private Const ptxe_Retu% = 12           '�I���@��
Private Const ptxe_Ren% = 13            '�I���@�A
Private Const ptxe_Dan% = 14            '�I���@�i

Private Const pcmbSHIMUKE% = 0

Private IN_cnt  As Integer
Private OK_cnt  As Integer
Private NG_cnt  As Integer


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    PI999981.MousePointer = vbHourglass

    Call Ctrl_Lock(PI999981)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(PI999981)


    PI999981.MousePointer = vbDefault

End Sub

Private Function Error_Check_Proc(Mode As Integer, chk As Integer, flg) As Integer
'----------------------------------------------------------------------------
'                   ���͍��ڂ̃G���[�`�F�b�N
'----------------------------------------------------------------------------
    
Dim sts         As Integer
Dim yn          As Integer
    
    
    Error_Check_Proc = True
    
        
        
    Error_Check_Proc = False
    

End Function







Private Sub Command1_Click(Index As Integer)

Dim ans             As Integer
Dim i               As Integer

Dim f               As New PI999982

Dim rpt2            As New PI99998F2


Dim com             As Integer
Dim sts             As Integer

Dim FileNo          As Long
Dim wkText          As String

Dim Skip_F          As Boolean


    Select Case Index
        Case 0              '���
            
            
            
            Beep
            ans = MsgBox("������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                
                List2.Clear
                
                OK_cnt = 0
                NG_cnt = 0
                Text1(1).Text = Format(OK_cnt, "#,##0")
                Text1(2).Text = Format(NG_cnt, "#,##0")
                
                For i = 0 To List1.ListCount - 1
            
            
            
                    Taget_SHIMUKE_CODE_KEY = Mid(Right(Combo1(pcmbSHIMUKE), 4), 1, 2)
                    Taget_JGYOBU_key = Mid(Right(Combo1(pcmbSHIMUKE), 4), 3, 1)
                    Taget_NAIGAI_key = Mid(Right(Combo1(pcmbSHIMUKE), 4), 4, 1)
                    Taget_Hin_key = Trim(Left(List1.List(i), 20))
                    
                    Skip_F = False
                    
                    
                    Call UniCode_Conv(K0_ITEM.JGYOBU, Taget_JGYOBU_key)
                    Call UniCode_Conv(K0_ITEM.NAIGAI, Taget_NAIGAI_key)
                    Call UniCode_Conv(K0_ITEM.HIN_GAI, Taget_Hin_key)
                    
                    
                    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            
                            Skip_F = True
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                            Exit Sub
                    
                    End Select
                    
                    
                    
                    Call UniCode_Conv(K0_P_COMPO.SHIMUKE_CODE, Taget_SHIMUKE_CODE_KEY)
                    Call UniCode_Conv(K0_P_COMPO.JGYOBU, Taget_JGYOBU_key)
                    Call UniCode_Conv(K0_P_COMPO.NAIGAI, Taget_NAIGAI_key)
                    Call UniCode_Conv(K0_P_COMPO.HIN_GAI, Taget_Hin_key)
                    
                    Call UniCode_Conv(K0_P_COMPO.DATA_KBN, "0")
                    Call UniCode_Conv(K0_P_COMPO.SEQNO, "000")
                
                    
                    
                    sts = BTRV(BtOpGetEqual, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
                        
                    Select Case sts
                        Case BtNoErr
                        Case BtErrKeyNotFound
                            
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "�\���}�X�^")
                            Exit Sub
                    
                    End Select
                    
                    
                    
                    
                    If Skip_F Then
                        List2.AddItem Taget_Hin_key & " " & "NG"
                        NG_cnt = NG_cnt + 1
                        Text1(2).Text = Format(NG_cnt, "#,##0")
                    Else
                    
                        Set rpt2 = New PI99998F2
                        '���|�[�g��������܂��B�itrue�F����_�C�A���O���� false�F�Ȃ��j
                        rpt2.PrintReport False
                        Set rpt2 = Nothing
                    
                        List2.AddItem Taget_Hin_key & " " & "OK"
                        OK_cnt = OK_cnt + 1
                        Text1(1).Text = Format(OK_cnt, "#,##0")
                    
                    
                    End If
                Next i
            
                MsgBox "������I�����܂����B"
            
            
            End If
        Case 1              '�I��
            Unload Me
    
        Case 2
    
            List1.Clear
            
            CommonDialog1.Filter = "���ׂẴt�@�C�� (*.*)|*.*|"
            CommonDialog1.FilterIndex = 2
        
            On Error GoTo ErrHandler
        
            CommonDialog1.ShowOpen
    
            FileNo = FreeFile
            Open CommonDialog1.fileName For Input As #FileNo
    
            IN_cnt = 0
            Text1(0).Text = Format(IN_cnt, "#,##0")
    
            Do Until eof(FileNo)
                Line Input #FileNo, wkText
                If Trim(wkText) = "" Then
                    Exit Do
                End If
    
                List1.AddItem Trim(wkText)
                IN_cnt = IN_cnt + 1
    
                Text1(0).Text = Format(IN_cnt, "#,##0")
    
            Loop
    
            Close #FileNo
    
    End Select

ErrHandler:
    
End Sub


Private Sub Form_DblClick()
    PrintForm
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()

Dim c           As String * 128
Dim sts         As Integer
Dim i           As Integer

Dim MUKE_CODE   As Variant


    If App.PrevInstance Then
        MsgBox "����v���O�������s���ł��B"
        End
    End If

                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���܂��B"
        End
    End If
    LOG_F = RTrim(c)
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�R�[�h�}�X�^�n�o�d�m
    If P_CODE_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�\���}�X�^�n�o�d�m
    If P_COMPO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�S���҃}�X�^�n�o�d�m
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌��ް��n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    '����Ͻ���`
    Call P_CODE_TBL_Proc
    
    
    
    Load PI999982
    
    
    
    
    
    '�d������̃Z�b�g
    If Code_Set_Proc(pcmbSHIMUKE, P_KBN04_CD, 0) Then
        Unload Me
    End If
    
    
    
    Doukon_Tbl_No(0) = "�@"
    Doukon_Tbl_No(1) = "�A"
    Doukon_Tbl_No(2) = "�B"
    Doukon_Tbl_No(3) = "�C"
    Doukon_Tbl_No(4) = "�D"
    Doukon_Tbl_No(5) = "�E"
    Doukon_Tbl_No(6) = "�F"
    Doukon_Tbl_No(7) = "�G"
    Doukon_Tbl_No(8) = "�H"
    Doukon_Tbl_No(9) = "�I"
    Doukon_Tbl_No(10) = "�J"
    Doukon_Tbl_No(11) = "�K"
    Doukon_Tbl_No(12) = "�L"
    Doukon_Tbl_No(13) = "�M"
    Doukon_Tbl_No(14) = "�N"
    Doukon_Tbl_No(15) = "�O"
    Doukon_Tbl_No(16) = "�P"
    Doukon_Tbl_No(17) = "�Q"
    Doukon_Tbl_No(18) = "�R"
    Doukon_Tbl_No(19) = "�S"
    
    
    
End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
                                            
                                            
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
    
                                            '�R�[�h�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�R�[�h�}�X�^")
        End If
    End If
    
                                            '�\���}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, P_COMPO_POS, P_COMPO_O_REC, Len(P_COMPO_O_REC), K0_P_COMPO, Len(K0_P_COMPO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�\���}�X�^")
        End If
    End If
    
                                            '�S���҃}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�S���҃}�X�^")
        End If
    End If
    
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    Set PI999981 = Nothing
    Set PI999982 = Nothing

    End
End Sub


Private Function Code_Set_Proc(Index As Integer, KBN As String, Mode As Integer) As Integer
'----------------------------------------------------------------------------
'                   �R�[�h�}�X�^���R���{�ɃZ�b�g����B
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer
Dim Key_Len     As Integer
Dim OPTION1     As Integer
Dim OPTION2     As Integer

Dim wkOption    As String



Dim i           As Integer
    
    Code_Set_Proc = True
    
    Combo1(Index).Clear
    
    For i = 0 To UBound(P_KBN_TBL)
    
        If KBN = P_KBN_TBL(i).KBN_CD Then
            Key_Len = P_KBN_TBL(i).KBN_Len
            Exit For
        End If
    
    Next i
    
    If i > UBound(P_KBN_TBL) Then
        Exit Function
    End If
    
    If Mode = 1 Then
        Combo1(Index).AddItem Space(Key_Len)
    End If
    
    Call UniCode_Conv(K0_P_CODE.DATA_KBN, KBN)
    Call UniCode_Conv(K0_P_CODE.C_Code, "")

    com = BtOpGetGreater

    Do
        DoEvents
    
        sts = BTRV(com, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
            
        Select Case sts
            Case BtNoErr
            
                                
                If StrConv(P_CODEREC.DATA_KBN, vbUnicode) <> KBN Then
                    
                    Exit Do
                
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�R�[�h�}�X�^")
                Exit Function
        
        End Select

        wkOption = ""
        If P_KBN_TBL(i).KBN_OP1 Then
            wkOption = Trim(StrConv(P_CODEREC.OPTION1, vbUnicode))
        End If
        If P_KBN_TBL(i).KBN_OP2 Then
            wkOption = wkOption & Trim(StrConv(P_CODEREC.OPTION2, vbUnicode))
        End If
        
        
        
        Combo1(Index).AddItem StrConv(P_CODEREC.C_RNAME, vbUnicode) & "            " & _
                                Left(StrConv(P_CODEREC.C_Code, vbUnicode), Key_Len) & wkOption
        
        
        com = BtOpGetNext
    
    Loop

    Code_Set_Proc = False
    



End Function


