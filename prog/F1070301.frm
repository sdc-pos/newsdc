VERSION 5.00
Begin VB.Form F1070301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "����X�g���([F107030] 2012.04.19 14:00)"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2250
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   5880
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   0
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�������f"
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
      Left            =   4680
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   8
      Left            =   7800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�ް�"
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   3
      Left            =   2640
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Index           =   0
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "�`"
      Height          =   255
      Index           =   1
      Left            =   5520
      TabIndex        =   17
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "������t�͈�"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   16
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  '�ׯ�
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   6480
      Width           =   180
   End
End
Attribute VB_Name = "F1070301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_DATE% = 0                '�J�n�@���t
Private Const ptxE_DATE% = 1                '�I���@���t

Private Const Text_Max% = 2                 '��ʍ��ڕʍő���ޯ��


Private Print_Jgyobu        As Variant      '����Ώێ��ƕ�
Private Print_Jgyobu_T()    As String * 1


Private Print_Yoin          As Variant      '����Ώۗv��
Private Print_Yoin_T()      As String * 2


Private Const LMAX% = 44                    '�œ��ő�s��
Private Const MGN_L% = 3                    '���]���i�����F�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j

Private Pdate           As String           '����J�n���t�iͯ�ް�p�j
Private Ptime           As String           '����J�n�����iͯ�ް�p�j

Private NormalFont      As New StdFont      '����t�H���g

Private PRT_CAN         As Boolean          '����r���L�����Z���v��

Private F107030CSV      As String           'CSV�o�̓t�@�C��


Private Function Print_Proc() As Integer
    
Dim sts             As Integer
Dim com             As Integer
    
Dim LCNT            As Integer

Dim i               As Integer
Dim j               As Integer


Dim RetBuf          As String

Dim Print_F         As Boolean



    Print_Proc = True


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "����X�g�����", Me.hwnd, 0)


'������́u������f�v�ȊO�̃C�x���g�擾�s��
    Call Input_Lock         '��ʍ��ڃ��b�N
    
    
    
    Set Printer.Font = NormalFont
    Printer.Orientation = vbPRORLandscape   '�p���̒��ӂ���ɂ��Ĉ��
    Pdate = Date
    Ptime = Time
    
    
    
    
    Command1.Visible = True
    Command1.Enabled = True


    Pdate = Date
    Ptime = Time





    PRT_CAN = False

    For i = 0 To UBound(Print_Jgyobu_T)
        Call UniCode_Conv(K0_IDO.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_IDO.JITU_DT, Format(Text1(ptxS_DATE).Text, "YYYYMMDD"))
        Call UniCode_Conv(K0_IDO.JITU_TM, "")
        com = BtOpGetGreaterEqual
        LCNT = 99
    
    
        Do
            DoEvents
                                                '������f�v��
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '��ʍ��ڃ��b�N����
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "����X�g������f", Me.hwnd, 0)
                Command1.Visible = False
                Print_Proc = False
                Exit Function
            End If
    
            
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> Print_Jgyobu_T(i) Then
                        Exit Do
                    End If
                    
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > Format(Text1(ptxE_DATE).Text, "YYYYMMDD") Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�݌Ɉړ���")
                    Exit Function
            End Select
    
    
            Print_F = False
            For j = 0 To UBound(Print_Yoin_T)
                If Print_Yoin_T(j) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                    Print_F = True
                    Exit For
                End If
            
            Next j
    
    
    
            If Print_F Then
    
                '�w�b�_�[�R���g���[��
                If LCNT > LMAX Then
                    Call Print_Head(LCNT)
                End If
        
        
        
                '���ƕ�
                For j = 0 To UBound(JGYOBU_T)
                    If JGYOBU_T(j).CODE = StrConv(IDOREC.JGYOBU, vbUnicode) Then
                        Exit For
                    End If
                Next j
        
        
                Printer.Print Tab(MGN_L);
                If j <= UBound(JGYOBU_T) Then
                    Call Moji_Cut_Proc(JGYOBU_T(j).NAME, RetBuf, 10)
                    Printer.Print RetBuf;
                End If
                Printer.Print Tab(MGN_L + 10);
                Printer.Print Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 7, 2);
        
                Printer.Print Tab(MGN_L + 22);
                Printer.Print Left(StrConv(IDOREC.HIN_GAI, vbUnicode), 14);
                Printer.Print Tab(MGN_L + 37);
                Call Moji_Cut_Proc(StrConv(IDOREC.HIN_NAME, vbUnicode), RetBuf, 35)
                Printer.Print RetBuf;
                
                Printer.Print Tab(MGN_L + 75);
                Printer.Print StrConv(IDOREC.FROM_SOKO, vbUnicode) & "-" & _
                                StrConv(IDOREC.FROM_RETU, vbUnicode) & "-" & _
                                StrConv(IDOREC.FROM_REN, vbUnicode) & "-" & _
                                StrConv(IDOREC.FROM_DAN, vbUnicode);
                RetBuf = Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)), "#,##0")
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print Tab(MGN_L + 90);
                Printer.Print RetBuf;
                
                RetBuf = Format(CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0")
                Printer.Print Tab(MGN_L + 100);
                If Len(RetBuf) < 9 Then
                    RetBuf = Space(9 - Len(RetBuf)) & RetBuf
                End If
                Printer.Print RetBuf;
                
                Printer.Print Tab(MGN_L + 112);
                Call Moji_Cut_Proc(StrConv(IDOREC.MEMO, vbUnicode), RetBuf, 20)
                Printer.Print RetBuf;
                LCNT = LCNT + 1
            End If
            com = BtOpGetNext
    
        Loop
    Next i

    If LCNT <> 99 Then
        Printer.EndDoc
    End If
    
    
    Call Input_UnLock         '��ʍ��ڃ��b�N����
    Command1.Visible = False


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "����X�g����I��", Me.hwnd, 0)

    Print_Proc = False
End Function

Private Sub Print_Head(LCNT As Integer)
                                        
Dim i As Integer
Dim RetBuf As String
Dim sts As Integer

    If LCNT <> 99 Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print
                                        '�w�b�_�[�i�P�j
    Printer.Print Tab(3);
    Printer.Print Tab(36);
    Printer.Print "������  ����X�g  ������";
    Printer.Print Tab(100);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    Printer.Print
                                        '���׈��
    Printer.Print Tab(MGN_L);
    Printer.Print "���ƕ�";
    Printer.Print Tab(MGN_L + 10);
    Printer.Print "�����";
    Printer.Print Tab(MGN_L + 22);
    Printer.Print "�i  ��";
    Printer.Print Tab(MGN_L + 37);
    Printer.Print "�@�@�i             ��";
    Printer.Print Tab(MGN_L + 75);
    Printer.Print "�I    ��";
    Printer.Print Tab(MGN_L + 92);
    Printer.Print "���ς�";
    Printer.Print Tab(MGN_L + 102);
    Printer.Print "�����i";
    Printer.Print Tab(MGN_L + 112);
    Printer.Print "���@��(�w�}�[��)"
    
    Printer.Print

    LCNT = 6 + MGN_U

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1070301.MousePointer = vbHourglass

    Call Ctrl_Lock(F1070301)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1070301)


    F1070301.MousePointer = vbDefault

End Sub
Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> �b�r�u�o��
        Case 7
            If Not IsDate(Text1(ptxS_DATE).Text) Then
                Text1(ptxS_DATE).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Not IsDate(Text1(ptxE_DATE).Text) Then
                Text1(ptxE_DATE).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Text1(ptxS_DATE).Text > Text1(ptxE_DATE).Text Then
                MsgBox "���͂������ڂ̓G���[�ł��B�i���t�͈́j"
                Text1(ptxS_DATE).SetFocus
                Exit Sub
            End If
            
            
            Beep
            ans = MsgBox("�u����X�g�v�f�[�^�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Output_Proc() Then
                    Unload Me
                End If
                Text1(ptxS_DATE).SetFocus
            End If
        
        
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> �b�r�u�o��
        Case 8                              '���
            
            If Not IsDate(Text1(ptxS_DATE).Text) Then
                Text1(ptxS_DATE).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Not IsDate(Text1(ptxE_DATE).Text) Then
                Text1(ptxE_DATE).Text = Format(Now, "YYYY/MM/DD")
            End If
            If Text1(ptxS_DATE).Text > Text1(ptxE_DATE).Text Then
                MsgBox "���͂������ڂ̓G���[�ł��B�i���t�͈́j"
                Text1(ptxS_DATE).SetFocus
                Exit Sub
            End If
            
            
            Beep
            ans = MsgBox("�u����X�g�v������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
                Text1(ptxS_DATE).SetFocus
            End If
                    
        Case 11                             '�I��
            Unload Me
    End Select
End Sub
Private Sub Command1_Click()
    PRT_CAN = True
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
            Command(KeyCode - vbKeyF1).Value = True
    End Select

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
Dim i   As Integer
Dim c   As String * 128
Dim sts As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    

    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "����X�g���", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)



    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = Trim(c)
                                
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
                                
                                
                                '����Ώێ��ƕ�
    If GetIni(App.EXEName, "JGYOBU_CODE", App.EXEName, c) Then
        MsgBox "����Ώێ��ƕ��̊l���Ɏ��s���܂���(JGYOBU_CODE=)�B�����𒆎~���܂��B"
        End
    Else
        Print_Jgyobu = Split(Trim(c), ",", -1)
        Erase Print_Jgyobu_T
        
        For i = 0 To UBound(Print_Jgyobu)
        
            ReDim Preserve Print_Jgyobu_T(0 To i)
            Print_Jgyobu_T(i) = Print_Jgyobu(i)
        Next i
        
        
    End If
                                '����Ώۗv��
    If GetIni(App.EXEName, "YOIN_CODE", App.EXEName, c) Then
        MsgBox "����Ώۗv���̊l���Ɏ��s���܂���(YOIN_CODE=)�B�����𒆎~���܂��B"
        End
    Else
        Print_Yoin = Split(Trim(c), ",", -1)
    
        Erase Print_Yoin_T
        
        For i = 0 To UBound(Print_Yoin)
        
            ReDim Preserve Print_Yoin_T(0 To i)
            Print_Yoin_T(i) = Print_Yoin(i)
        Next i
    
    
    End If
                                
                                '�b�r�u̧��
    If GetIni(App.EXEName, "F107030CSV", App.EXEName, c) Then
    Else
        F107030CSV = Trim(c)
        Command(7).Enabled = True
    End If
                                
                                
                                
                                
                                
                                
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1070301.FontName
        .Size = F1070301.FontSize
    End With
    Set Printer.Font = NormalFont
    
    Text1(ptxS_DATE).Text = Format(Now, "YYYY/MM/DD")
    Text1(ptxE_DATE).Text = Format(Now, "YYYY/MM/DD")
    
    Text1(ptxS_DATE).SetFocus

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
Dim yn  As Integer
    
    
    
    yn = MsgBox("[����X�g���]�������I�����܂����H", vbYesNo, "�m�F����")
    If yn = vbNo Then
        Cancel = True
        Exit Sub
    End If
    
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
    
    sts = BTRV(BtOpReset, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1070301 = Nothing

    End
End Sub
Private Sub Text1_GotFocus(Index As Integer)
    If Text1(Index).TabStop = True Then
        Text1(Index) = Trim(Text1(Index).Text)
        Text1(Index).SelStart = 0
        Text1(Index).SelLength = Len(Text1(Index).Text)
    End If

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i   As Integer

    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    For i = Index + 1 To Text_Max
        If Text1(i).Enabled And Text1(i).Visible And Text1(i).TabStop Then
            Text1(i).SetFocus
            Exit For
        End If
    Next i

End Sub

Private Function Output_Proc() As Integer
    
Dim sts             As Integer
Dim com             As Integer
    
Dim LCNT            As Integer

Dim i               As Integer
Dim j               As Integer


Dim RetBuf          As String

Dim Print_F         As Boolean

Dim FileNo          As Integer


    Output_Proc = True


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "����X�g�f�[�^�o�͒�", Me.hwnd, 0)


'������́u������f�v�ȊO�̃C�x���g�擾�s��
    Call Input_Lock         '��ʍ��ڃ��b�N
    Command1.Visible = True
    Command1.Enabled = True


    Pdate = Date
    Ptime = Time


    FileNo = FreeFile
    
    On Error GoTo Error_Proc
    
    Open (F107030CSV) For Output As FileNo



    LCNT = 99

    PRT_CAN = False

    For i = 0 To UBound(Print_Jgyobu_T)
        Call UniCode_Conv(K0_IDO.JGYOBU, Print_Jgyobu_T(i))
        Call UniCode_Conv(K0_IDO.JITU_DT, Format(Text1(ptxS_DATE).Text, "YYYYMMDD"))
        Call UniCode_Conv(K0_IDO.JITU_TM, "")
        com = BtOpGetGreaterEqual
    
    
        Do
            DoEvents
                                                '������f�v��
            If PRT_CAN Then
                Printer.KillDoc
                Call Input_UnLock               '��ʍ��ڃ��b�N����
                hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
                    "����X�g�f�[�^�o�͒��f", Me.hwnd, 0)
                Command1.Visible = False
                Output_Proc = False
                Exit Function
            End If
    
            
            sts = BTRV(com, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
            Select Case sts
                Case BtNoErr
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> Print_Jgyobu_T(i) Then
                        Exit Do
                    End If
                    
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > Format(Text1(ptxE_DATE).Text, "YYYYMMDD") Then
                        Exit Do
                    End If
                
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�݌Ɉړ���")
                    Exit Function
            End Select
    
    
            Print_F = False
            For j = 0 To UBound(Print_Yoin_T)
                If Print_Yoin_T(j) = StrConv(IDOREC.RIRK_ID, vbUnicode) Then
                    Print_F = True
                    Exit For
                End If
            
            Next j
    
    
    
            If Print_F Then
    
                '�w�b�_�[�R���g���[��
                If LCNT = 99 Then
                    Write #FileNo, "���ƕ�", "�����", "�i��", "�i��", "�I��", "���ς�", "�����i", "�����i�w�}�[���j"
                    LCNT = 0
                End If
                '���ƕ�
                For j = 0 To UBound(JGYOBU_T)
                    If JGYOBU_T(j).CODE = StrConv(IDOREC.JGYOBU, vbUnicode) Then
                        Exit For
                    End If
                Next j
                If j <= UBound(JGYOBU_T) Then
                    Write #FileNo, RTrim(JGYOBU_T(j).NAME),
                End If
                '�����
                Write #FileNo, Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 1, 4) & "/" & _
                                Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 5, 2) & "/" & _
                                Mid(StrConv(IDOREC.JITU_DT, vbUnicode), 7, 2),
        
                '�i��
                Write #FileNo, RTrim(StrConv(IDOREC.HIN_GAI, vbUnicode)),
                '�i��
                Write #FileNo, RTrim(StrConv(IDOREC.HIN_NAME, vbUnicode)),
                '�I��
                Write #FileNo, StrConv(IDOREC.FROM_SOKO, vbUnicode) & "-" & _
                                StrConv(IDOREC.FROM_RETU, vbUnicode) & "-" & _
                                StrConv(IDOREC.FROM_REN, vbUnicode) & "-" & _
                                StrConv(IDOREC.FROM_DAN, vbUnicode),
                '���i���ς�
                Write #FileNo, Format(CLng(StrConv(IDOREC.SUMI_JITU_QTY, vbUnicode)), "#,##0"),
                '�����i
                Write #FileNo, Format(CLng(StrConv(IDOREC.MI_JITU_QTY, vbUnicode)), "#,##0"),
                '�����i�w�}�[���j
                Write #FileNo, RTrim(StrConv(IDOREC.MEMO, vbUnicode)),
                                    
                Write #FileNo,
            
            End If
            com = BtOpGetNext
    
        Loop
    Next i

    
    
    Close #FileNo
    
    MsgBox "�u" & F107030CSV & "�v�͐���ɏo�͂���܂����B"
    
    
    
    
    Call Input_UnLock         '��ʍ��ڃ��b�N����
    Command1.Visible = False

    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "����X�g�f�[�^�o�͏I��", Me.hwnd, 0)

    Output_Proc = False
    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox F107030CSV & "���g�p���ł��B"
        Call Input_UnLock         '��ʍ��ڃ��b�N����
        Output_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
    End If

End Function


