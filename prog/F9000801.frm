VERSION 5.00
Begin VB.Form F9000801 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�ړ����݌ɏW�v����"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2550
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
      Index           =   5
      Left            =   7770
      MaxLength       =   2
      TabIndex        =   27
      Top             =   1680
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   7035
      MaxLength       =   2
      TabIndex        =   25
      Top             =   1680
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   5985
      MaxLength       =   4
      TabIndex        =   23
      Top             =   1680
      Width           =   645
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   21
      Top             =   1680
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   4305
      MaxLength       =   2
      TabIndex        =   19
      Top             =   1680
      Width           =   330
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   3255
      MaxLength       =   4
      TabIndex        =   17
      Top             =   1680
      Width           =   645
   End
   Begin VB.CommandButton Command 
      Caption         =   "�I  ��"
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      Index           =   8
      Left            =   7800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�ް�"
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
      TabIndex        =   7
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
      Index           =   5
      Left            =   4800
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
      Index           =   4
      Left            =   3960
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
      Index           =   3
      Left            =   2640
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
      Index           =   2
      Left            =   1800
      TabIndex        =   2
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
      TabIndex        =   1
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
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   5
      Left            =   8190
      TabIndex        =   28
      Top             =   1800
      Width           =   225
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   4
      Left            =   7455
      TabIndex        =   26
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�N"
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   24
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���`"
      Height          =   255
      Index           =   2
      Left            =   5460
      TabIndex        =   22
      Top             =   1800
      Width           =   540
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   1
      Left            =   4725
      TabIndex        =   20
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�N"
      Height          =   255
      Index           =   0
      Left            =   3990
      TabIndex        =   18
      Top             =   1800
      Width           =   330
   End
   Begin VB.Label lbloUT_Cnt 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6645
      TabIndex        =   16
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�o�͌�����"
      Height          =   255
      Index           =   0
      Left            =   5445
      TabIndex        =   15
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblIn_Cnt 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4545
      TabIndex        =   14
      Top             =   2520
      Width           =   915
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���͌�����"
      Height          =   255
      Index           =   5
      Left            =   3345
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
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
      TabIndex        =   12
      Top             =   6480
      Width           =   180
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F9000801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim SYUKA_DATA      As String               '�z�X�g�I�Ԑݒ�f�[�^�t���p�X
Dim SOKO_CODE       As String * 3           '�q�ɺ���
Dim SEQ_NO          As Double               '�A��

Private Function OUTPUT_Proc() As Integer

Dim sts                     As Integer
Dim FileNo                  As Long
Dim fileName                As String       '̧�ٖ�

Dim In_Cnt                  As Long
Dim Out_Cnt                 As Long

Dim com                     As Integer

Dim wkData_Kbn              As String * 1




    OUTPUT_Proc = True
'���s���̓C�x���g�擾�s��
    Call Input_Lock         '��ʍ��ڃ��b�N

    FileNo = FreeFile
    fileName = SYUKA_DATA


    
    DoEvents

    On Error GoTo Error_Proc
    Open (fileName) For Output As FileNo
                    
    In_Cnt = 0
    Out_Cnt = 0

    com = BtOpGetFirst
    Do
        DoEvents
        
        Do
        
            sts = BTRV(com + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    DoEvents
                Case Else
                    Call File_Error(sts, com, "�o�ח\��f�[�^")
                    Exit Do
            End Select
        
        Loop
        
        If sts <> BtNoErr Then
            Exit Do
        End If
        
        In_Cnt = In_Cnt + 1
        lblIn_Cnt = Format(In_Cnt, "#0")
        
        
        
        If (StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "25" Or StrConv(Y_SYUREC.TORI_KBN, vbUnicode) = "29") Then
'''            If StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode) <= Format(Now, "YYYYMMDD") Then    2006.10.13 ���t�������O��
                If StrConv(Y_SYUREC.CYU_KBN, vbUnicode) <> CYU_KBN_BOU Then
                    
                    
                    If StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "1" Or StrConv(Y_SYUREC.DATA_KBN, vbUnicode) = "3" Then
                    
                        If Trim(StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode)) <> "" Then
                        
                        
                            If Trim(StrConv(Y_SYUREC.LK_SEQ_NO, vbUnicode)) = "" Then
        
        
        
        
                                Out_Cnt = Out_Cnt + 1
                                lbloUT_Cnt = Format(Out_Cnt, "#0")
                
                
                
                                If CDbl(Right(Format(SEQ_NO, "000000000000"), 11)) = 99999999999# Then
                                    SEQ_NO = CDbl(Left(Format(SEQ_NO, "000000000000"), 1) & "00000000000")
                                End If
                                SEQ_NO = SEQ_NO + 1
                                Print #FileNo, Format(SEQ_NO, "000000000000");                  '�ް����ݽ�ԍ�
                                Print #FileNo, StrConv(Y_SYUREC.JGYOBA, vbUnicode);             '���Ə꺰��
                                
                                wkData_Kbn = StrConv(Y_SYUREC.DATA_KBN, vbUnicode)
                                If wkData_Kbn = "3" Then
                                    wkData_Kbn = "2"
                                End If
                                Print #FileNo, wkData_Kbn;                                      '�ް��敪
                                
                                Print #FileNo, StrConv(Y_SYUREC.TORI_KBN, vbUnicode);           '����敪
                                Print #FileNo, StrConv(Y_SYUREC.ID_NO, vbUnicode);              'ID-NO
                                Print #FileNo, StrConv(Y_SYUREC.KAIKEI_JGYOBA, vbUnicode);      '��v�p���Ə꺰��
                                Print #FileNo, StrConv(Y_SYUREC.SHISAN_JGYOBA, vbUnicode);      '���Y�Ǘ����Ə꺰��
                                Print #FileNo, StrConv(Y_SYUREC.HIN_NO, vbUnicode);             '�i�ڔԍ�
                                Print #FileNo, StrConv(Y_SYUREC.DEN_NO, vbUnicode);             '�`�[�ԍ�
                                Print #FileNo, Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), _
                                                "000000000.0000");                              '�o�Ɏ��ѐ�
                                Print #FileNo, StrConv(Y_SYUREC.LK_MUKE_CODE, vbUnicode);       '����溰��
                                Print #FileNo, StrConv(Y_SYUREC.SYUKO_SYUSI, vbUnicode);        '�݌Ɏ��x
                                Print #FileNo, StrConv(Y_SYUREC.SHISAN_SYUSI, vbUnicode);       '���Y�Ǘ��݌Ɏ��x
                                Print #FileNo, StrConv(Y_SYUREC.HOJYO_SYUSI, vbUnicode);        '�⏕�݌Ɏ��x
                                Print #FileNo, StrConv(Y_SYUREC.KEY_SYUKA_YMD, vbUnicode);      '�o�׎��ѓ��t
                                Print #FileNo, StrConv(Y_SYUREC.KENPIN_TANTO_CODE, vbUnicode) & _
                                                Space(3);                                       '�S���ҁi���i�S���ҁj
                                Print #FileNo, String(12, "0");                                 '����ԍ�
                                Print #FileNo, Space(20);                                       '�S���Җ�
                                Print #FileNo, StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode);         '�����J�n���t�i���i���j
                                Print #FileNo, Left(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 4) & _
                                                "0000";                                         '�����J�n�����i���i�����j
                                Print #FileNo, StrConv(Y_SYUREC.KENPIN_YMD, vbUnicode);         'HT�������t�i���i���j
                                Print #FileNo, Left(StrConv(Y_SYUREC.KENPIN_HMS, vbUnicode), 4) & _
                                                "0000";                                         'HT���������i���i�����j
                                Print #FileNo, SOKO_CODE                                        '�q�ɺ���
                            
                                Call UniCode_Conv(Y_SYUREC.LK_SEQ_NO, Format(SEQ_NO, "000000000000"))
                            
                                Do
                                
                                    sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                                
                                    Select Case sts
                                        Case BtNoErr
                                            Exit Do
                                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                                            DoEvents
                                        Case Else
                                            Call File_Error(sts, com, "�o�ח\��f�[�^")
                                            Exit Do
                                    End Select
                                
                                Loop
                            
                            
                                If sts <> BtNoErr Then
                                    Exit Do
                                End If
                            End If
                        End If
                    End If
                End If
'''            End If
        End If
    
    
    
        com = BtOpGetNext
    
    Loop


    sts = WriteIni(App.EXEName, "SEQ_NO", App.EXEName, Format(SEQ_NO, "000000000000"))

    Close #FileNo

    Call Input_UnLock         '��ʍ��ڃ��b�N����

    OUTPUT_Proc = False


    Exit Function

Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "���g�p���ł��B"
        Call Input_UnLock         '��ʍ��ڃ��b�N����
        OUTPUT_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        OUTPUT_Proc = True
    End If
End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1200901.MousePointer = vbHourglass

    Call Ctrl_Lock(F1200901)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1200901)


    F1200901.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Dim i       As Integer
Dim c       As String
Dim sts     As Integer

    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If


    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̎擾�Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = Trim(c)
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����n�o�d�m
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer

                                            '�o�ח\���ް��b�k�n�r�d
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�o�ח\��}�X�^")
        End If
    End If

    sts = BTRV(BtOpReset, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If

    Set F1200901 = Nothing

    End
End Sub




