VERSION 5.00
Begin VB.Form F1040151 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�݌ɖ⍇�킹�i�I�ԕʁj"
   ClientHeight    =   13605
   ClientLeft      =   2025
   ClientTop       =   2550
   ClientWidth     =   11475
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
   ScaleHeight     =   13605
   ScaleWidth      =   11475
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.PictureBox Picture1 
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   30
      Top             =   3000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʈ��"
      Height          =   495
      Left            =   9000
      TabIndex        =   29
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   11100
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   9780
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   720
      MaxLength       =   2
      TabIndex        =   0
      Top             =   120
      Width           =   375
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
      Left            =   10260
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   12420
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
      Left            =   9420
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   12420
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
      Left            =   8580
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   12420
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
      Left            =   7740
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   12420
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�� �V"
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
      Left            =   6420
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   12420
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
      Left            =   5580
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   12420
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
      Left            =   4740
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   12420
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
      Left            =   3900
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   12420
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
      Left            =   2580
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   12420
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
      Left            =   1740
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   12420
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
      Left            =   900
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   12420
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
      Left            =   60
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   12420
      Width           =   855
   End
   Begin VB.Label lblDateTime 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8640
      TabIndex        =   31
      Top             =   12960
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���i��"
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   28
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�W���I��"
      Height          =   255
      Index           =   5
      Left            =   9060
      TabIndex        =   27
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��(*)"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   26
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�݌ɐ�"
      Height          =   255
      Index           =   3
      Left            =   7590
      TabIndex        =   25
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԁi���j"
      Height          =   255
      Index           =   2
      Left            =   5595
      TabIndex        =   24
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ד�"
      Height          =   255
      Index           =   1
      Left            =   4335
      TabIndex        =   23
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԁi�O�j"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   22
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   21
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   20
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   19
      Top             =   240
      Width           =   135
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
      Left            =   180
      TabIndex        =   18
      Top             =   12840
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�I��"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   495
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1040151"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WS_NO As String


Private Const ptxSoko_No% = 0           '�q�ɇ�
Private Const ptxRetu% = 1              '��
Private Const ptxRen% = 2               '�A
Private Const ptxDan% = 3               '�i
    
Private Const pLstZaiko% = 0            '�݌�ؽ�
    
Private Const Text_Max% = 3


'Private Const Last_Update_Day$ = "[F104015]2016.01.26 08:30"
Private Const Last_Update_Day$ = "[F104015]2018.10.02 13:30"

Private Function List_Dsp() As Integer

Dim sts         As Integer
Dim com         As Integer
Dim i           As Integer
Dim RetBuf      As String
Dim Edit        As String
    
    
    List_Dsp = True
    
    Call Input_Lock
    
    List1.Clear
    
    Call UniCode_Conv(K0_ZAIKO.Soko_No, Text(ptxSoko_No).Text)
    Call UniCode_Conv(K0_ZAIKO.Retu, Text(ptxRetu).Text)
    Call UniCode_Conv(K0_ZAIKO.Ren, Text(ptxRen).Text)
    Call UniCode_Conv(K0_ZAIKO.Dan, Text(ptxDan).Text)
    Call UniCode_Conv(K0_ZAIKO.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ZAIKO.NAIGAI, "")
    Call UniCode_Conv(K0_ZAIKO.HIN_GAI, "")
    Call UniCode_Conv(K0_ZAIKO.GOODS_ON, "")
    Call UniCode_Conv(K0_ZAIKO.NYUKA_DT, "")
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
        Select Case sts
            Case BtNoErr
                                                        '�I�ԃu���[�N
                If StrConv(ZAIKOREC.Soko_No, vbUnicode) <> Text(ptxSoko_No).Text Or _
                    StrConv(ZAIKOREC.Retu, vbUnicode) <> Text(ptxRetu).Text Or _
                    StrConv(ZAIKOREC.Ren, vbUnicode) <> Text(ptxRen).Text Or _
                    StrConv(ZAIKOREC.Dan, vbUnicode) <> Text(ptxDan).Text Then
                    Exit Do
                End If
                                                        '���ƕ��u���[�N
                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If
                        
                If StrConv(ZAIKOREC.GOODS_ON, vbUnicode) = GOODS_OFF Then
                    Edit = "  "
                Else
                    Edit = "* "
                End If
                If StrConv(ZAIKOREC.NAIGAI, vbUnicode) = NAIGAI_GAI Then
                    Edit = Edit & "�O" & "   "
                Else
                    Edit = Edit & "  " & "   "
                End If
                
                Edit = Edit & StrConv(ZAIKOREC.HIN_GAI, vbUnicode) & " "
                Edit = Edit & Mid(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 1, 4) & "/" & Mid$(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 5, 2) & "/" & Mid$(StrConv(ZAIKOREC.NYUKA_DT, vbUnicode), 7, 2) & " "
                
                RetBuf = Replace(StrConv(ZAIKOREC.HIN_NAI, vbUnicode), Chr(0), " ")
                
                Edit = Edit & Left(RetBuf, 13) & " "
                
                
                RetBuf = Format(CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode)), "#0")
                If Len(Trim(RetBuf)) < 8 Then
                    RetBuf = Space(8 - Len(Trim(RetBuf))) & Trim(RetBuf)
                End If
                Edit = Edit & RetBuf & " "
                
                Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ZAIKOREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ZAIKOREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ZAIKOREC.HIN_GAI, vbUnicode))
                sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                Select Case sts
                    Case BtNoErr
                        Edit = Edit & "    " & StrConv(ITEMREC.ST_SOKO, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                                                            StrConv(ITEMREC.ST_DAN, vbUnicode)
                    Case BtErrKeyNotFound


                    Case Else
                        Call Input_UnLock
                        Call File_Error(sts, BtOpGetEqual, "�i�ڃ}�X�^")
                        List_Dsp = True
                        Exit Function
                End Select
                
                List1.AddItem Edit
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call Input_UnLock
                Call File_Error(sts, com, "�݌Ƀf�[�^")
                List_Dsp = True
                Exit Function
        End Select
        
        
        com = BtOpGetNext
    
    Loop


    lblDateTime.Caption = Format(Now, "yyyy/mm/dd HH:MM")       '2018.10.02

    Call Input_UnLock
    
    Text(ptxSoko_No).SetFocus
    
    List_Dsp = False
    
End Function
                                    '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field(Mode As Integer)
Dim i  As Integer

    For i = Mode To Text_Max
       Text(i).Text = ""
    Next i
    
    List1.Clear
End Sub

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1040151.MousePointer = vbHourglass

    Call Ctrl_Lock(F1040151)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1040151)


    F1040151.MousePointer = vbDefault

End Sub

Private Sub Command_Click(Index As Integer)
    
    Select Case Index
        Case 7                              '�ŐV�\��
            If List_Dsp() Then
                Unload Me
            End If
                        
        Case 11                             '�I��
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub


Private Sub Command1_Click()
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�݌ɖ⍇�킹�i�I�ԕʁj ��ʈ�����J�n���܂��� ", Me.hwnd, 0)


Call Form_HCopy_Win7_NEW(Picture1, vbPRPSA4, vbPRORLandscape)       '2018.10.02


    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�݌ɖ⍇�킹�i�I�ԕʁj ��ʈ�����I�����܂��� ", Me.hwnd, 0)

End Sub

Private Sub Form_DblClick()
'2018.10.02    PrintForm
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
Dim i As Integer
Dim c As String * 128
Dim sts As Integer
    
    If App.PrevInstance Then
        Beep
        MsgBox "����v���O�������s���ł��B"
        End
    End If
    
    Show
                                
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�݌ɖ⍇�킹�i�I�ԕʁj", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
                                
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = RTrim(c)
                                

                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1040151.Caption = "�݌ɖ⍇�킹�i�I�ԕʁj�i" + RTrim(JGYOBU_T(i).NAME) + ") " & Last_Update_Day
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)

                                
                                '�i�ڃ}�X�^OPEN
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^OPEN
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '��ʏ����ݒ�
    Call Clear_Field(ptxSoko_No)
    
    Text(ptxSoko_No).SetFocus
    
    End Sub

Private Sub Form_Unload(CANCEL As Integer)
Dim sts As Integer
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
    
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1040151 = Nothing

    End
End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1040151.Caption = "�݌ɖ⍇�킹�i�I�ԕʁj�i" + RTrim(JGYOBU_T(Index).NAME) + ") " & Last_Update_Day
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

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
    
    If KeyCode <> vbKeyReturn Then Exit Sub

    Select Case Index
        Case ptxSoko_No                 '�q�ɇ�
            Text(Index).Text = StrConv(Text(Index).Text, vbUpperCase)       '2016.01.25
        Case ptxRetu, ptxRen, ptxDan    '��^�A�^�i
            If Not IsNumeric(Text(Index).Text) Then
                Beep
                MsgBox "���͂������ڂ̓G���[�ł��B���l�œ��͂��ĉ�����"
                Text(Index).SetFocus
                Exit Sub
            Else
                Text(Index).Text = StrConv(Text(Index).Text, vbUpperCase)       '2016.01.25
                Text(Index).Text = Format(CInt(Text(Index).Text), "00")
            End If
                        
            If Index = ptxDan Then
                If List_Dsp() Then
                    Unload Me
                End If
            End If
    End Select

    For i = Index + 1 To Text_Max
        If Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i

End Sub


Private Sub Text_LostFocus(Index As Integer)

    Text(Index).Text = StrConv(Text(Index).Text, vbUpperCase)       '2016.01.25


End Sub

