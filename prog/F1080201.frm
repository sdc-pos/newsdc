VERSION 5.00
Begin VB.Form F1080201 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�����ʒI���X�g���"
   ClientHeight    =   6960
   ClientLeft      =   2025
   ClientTop       =   2265
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
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   5280
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1560
      Width           =   615
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
      Index           =   10
      Left            =   9480
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
      Index           =   9
      Left            =   8640
      TabIndex        =   10
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�f�[�^"
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
      Index           =   6
      Left            =   5640
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
      Index           =   5
      Left            =   4800
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
      Index           =   4
      Left            =   3960
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
      Index           =   3
      Left            =   2640
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
      Index           =   2
      Left            =   1800
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
      Index           =   1
      Left            =   960
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
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
      TabIndex        =   14
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������"
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   13
      Top             =   1680
      Width           =   1095
   End
End
Attribute VB_Name = "F1080201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxPACKING_NO% = 0            '�J�n�@�q��

Private Const Text_Max% = 0                 '��ʍ��ڕʍő���ޯ��

Private Const LMAX% = 65                    '�œ��ő�s��
Private Const LCTL% = 99                    '

Private Const MGN_L% = 5                    '���]���i�����F�P����j
Private Const MGN_U% = 1                    '��]���i�s���F�P����j

Private Const Page_Soko_cnt% = 10           '�œ��q�ɐ�

Private Pdate           As String           '����J�n���t�iͯ�ް�p�j
Private Ptime           As String           '����J�n�����iͯ�ް�p�j

Private PTANA_DATA      As String           'CSV�f�[�^�t���p�X

Private NormalFont      As New StdFont      '����t�H���g


Private Type Soko_tbl_Tag
    Soko_No             As String * 2
'    Page_cnt            As Integer
End Type

Dim Soko_Tbl()          As Soko_tbl_Tag     '�q�ɏ��e�[�u��
     

Private Type RetuRen_Tag
    Retu                As String * 2
    Ren                 As String * 2
End Type

Private Type Retu_tag
    RetuRen()           As RetuRen_Tag
End Type

Private Location()      As Retu_tag

'Private Max_Gyo         As Integer
Private Const Retu_Max% = 10



Private Function Print_Proc() As Integer
'----------------------------------------------------------------------------
'                  �u�����ʒI���X�g�v�������
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
    
Dim LCNT            As Integer

Dim Save_Packing_No As String * 4
Dim Save_Rank       As String * 3
Dim Pri_Rank        As String * 3
Dim Save_Page       As String * 1

Dim Tab_Pos         As Integer

Dim Print_tab       As Integer


    Print_Proc = True

    Call Input_Lock         '��ʍ��ڃ��b�N
    
    If Data_Make_Proc() Then
        Exit Function
    End If
    
    Printer.Orientation = vbPRORPortrait   '�p���̒��ӂ���ɂ��Ĉ��
    Pdate = Date
    Ptime = Time
    
    
    LCNT = LCTL

    com = BtOpGetFirst

    Do
        DoEvents
        sts = BTRV(com, PTANA_POS, PTANAREC, Len(PTANAREC), K0_PTANA, Len(K0_PTANA), 0)
        Select Case sts
            Case BtNoErr
                
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�����ʒI���X�g����t�@�C��")
                Exit Function
        End Select
    
    
        If Not IsNumeric(Left(StrConv(PTANAREC.RETUREN01, vbUnicode), 2)) And _
            Not IsNumeric(Left(StrConv(PTANAREC.RETUREN02, vbUnicode), 2)) And _
            Not IsNumeric(Left(StrConv(PTANAREC.RETUREN03, vbUnicode), 2)) And _
            Not IsNumeric(Left(StrConv(PTANAREC.RETUREN04, vbUnicode), 2)) And _
            Not IsNumeric(Left(StrConv(PTANAREC.RETUREN05, vbUnicode), 2)) And _
            Not IsNumeric(Left(StrConv(PTANAREC.RETUREN06, vbUnicode), 2)) And _
            Not IsNumeric(Left(StrConv(PTANAREC.RETUREN07, vbUnicode), 2)) And _
            Not IsNumeric(Left(StrConv(PTANAREC.RETUREN08, vbUnicode), 2)) And _
            Not IsNumeric(Left(StrConv(PTANAREC.RETUREN09, vbUnicode), 2)) And _
            Not IsNumeric(Left(StrConv(PTANAREC.RETUREN10, vbUnicode), 2)) Then
        Else
    
            If com = BtOpGetFirst Then
                Save_Packing_No = StrConv(PTANAREC.Packing_No, vbUnicode)
                Save_Rank = StrConv(PTANAREC.Rank, vbUnicode)
                Save_Page = StrConv(PTANAREC.Page_cnt, vbUnicode)
            End If
        
            If Save_Packing_No <> StrConv(PTANAREC.Packing_No, vbUnicode) Then
                
                Save_Packing_No = StrConv(PTANAREC.Packing_No, vbUnicode)
                Save_Rank = StrConv(PTANAREC.Rank, vbUnicode)
                Save_Page = StrConv(PTANAREC.Page_cnt, vbUnicode)
                LCNT = LMAX + 1
            
            End If
                
            If Save_Rank <> StrConv(PTANAREC.Rank, vbUnicode) Then
                
                Printer.Print
                Printer.Print
                LCNT = LCNT + 2
                
                Save_Rank = StrConv(PTANAREC.Rank, vbUnicode)
                Save_Page = StrConv(PTANAREC.Page_cnt, vbUnicode)
            End If
            
            If Save_Page <> StrConv(PTANAREC.Page_cnt, vbUnicode) Then
                
                LCNT = LMAX + 1
                Save_Page = StrConv(PTANAREC.Page_cnt, vbUnicode)
            
            End If
            
            
            If LCNT > LMAX Then
            
                Call Print_Head(LCNT, Save_Page)
                Pri_Rank = ""
            End If
            
            
            Tab_Pos = MGN_L + 2
            If Pri_Rank <> StrConv(PTANAREC.Rank, vbUnicode) Then
                Printer.Print Tab(Tab_Pos);
                Printer.Print StrConv(PTANAREC.Rank, vbUnicode);
                Pri_Rank = StrConv(PTANAREC.Rank, vbUnicode)
            End If
            
            Tab_Pos = Tab_Pos + 5
            Printer.Print Tab(Tab_Pos);
            If IsNumeric(Left(StrConv(PTANAREC.RETUREN01, vbUnicode), 2)) Then
                Printer.Print StrConv(PTANAREC.RETUREN01, vbUnicode);
            End If
            
            Tab_Pos = Tab_Pos + 8
            Printer.Print Tab(Tab_Pos);
            If IsNumeric(Left(StrConv(PTANAREC.RETUREN02, vbUnicode), 2)) Then
                Printer.Print StrConv(PTANAREC.RETUREN02, vbUnicode);
            End If
            
            Tab_Pos = Tab_Pos + 8
            Printer.Print Tab(Tab_Pos);
            If IsNumeric(Left(StrConv(PTANAREC.RETUREN03, vbUnicode), 2)) Then
                Printer.Print StrConv(PTANAREC.RETUREN03, vbUnicode);
            End If
            
            Tab_Pos = Tab_Pos + 8
            Printer.Print Tab(Tab_Pos);
            If IsNumeric(Left(StrConv(PTANAREC.RETUREN04, vbUnicode), 2)) Then
                Printer.Print StrConv(PTANAREC.RETUREN04, vbUnicode);
            End If
            
            Tab_Pos = Tab_Pos + 8
            Printer.Print Tab(Tab_Pos);
            If IsNumeric(Left(StrConv(PTANAREC.RETUREN05, vbUnicode), 2)) Then
                Printer.Print StrConv(PTANAREC.RETUREN05, vbUnicode);
            End If
            
            Tab_Pos = Tab_Pos + 8
            Printer.Print Tab(Tab_Pos);
            If IsNumeric(Left(StrConv(PTANAREC.RETUREN06, vbUnicode), 2)) Then
                Printer.Print StrConv(PTANAREC.RETUREN06, vbUnicode);
            End If
            
            Tab_Pos = Tab_Pos + 8
            Printer.Print Tab(Tab_Pos);
            If IsNumeric(Left(StrConv(PTANAREC.RETUREN07, vbUnicode), 2)) Then
                Printer.Print StrConv(PTANAREC.RETUREN07, vbUnicode);
            End If
            
            Tab_Pos = Tab_Pos + 8
            Printer.Print Tab(Tab_Pos);
            If IsNumeric(Left(StrConv(PTANAREC.RETUREN08, vbUnicode), 2)) Then
                Printer.Print StrConv(PTANAREC.RETUREN08, vbUnicode);
            End If
            
            Tab_Pos = Tab_Pos + 8
            Printer.Print Tab(Tab_Pos);
            If IsNumeric(Left(StrConv(PTANAREC.RETUREN09, vbUnicode), 2)) Then
                Printer.Print StrConv(PTANAREC.RETUREN09, vbUnicode);
            End If
            
            Tab_Pos = Tab_Pos + 8
            Printer.Print Tab(Tab_Pos);
            If IsNumeric(Left(StrConv(PTANAREC.RETUREN10, vbUnicode), 2)) Then
                Printer.Print StrConv(PTANAREC.RETUREN10, vbUnicode);
            End If
            Printer.Print
            LCNT = LCNT + 1
        
        End If
        
        com = BtOpGetNext
    Loop

    Printer.EndDoc

    
    Call Input_UnLock         '��ʍ��ڃ��b�N����

    Print_Proc = False

End Function

Private Sub Print_Head(LCNT As Integer, Page_No As String)
'----------------------------------------------------------------------------
'                  �w�b�_�[�R���g���[������
'----------------------------------------------------------------------------
Dim Start_Page  As Integer
Dim End_Page    As Integer
Dim i           As Integer
Dim Tab_Pos     As Integer
    
    If LCNT <> LCTL Then
        Printer.NewPage
    End If

    For i = 1 To MGN_U
        Printer.Print
    Next i

    Printer.Print


    Printer.Print Tab(20);
    Printer.Print "������  �����ʒI���X�g  ������";
    Printer.Print Tab(60);
    Printer.Print Pdate & "  " & Ptime;
    Printer.Print "     P." & Format(Printer.Page, "000")
    
    Printer.Print

    Printer.Print Tab(MGN_L);
    Printer.Print "�������F" & StrConv(PTANAREC.Packing_No, vbUnicode)
    Printer.Print

    Printer.Print Tab(MGN_L);
    Printer.Print "�����N";
    Printer.Print Tab(MGN_L + 8);
    Printer.Print "�I�@��"
    Printer.Print
    
    Start_Page = CInt(Page_No & "0")
    End_Page = CInt(Page_No & "9")
    
    Tab_Pos = MGN_L + 8
    For i = Start_Page To End_Page
        If i > UBound(Soko_Tbl) Then
            Exit For
        End If
    
        Printer.Print Tab(Tab_Pos);
        Printer.Print Soko_Tbl(i).Soko_No;
        Tab_Pos = Tab_Pos + 8
    Next i
    Printer.Print
    Printer.Print

    LCNT = MGN_U + 8

End Sub
Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1080201.MousePointer = vbHourglass

    Call Ctrl_Lock(F1080201)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1080201)


    F1080201.MousePointer = vbDefault

End Sub


Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        
        Case 7          '�f�[�^�o��
        
            Beep
            ans = MsgBox("�u�����ʒI�f�[�^�v�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Data_Proc() Then
                    Unload Me
                End If
            End If
            Text(ptxPACKING_NO).SetFocus
        
        
        Case 8          '���
            
            
            Beep
            ans = MsgBox("�u�����ʒI���X�g�v������܂����H", vbYesNo + vbQuestion, "�m�F����")
            If ans = vbYes Then
                If Print_Proc() Then
                    Unload Me
                End If
            End If
            Text(ptxPACKING_NO).SetFocus
                    
        Case 11         '�I��
            Unload Me
        Case Else
            Beep
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
    
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = Trim(c)
                                '�����ʒI���X�g�t�@�C������荞��
    If GetIni("FILE", "PTANA_DATA", "SYS", c) Then
        Beep
        MsgBox "�����ʒI���X�g�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    PTANA_DATA = Trim(c)
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�I�ʌ����}�X�^�n�o�d�m
    If TPACKING_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�����ʒI���X�g����t�@�C���n�o�d�m
    If PTANA_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
    If Soko_INF_Proc() Then     '�q�ɏ��̓W�J
        Unload Me
    End If
                                
                                '����t�H���g�ݒ�
    With NormalFont
        .NAME = F1080201.FontName
        .Size = F1080201.FontSize
    End With
    Set Printer.Font = NormalFont
                                

    Show
                                
    
    Text(ptxPACKING_NO).SetFocus
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

Dim sts As Integer
    
                                            
                                            '�q�Ƀ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�q�Ƀ}�X�^")
        End If
    End If
                                            
                                            '�I�ʌ����}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�ʌ����}�X�^")
        End If
    End If
                                            '�����ʒI���X�g����t�@�C���b�k�n�r�d
    sts = BTRV(BtOpClose, PTANA_POS, PTANAREC, Len(PTANAREC), K0_PTANA, Len(K0_PTANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�����ʒI���X�g����t�@�C��")
        End If
    End If
    
    sts = BTRV(BtOpReset, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K0_TPACKING, Len(K0_TPACKING), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1080201 = Nothing

    End
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
                
        
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i

End Sub

Private Function Data_Make_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u�����ʒI�ԃ��X�g����t�@�C���v�쐬 ����
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim Soko_Cnt        As Integer
Dim Seq_No          As Integer

Dim Save_Packing_No As String * 4
Dim Save_Rank       As String * 3

Dim i               As Integer
Dim j               As Integer
Dim k               As Integer


    Data_Make_Proc = True
        
    '---------------------------------------------------------- '�S���R�[�h�폜
    com = BtOpGetFirst
    Do
        
        Do
            DoEvents
            sts = BTRV(com + BtSNoWait, PTANA_POS, PTANAREC, Len(PTANAREC), K0_PTANA, Len(K0_PTANA), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrEOF
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<PTANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                
                Case Else
                    Call File_Error(sts, com + BtSNoWait, "�����ʒI���X�g����t�@�C��")
                    Exit Function
            End Select
        
        Loop
    
        If sts = BtErrEOF Then
            Exit Do
        End If
        
        Do
            
            sts = BTRV(BtOpDelete, PTANA_POS, PTANAREC, Len(PTANAREC), K0_PTANA, Len(K0_PTANA), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<PTANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpDelete, "�����ʒI���X�g����t�@�C��")
                    Exit Function
            End Select
        
        Loop
        
        com = BtOpGetNext
    
    Loop
    
    '---------------------------------------------------------- '�������[�W�J
    Call UniCode_Conv(K1_TPACKING.Packing_No, Text(ptxPACKING_NO).Text)
    Call UniCode_Conv(K1_TPACKING.Rank, "")
    Call UniCode_Conv(K1_TPACKING.Soko_No, "")
    Call UniCode_Conv(K1_TPACKING.Retu, "")
    Call UniCode_Conv(K1_TPACKING.Ren, "")
    
    
    Save_Packing_No = ""
'    Max_Gyo = -1
    
    com = BtOpGetGreater
    Do
        
        DoEvents
        
        sts = BTRV(com, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K1_TPACKING, Len(K1_TPACKING), 1)
        Select Case sts
            Case BtNoErr
                If Len(Trim(Text(ptxPACKING_NO).Text)) = 0 Then
                Else
                    If Trim(Text(ptxPACKING_NO).Text) <> Trim(StrConv(TPACKINGREC.Packing_No, vbUnicode)) Then
                        If Len(Trim(Save_Packing_No)) <> 0 Then
                        
                            If Data_Put_Proc(Save_Packing_No, Save_Rank) Then
                                Exit Function
                            End If
                        
                        End If
                        
                        
                        
                        Exit Do
                    End If
                End If
            Case BtErrEOF
                If Len(Trim(Save_Packing_No)) <> 0 Then
                        
                    If Data_Put_Proc(Save_Packing_No, Save_Rank) Then
                        Exit Function
                    End If
                        
                End If
            
                Exit Do
            
            Case Else
                Call File_Error(sts, com, "�I�ʌ����}�X�^����t�@�C��")
                Exit Function
        End Select
    
        If com = BtOpGetGreater Then
            Save_Packing_No = StrConv(TPACKINGREC.Packing_No, vbUnicode)
            Save_Rank = StrConv(TPACKINGREC.Rank, vbUnicode)
            i = -1
        
            Erase Location
        
        End If
    
        If Trim(Text(ptxPACKING_NO).Text) <> Trim(StrConv(TPACKINGREC.Packing_No, vbUnicode)) Or _
            Save_Rank <> StrConv(TPACKINGREC.Rank, vbUnicode) Then
            
        
            '�����^�����N�u���[�N�Ńf�[�^�o��
            If Data_Put_Proc(Save_Packing_No, Save_Rank) Then
                Exit Function
            End If
            
            Save_Rank = StrConv(TPACKINGREC.Rank, vbUnicode)
            
            
            Erase Location
            
            i = -1
        
        End If
        
'        i = i + 1
'        If i > Max_Gyo Then
'            Max_Gyo = i
'        End If
        
'        ReDim Preserve Location(0 To i)
'        ReDim Preserve Location(i).RetuRen(UBound(Soko_Tbl))
            
        For j = 0 To UBound(Soko_Tbl)
            
            If Soko_Tbl(j).Soko_No = StrConv(TPACKINGREC.Soko_No, vbUnicode) Then
                                
                
                
                If i = -1 Then
                
                    i = i + 1
'                    If i > Max_Gyo Then
'                        Max_Gyo = i
'                    End If
        
                    ReDim Preserve Location(0 To i)
                    ReDim Preserve Location(i).RetuRen(UBound(Soko_Tbl))
                    Location(i).RetuRen(j).Retu = StrConv(TPACKINGREC.Retu, vbUnicode)
                    Location(i).RetuRen(j).Ren = StrConv(TPACKINGREC.Ren, vbUnicode)
                Else
                    For i = 0 To UBound(Location)
                    
                        If Not IsNumeric(Location(i).RetuRen(j).Retu) Then
                            Location(i).RetuRen(j).Retu = StrConv(TPACKINGREC.Retu, vbUnicode)
                            Location(i).RetuRen(j).Ren = StrConv(TPACKINGREC.Ren, vbUnicode)
                            Exit For
                        End If
                    Next i
                            
                    If i > UBound(Location) Then
                            
                        i = UBound(Location) + 1
'                        If i > Max_Gyo Then
'                            Max_Gyo = i
'                        End If
                        ReDim Preserve Location(0 To i)
                        ReDim Preserve Location(i).RetuRen(UBound(Soko_Tbl))
                        
                        Location(i).RetuRen(j).Retu = StrConv(TPACKINGREC.Retu, vbUnicode)
                        Location(i).RetuRen(j).Ren = StrConv(TPACKINGREC.Ren, vbUnicode)
                    End If
                End If
                
                Exit For
            
            End If
        
        Next j
        
        com = BtOpGetNext
    
    Loop


    Data_Make_Proc = False

End Function

Private Function Soko_INF_Proc() As Integer
'----------------------------------------------------------------------------
'                   �q�ɏ��̓W�J����
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim com         As Integer

Dim i           As Integer
    
    
    Soko_INF_Proc = True
    
    
    i = -1
    
    
    com = BtOpGetFirst
    Do
        
        DoEvents
        sts = BTRV(com, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
                
            Case Else
                Call File_Error(sts, com, "�����ʒI���X�g����t�@�C��")
                Exit Function
        End Select
        
        If StrConv(SOKOREC.SOKO_BUN, vbUnicode) <> BUN_KASO Then
            i = i + 1
            ReDim Preserve Soko_Tbl(0 To i)
            Soko_Tbl(i).Soko_No = StrConv(SOKOREC.Soko_No, vbUnicode)
                                '����|�Q���u1�y�[�W������̑q�ɐ��i�q�Ɂ^�Łj�v
'            Soko_Tbl(i).Page_cnt = Fix((i + 1) / (Retu_Max + 2)) + 1
        
        End If
    
        com = BtOpGetNext
   
    Loop

    If i = -1 Then
        
        Beep
        MsgBox "���q�ɂ��L��܂���B�}�X�^���e���m�F���ĉ������B"
        Exit Function
    
    End If
    Soko_INF_Proc = False

End Function

Private Function Data_Put_Proc(Packing_No As String, Rank As String) As Integer
'----------------------------------------------------------------------------
'                   �u�����p�I���X�g����t�@�C���v�o�͏���
'----------------------------------------------------------------------------
Dim i           As Integer
Dim j           As Integer
Dim k           As Integer

Dim Soko_Retu   As String * 1
Dim Pos         As Integer
    
    
Dim Seq_No      As Long
    
Dim sts         As Integer
Dim ans         As Integer
    
    
    Data_Put_Proc = True
    
    
    Seq_No = 0
    
    i = -1
    Do
        
        DoEvents
        
        i = i + 1
        
        If i > UBound(Location) Then
            Exit Do
        End If
              
        
        
        Call UniCode_Conv(PTANAREC.Packing_No, Packing_No)
        Call UniCode_Conv(PTANAREC.Rank, Rank)
        
        Soko_Retu = "0"
        
        For j = 0 To UBound(Soko_Tbl)
            
            
            If Soko_Retu <> Left(Format(j, "00"), 1) Then
                
                
                Seq_No = Seq_No + 1
                Call UniCode_Conv(PTANAREC.SEQ, Format(Seq_No, "00000"))
                Call UniCode_Conv(PTANAREC.Page_cnt, Soko_Retu)
                Do
                    sts = BTRV(BtOpInsert, PTANA_POS, PTANAREC, Len(PTANAREC), K0_PTANA, Len(K0_PTANA), 0)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<PTANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Exit Function
                            End If
                        Case Else
                            Call File_Error(sts, BtOpInsert, "�����ʒI���X�g����t�@�C��")
                            Exit Function
                    End Select
                
                
                Loop
                
                
                Call UniCode_Conv(PTANAREC.SOKO_NO01, "")
                Call UniCode_Conv(PTANAREC.RETUREN01, "")
                Call UniCode_Conv(PTANAREC.SOKO_NO02, "")
                Call UniCode_Conv(PTANAREC.RETUREN02, "")
                Call UniCode_Conv(PTANAREC.SOKO_NO03, "")
                Call UniCode_Conv(PTANAREC.RETUREN03, "")
                Call UniCode_Conv(PTANAREC.SOKO_NO04, "")
                Call UniCode_Conv(PTANAREC.RETUREN04, "")
                Call UniCode_Conv(PTANAREC.SOKO_NO05, "")
                Call UniCode_Conv(PTANAREC.RETUREN05, "")
                Call UniCode_Conv(PTANAREC.SOKO_NO06, "")
                Call UniCode_Conv(PTANAREC.RETUREN06, "")
                Call UniCode_Conv(PTANAREC.SOKO_NO07, "")
                Call UniCode_Conv(PTANAREC.RETUREN07, "")
                Call UniCode_Conv(PTANAREC.SOKO_NO08, "")
                Call UniCode_Conv(PTANAREC.RETUREN08, "")
                Call UniCode_Conv(PTANAREC.SOKO_NO09, "")
                Call UniCode_Conv(PTANAREC.RETUREN09, "")
                Call UniCode_Conv(PTANAREC.SOKO_NO10, "")
                Call UniCode_Conv(PTANAREC.RETUREN10, "")
                
                
                Soko_Retu = Left(Format(j, "00"), 1)
            
            End If
            
            
            Pos = CInt(Right(Format(j, "00"), 1))
            
            Select Case Pos
                Case 0
                    Call UniCode_Conv(PTANAREC.SOKO_NO01, Soko_Tbl(j).Soko_No)
                    Call UniCode_Conv(PTANAREC.RETUREN01, (Location(i).RetuRen(j).Retu & "-" & Location(i).RetuRen(j).Ren))
                    
                Case 1
                    Call UniCode_Conv(PTANAREC.SOKO_NO02, Soko_Tbl(j).Soko_No)
                    Call UniCode_Conv(PTANAREC.RETUREN02, (Location(i).RetuRen(j).Retu & "-" & Location(i).RetuRen(j).Ren))
                Case 2
                    Call UniCode_Conv(PTANAREC.SOKO_NO03, Soko_Tbl(j).Soko_No)
                    Call UniCode_Conv(PTANAREC.RETUREN03, (Location(i).RetuRen(j).Retu & "-" & Location(i).RetuRen(j).Ren))
                Case 3
                    Call UniCode_Conv(PTANAREC.SOKO_NO04, Soko_Tbl(j).Soko_No)
                    Call UniCode_Conv(PTANAREC.RETUREN04, (Location(i).RetuRen(j).Retu & "-" & Location(i).RetuRen(j).Ren))
                Case 4
                    Call UniCode_Conv(PTANAREC.SOKO_NO05, Soko_Tbl(j).Soko_No)
                    Call UniCode_Conv(PTANAREC.RETUREN05, (Location(i).RetuRen(j).Retu & "-" & Location(i).RetuRen(j).Ren))
                Case 5
                    Call UniCode_Conv(PTANAREC.SOKO_NO06, Soko_Tbl(j).Soko_No)
                    Call UniCode_Conv(PTANAREC.RETUREN06, (Location(i).RetuRen(j).Retu & "-" & Location(i).RetuRen(j).Ren))
                Case 6
                    Call UniCode_Conv(PTANAREC.SOKO_NO07, Soko_Tbl(j).Soko_No)
                    Call UniCode_Conv(PTANAREC.RETUREN07, (Location(i).RetuRen(j).Retu & "-" & Location(i).RetuRen(j).Ren))
                Case 7
                    Call UniCode_Conv(PTANAREC.SOKO_NO08, Soko_Tbl(j).Soko_No)
                    Call UniCode_Conv(PTANAREC.RETUREN08, (Location(i).RetuRen(j).Retu & "-" & Location(i).RetuRen(j).Ren))
                Case 8
                    Call UniCode_Conv(PTANAREC.SOKO_NO09, Soko_Tbl(j).Soko_No)
                    Call UniCode_Conv(PTANAREC.RETUREN09, (Location(i).RetuRen(j).Retu & "-" & Location(i).RetuRen(j).Ren))
                Case 9
                    Call UniCode_Conv(PTANAREC.SOKO_NO10, Soko_Tbl(j).Soko_No)
                    Call UniCode_Conv(PTANAREC.RETUREN10, (Location(i).RetuRen(j).Retu & "-" & Location(i).RetuRen(j).Ren))
            End Select
            
            
        Next j
    
        If Soko_Retu <> "0" Then
        
            Seq_No = Seq_No + 1
            Call UniCode_Conv(PTANAREC.SEQ, Format(Seq_No, "00000"))
            Call UniCode_Conv(PTANAREC.Page_cnt, Soko_Retu)
            Do
                sts = BTRV(BtOpInsert, PTANA_POS, PTANAREC, Len(PTANAREC), K0_PTANA, Len(K0_PTANA), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<PTANA.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpInsert, "�����ʒI���X�g����t�@�C��")
                        Exit Function
                End Select
                
            Loop
        
            Call UniCode_Conv(PTANAREC.SOKO_NO01, "")
            Call UniCode_Conv(PTANAREC.RETUREN01, "")
            Call UniCode_Conv(PTANAREC.SOKO_NO02, "")
            Call UniCode_Conv(PTANAREC.RETUREN02, "")
            Call UniCode_Conv(PTANAREC.SOKO_NO03, "")
            Call UniCode_Conv(PTANAREC.RETUREN03, "")
            Call UniCode_Conv(PTANAREC.SOKO_NO04, "")
            Call UniCode_Conv(PTANAREC.RETUREN04, "")
            Call UniCode_Conv(PTANAREC.SOKO_NO05, "")
            Call UniCode_Conv(PTANAREC.RETUREN05, "")
            Call UniCode_Conv(PTANAREC.SOKO_NO06, "")
            Call UniCode_Conv(PTANAREC.RETUREN06, "")
            Call UniCode_Conv(PTANAREC.SOKO_NO07, "")
            Call UniCode_Conv(PTANAREC.RETUREN07, "")
            Call UniCode_Conv(PTANAREC.SOKO_NO08, "")
            Call UniCode_Conv(PTANAREC.RETUREN08, "")
            Call UniCode_Conv(PTANAREC.SOKO_NO09, "")
            Call UniCode_Conv(PTANAREC.RETUREN09, "")
            Call UniCode_Conv(PTANAREC.SOKO_NO10, "")
            Call UniCode_Conv(PTANAREC.RETUREN10, "")
        
        
        End If
    
    Loop
        
    
    Data_Put_Proc = False

End Function

Private Function Data_Proc() As Integer
'----------------------------------------------------------------------------
'                   �u�����ʒI�f�[�^�i�b�r�u�j�v�o�͏���
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim com             As Integer

Dim Ret             As Integer
Dim FileNo          As Integer
Dim fileName        As String

Dim c               As String * 128
Dim Soko_No         As String * 2

'Dim Save_Packing_No As String * 4
'Dim Save_Rank       As String * 3
    
    Data_Proc = True
    
    Call Input_Lock         '��ʍ��ڃ��b�N
    
    FileNo = FreeFile
    fileName = PTANA_DATA
    
    Ret = InStr(1, Trim(fileName), ".") - 1
    fileName = Left(Trim(fileName), Ret) & "_" & Last_JGYOBU & Right(Trim(fileName), Len(Trim(fileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (fileName) For Output As FileNo
    
    
    Write #FileNo, "�����ʒI�f�[�^"
    Write #FileNo, "������", "�����N", "�I��"
    
    
    Call UniCode_Conv(K1_TPACKING.Packing_No, Text(ptxPACKING_NO).Text)
    Call UniCode_Conv(K1_TPACKING.Rank, "")
    Call UniCode_Conv(K1_TPACKING.Soko_No, "")
    Call UniCode_Conv(K1_TPACKING.Retu, "")
    Call UniCode_Conv(K1_TPACKING.Ren, "")
        
    com = BtOpGetGreater
    
'    Save_Packing_No = ""
'    Save_Rank = ""
    
    
    Do
        DoEvents
        sts = BTRV(com, TPACKING_POS, TPACKINGREC, Len(TPACKINGREC), K1_TPACKING, Len(K1_TPACKING), 1)
        Select Case sts
            Case BtNoErr
                
                If Len(Trim(Text(ptxPACKING_NO).Text)) <> 0 Then
                    If Trim(Text(ptxPACKING_NO).Text) <> Trim(StrConv(TPACKINGREC.Packing_No, vbUnicode)) Then
                        Exit Do
                    End If
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�I�ʌ����}�X�^")
                Exit Function
        End Select
    
    
'        If Save_Packing_No <> StrConv(TPACKINGREC.Packing_No, vbUnicode) Then
'            Write #FileNo, StrConv(TPACKINGREC.Packing_No, vbUnicode),
'            Write #FileNo, StrConv(TPACKINGREC.Rank, vbUnicode),
'            Save_Packing_No = StrConv(TPACKINGREC.Packing_No, vbUnicode)
'            Save_Rank = StrConv(TPACKINGREC.Rank, vbUnicode)
'        Else
'            Write #FileNo, "", "",
'
'        End If
        
'        If Save_Rank <> StrConv(TPACKINGREC.Rank, vbUnicode) Then
'            Write #FileNo, "", StrConv(TPACKINGREC.Rank, vbUnicode),
'            Save_Rank = StrConv(TPACKINGREC.Rank, vbUnicode)
'        Else
'            Write #FileNo, "", "",
'        End If
        
        Write #FileNo, StrConv(TPACKINGREC.Packing_No, vbUnicode),
        Write #FileNo, StrConv(TPACKINGREC.Rank, vbUnicode),
        
        
        If GetIni("SOKO_NO", StrConv(TPACKINGREC.Soko_No, vbUnicode), "SYS", c) Then
            Soko_No = StrConv(TPACKINGREC.Soko_No, vbUnicode)
        Else
            Soko_No = Trim(c)
        End If
        
        Write #FileNo, Soko_No & "-" & _
                        StrConv(TPACKINGREC.Retu, vbUnicode) & "-" & _
                        StrConv(TPACKINGREC.Ren, vbUnicode)
        
        
        
        
        com = BtOpGetNext
    Loop
    
    
    
    Close #FileNo
    
    Call Input_UnLock         '��ʍ��ڃ��b�N����
    Beep
    MsgBox "�u" & fileName & "�v�͐���ɏo�͂���܂����B"
    
    Data_Proc = False
    
    Exit Function
    
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox fileName & "���g�p���ł��B"
        Call Input_UnLock         '��ʍ��ڃ��b�N����
        Data_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
    End If

    Call Input_UnLock       '��ʍ��ډ���
    
    Data_Proc = False

End Function
