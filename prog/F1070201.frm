VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form F1070201 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�I��������"
   ClientHeight    =   9885
   ClientLeft      =   795
   ClientTop       =   2295
   ClientWidth     =   15045
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
      Size            =   10.5
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   15045
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   1095
      MaxLength       =   2
      TabIndex        =   17
      Top             =   120
      Width           =   435
   End
   Begin VB.ComboBox Combo 
      Height          =   330
      Index           =   0
      Left            =   13800
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   14
      Top             =   8880
      Visible         =   0   'False
      Width           =   855
   End
   Begin TrueDBGrid60.TDBGrid TDBGrid1 
      Height          =   7815
      Left            =   0
      OleObjectBlob   =   "F1070201.frx":0000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   14775
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
      Top             =   9000
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
      Top             =   9000
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
      Top             =   9000
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�ĕ\��"
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
      Top             =   9000
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
      Top             =   9000
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
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�񍐏�"
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
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command 
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
      Top             =   9000
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
      Top             =   9000
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
      Top             =   9000
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "�X�@�V"
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9000
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   9
      Left            =   8865
      TabIndex        =   29
      Top             =   120
      Width           =   330
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   8025
      TabIndex        =   28
      Top             =   120
      Width           =   645
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�� �� ���F"
      Height          =   255
      Index           =   7
      Left            =   6870
      TabIndex        =   27
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   6
      Left            =   11280
      TabIndex        =   26
      Top             =   600
      Width           =   330
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   10650
      TabIndex        =   25
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�m�f�����F"
      Height          =   255
      Index           =   4
      Left            =   9600
      TabIndex        =   24
      Top             =   600
      Width           =   960
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   255
      Index           =   3
      Left            =   8865
      TabIndex        =   23
      Top             =   600
      Width           =   330
   End
   Begin VB.Label Label 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   8025
      TabIndex        =   22
      Top             =   600
      Width           =   645
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�n�j�����F"
      Height          =   255
      Index           =   1
      Left            =   6870
      TabIndex        =   21
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "�i�w���͈�"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   20
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblZEN_LOC 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   ")"
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   18
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�q�Ɏw��"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   16
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   33
      Left            =   12960
      TabIndex        =   15
      Top             =   9000
      Visible         =   0   'False
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
      Left            =   11340
      TabIndex        =   13
      Top             =   9000
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
Attribute VB_Name = "F1070201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxSOKO% = 0                  '�Ώۑq��

Private Const Text_Max% = 0                 '��ʍ��ڕʍő���ޯ��



Private Const pcmbNAIGAI% = 0           '�����O



Private Const colNo% = 0                '�s�ԍ� 2008.10.21


Private Const colHin_Gai% = 1           '�i�ԁi�O���j
Private Const colHOST_ZAIKO% = 2        '���_�݌�

Private Const colPPSC_ZAI_QTY% = 3      'PPSC�݌�
Private Const colBU_ZAI_QTY% = 4        'BU�݌�


Private Const colPOS_ZAIKO% = 5         '�o�n�r�݌�
Private Const colST_Zaiko% = 6          '�W���I�ԁ��W���I�ԍ݌�
Private Const colEE1_ZAIKO% = 7         '�ʒu���P�I�ԁ��݌�
Private Const colEE2_ZAIKO% = 8         '�ʒu���Q�I�ԁ��݌�
Private Const colEE3_ZAIKO% = 9         '�ʒu���R�I�ԁ��݌�
Private Const colETC_ZAIKO% = 10         '���̑��݌�
Private Const colCHECK_MARK_OK% = 11     '�n�j
Private Const colCHECK_MARK_NG% = 12    '�m�f
Private Const colPRINT_YMD% = 13        '������t
Private Const colINPUT_YMD% = 14        '���͓��t
Private Const colSAI_QTY% = 15          '���ِ�

Private Const colJGYOBU% = 16           '���ƕ�



Private STOCK       As New XArrayDB
Private Data_Flg    As Boolean


Private Const Min_Row% = 1              '�ŏ��s��
'Private Const Max_Row& = 2000           '�ő�s��
Dim Max_Row     As Long                 '���X�g�{�b�N�X�ő�\������

Private Const Min_Col% = 0              '�ŏ���
Private Const Max_Col% = 16             '�ő��

Private Stock_OK_DATA   As String       'OK�f�[�^
Private Stock_NG_DATA   As String       'NG�f�[�^

Private Sort_Tbl(colNo To colJGYOBU) _
                As Integer              '��Ă̐��� 0:���� 1:�~��    2008.10.21


Private OK_CNT      As Long             '2008.10.21
Private NG_CNT      As Long             '2008.10.21


'Private Const Lost_Update_Day$ = "[F107020] 2018.11.16 13:00"
Private Const Lost_Update_Day$ = "[F107020] 2018.12.01 11:00"

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   �c�a�f�������ݒ胁�C������
'----------------------------------------------------------------------------

Dim sts         As Integer
Dim com         As Integer

Dim ans         As Integer
Dim i           As Integer
Dim Row         As Long
    
    List_Disp_Proc = True
                                    
    Call Input_Lock
                                    
    '��ď��̏�����    2008.10.21
    For i = 0 To UBound(Sort_Tbl)
        Sort_Tbl(i) = 0             '��̫�ď���
    Next i

    Sort_Tbl(colNo) = 9             '��ď��O
    Sort_Tbl(colJGYOBU) = 9         '��ď��O
                                    
                                    
                                    
    OK_CNT = 0  '2008.10.21
    NG_CNT = 0  '2008.10.21
                                    
                                    '�e�[�u�����Z�b�g
    Data_Flg = False
    Set STOCK = Nothing
    
    
    If Last_JGYOBU = "*" Then
        '�SBU
        For i = 0 To UBound(JGYOBU_T)
        
            If JGYOBU_T(i).CODE = "*" Or JGYOBU_T(i).CODE = SHIZAI Then
            Else
        
                Call UniCode_Conv(K4_STOCK.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
                Call UniCode_Conv(K4_STOCK.ST_SOKO, Text(ptxSOKO).Text)
                Call UniCode_Conv(K4_STOCK.ST_RETU, "")
                Call UniCode_Conv(K4_STOCK.ST_REN, "")
                Call UniCode_Conv(K4_STOCK.ST_DAN, "")
                Call UniCode_Conv(K4_STOCK.HIN_GAI, "")
                
                
                Row = Min_Row - 1
                    
                com = BtOpGetGreater
                
                Do
                    DoEvents
                    
                    sts = BTRV(com, STOCK_POS, STOCKREC, Len(STOCKREC), K4_STOCK, Len(K4_STOCK), 4)
                
                    Select Case sts
                        Case BtNoErr
                    
                            If StrConv(STOCKREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                                
                                Exit Do
                            
                            End If
                            
                            
                            If Trim(Text(ptxSOKO).Text) <> StrConv(STOCKREC.ST_SOKO, vbUnicode) Then
                                Exit Do
                            End If
                            
                        Case BtErrEOF
                            Exit Do
                        Case Else
                            Call File_Error(sts, com, "�I�����f�[�^")
                            List_Disp_Proc = SYS_ERR
                    End Select
                                            '���ƕ� KEY��ڰ�
                    Data_Flg = True
                
                    Row = Row + 1
                    If Row > Max_Row Then
                        Beep
                        MsgBox "�ő�\���s���𒴂��܂����B"
                        Exit Do
                    End If
                            
                    Call Grid_Set_Proc(Row, StrConv(STOCKREC.JGYOBU, vbUnicode))
                
                    com = BtOpGetNext   '����
                Loop
            End If
        Next i
    
    Else
        '�P��BU
        Call UniCode_Conv(K1_STOCK.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K1_STOCK.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
        Call UniCode_Conv(K1_STOCK.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K1_STOCK.ST_RETU, "")
        Call UniCode_Conv(K1_STOCK.ST_REN, "")
        Call UniCode_Conv(K1_STOCK.ST_DAN, "")
        Call UniCode_Conv(K1_STOCK.HIN_GAI, "")
        
        
        Row = Min_Row - 1
            
        com = BtOpGetGreater
        
        Do
            DoEvents
            
            sts = BTRV(com, STOCK_POS, STOCKREC, Len(STOCKREC), K1_STOCK, Len(K1_STOCK), 1)
        
            Select Case sts
                Case BtNoErr
            
                    If StrConv(STOCKREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                        StrConv(STOCKREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                        
                        Exit Do
                    
                    End If
                    
                    
                    If Trim(Text(ptxSOKO).Text) <> StrConv(STOCKREC.ST_SOKO, vbUnicode) Then
                        Exit Do
                    End If
                    
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com, "�I�����f�[�^")
                    List_Disp_Proc = SYS_ERR
            End Select
                                    '���ƕ� KEY��ڰ�
            Data_Flg = True
        
            Row = Row + 1
            If Row > Max_Row Then
                Beep
                MsgBox "�ő�\���s���𒴂��܂����B"
                Exit Do
            End If
                    
            Call Grid_Set_Proc(Row, StrConv(STOCKREC.JGYOBU, vbUnicode))
        
            com = BtOpGetNext   '����
        Loop
                                
    End If
                                
                                
                                'DB�e�[�u�������N
    Set TDBGrid1.Array = STOCK
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    
    
    
    If TDBGrid1.ApproxCount > 0 Then
        Label(8).Caption = Format(STOCK.UpperBound(1), "#0")
    Else
        Label(8).Caption = 0
    End If
    Label(2).Caption = Format(OK_CNT, "#0")
    Label(5).Caption = Format(NG_CNT, "#0")
    
    
    Call Input_UnLock
    
    List_Disp_Proc = False

End Function
                                    '��ʏ�����Ԃ�ݒ肷��
Private Sub Clear_Field(Mode As Integer)
    Sort_Tbl(colJGYOBU) = 9         '��ď��O
'Dim i As Integer
'
'    For i = Mode To Text_Max
'        Text(i).Text = ""
'    Next i
'
End Sub


Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> vbKeyReturn Then
        Exit Sub
    End If
    

    Text(ptxSOKO).SetFocus


End Sub


Private Sub Command_Click(Index As Integer)

Dim sts As Integer
Dim yn  As Integer
Dim i       As Integer  '2008.10.21
    
    Select Case Index
        
        Case 0                              '�X�V
        
        
            If Not Data_Flg Then
                Exit Sub
            End If

            OK_CNT = 0
            NG_CNT = 0
            
            
            
            If TDBGrid1.ApproxCount > 0 Then
            
            
                For i = Min_Row To STOCK.UpperBound(1)
                    
                    If STOCK(i, colCHECK_MARK_OK) = True Then
                        OK_CNT = OK_CNT + 1
                    End If
                    
                    If STOCK(i, colCHECK_MARK_NG) = True Then
                        NG_CNT = NG_CNT + 1
                    End If
                    
                    
                Next i
                    
                        
                Label(2).Caption = Format(OK_CNT, "#0")
                Label(5).Caption = Format(NG_CNT, "#0")
            End If
        
        
        
            yn = MsgBox("�X�V���܂����H", vbYesNo, "�m�F����")
            
            If yn = vbYes Then
        
                If Update_Proc() Then
                    Unload Me
                End If
            
            
                MsgBox ("�X�V���I�����܂���")       '2018.11.16
            
            
            End If
        
        Case 4                              '�񍐏��쐬
                        
            If Not Data_Flg Then
                Exit Sub
            End If
            
            
            OK_CNT = 0
            NG_CNT = 0
            
            If TDBGrid1.ApproxCount > 0 Then
            
                For i = Min_Row To STOCK.UpperBound(1)
                    
                    If STOCK(i, colCHECK_MARK_OK) = True Then
                        OK_CNT = OK_CNT + 1
                    End If
                    
                    If STOCK(i, colCHECK_MARK_NG) = True Then
                        NG_CNT = NG_CNT + 1
                    End If
                    
                    
                Next i
                
            End If
                
                    
            Label(2).Caption = Format(OK_CNT, "#0")
            Label(5).Caption = Format(NG_CNT, "#0")
            
            
            yn = MsgBox("�񍐏��쐬���܂����H", vbYesNo, "�m�F����")
            If yn = vbYes Then
            
                If Report_Proc() Then
                    Unload Me
                End If
                
                If List_Disp_Proc() Then
                    Unload Me
                End If
            
'                MsgBox ("�񍐏��쐬���I�����܂���")       '2018.11.16
            
            
            
            End If
            
        Case 7                              '�ĕ\��
            If List_Disp_Proc() Then
                Unload Me
            End If
        Case 11                            '�I��
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub

Private Sub Form_DblClick()
'    PrintForm
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

Dim i       As Integer
Dim c       As String * 128
Dim sts     As Integer
    
'    If App.PrevInstance Then
'        Beep
'        MsgBox "����v���O�������s���ł��B"
'        End
'    End If
    
    
    
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�I�������́@�J�n", Me.hwnd, 0)
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
    LOG_F = RTrim(c)
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If



    '�SBU�^�s��
    ReDim Preserve JGYOBU_T(UBound(JGYOBU_T) + 1)
    JGYOBU_T(UBound(JGYOBU_T)).CODE = ""
    JGYOBU_T(UBound(JGYOBU_T)).NAME = "-"
    JGYOBU_T(UBound(JGYOBU_T)).COLOR = 12
    
    
    
    ReDim Preserve JGYOBU_T(UBound(JGYOBU_T) + 1)
    JGYOBU_T(UBound(JGYOBU_T)).CODE = "*"
    JGYOBU_T(UBound(JGYOBU_T)).NAME = "�SBU"
    JGYOBU_T(UBound(JGYOBU_T)).COLOR = 12


    ReDim Preserve JGYOBU_T(UBound(JGYOBU_T) + 1)
    JGYOBU_T(UBound(JGYOBU_T)).CODE = ""
    JGYOBU_T(UBound(JGYOBU_T)).NAME = "-"
    JGYOBU_T(UBound(JGYOBU_T)).COLOR = 12






    For i = 0 To UBound(JGYOBU_T)
'        If JGYOBU_T(i).CODE = " " Then
'            Unload SubMenu(i)
'            Exit For
'        End If

        Load SubMenu(i + 1)
        
        If RTrim(JGYOBU_T(i).NAME) = "-" Then
            SubMenu(i).Checked = False
        End If
        
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If Trim(JGYOBU_T(i).CODE) = "" Then
        Else
            If JGYOBU_T(i).CODE = Last_JGYOBU Then
                F1070201.Caption = "�I�������́i" + RTrim(JGYOBU_T(i).NAME) + ")" & " " & Lost_Update_Day
                SubMenu(i).Checked = True
                LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
                LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
    '            LabJIGYO.BorderStyle = 1
            Else
                SubMenu(i).Checked = False
            End If
        End If
    Next i
    Unload SubMenu(i)

                                
                                '�ő�\��������荞��
'    If GetIni(App.EXEName, "LISTMAX", "SYS", c) Then
    If GetIni(App.EXEName, "LISTMAX", App.EXEName, c) Then
        Beep
        MsgBox "�ő�\�������̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    Max_Row = CLng(RTrim(c))
                                '�񍐏��t�@�C����荞��
    If GetIni("FILE", "Stock_OK_DATA", "SYS", c) Then
        Beep
        MsgBox "�񍐏��o�͗p�t�@�C����[Stock_OK_DATA]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    Stock_OK_DATA = RTrim(c)
    If GetIni("FILE", "Stock_NG_DATA", "SYS", c) Then
        Beep
        MsgBox "�񍐏��o�͗p�t�@�C����[Stock_NG_DATA]�̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    Stock_NG_DATA = RTrim(c)
                                
                                
                                
                                '�q�Ƀ}�X�^�n�o�d�m
    If SOKO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                '�I�����f�[�^�n�o�d�m
    If STOCK_Open(BtOpenNomal) Then
        Unload Me
    End If

'                                '�����O��荞��
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & "    " & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & "    " & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
                                
'    Combo(pcmbNAIGAI).SetFocus
    Text(ptxSOKO).SetFocus
                                '��ʏ����ݒ�
    Call Clear_Field(0)
    
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
                                            
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            
                                            '�I�����f�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�����f�[�^")
        End If
    End If
    
    
    sts = BTRV(BtOpReset, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1070201 = Nothing

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
    F1070201.Caption = "�I�������́i" + RTrim(JGYOBU_T(Index).NAME) + ")" & " " & Lost_Update_Day
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)

End Sub


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1070201.MousePointer = vbHourglass

    Call Ctrl_Lock(F1070201)


End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1070201)


    F1070201.MousePointer = vbDefault

End Sub
Private Sub Grid_Set_Proc(Row As Long, stJGYOBU As String)
'----------------------------------------------------------------------------
'                   �I�����f�[�^���f����������
'----------------------------------------------------------------------------

Dim Edit        As String
Dim Num_Edit    As String


    STOCK.ReDim Min_Row, Row, Min_Col, Max_Col
                                            
                                            
                                            '�s�ԍ� 2008.10.21
    STOCK(Row, colNo) = Row
                                            
                                            
                                            '�i�ځi�O���j
    STOCK(Row, colHin_Gai) = StrConv(STOCKREC.HIN_GAI, vbUnicode)
                                            '�z�X�g�݌�
    STOCK(Row, colHOST_ZAIKO) = Format(CLng(StrConv(STOCKREC.HOST_ZAIKO, vbUnicode)), "#0")
                                            
                                            'PPSC�݌�
    
    If IsNumeric(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)) Then
    
        STOCK(Row, colPPSC_ZAI_QTY) = Format(CLng(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)), "#0")
    Else
        STOCK(Row, colPPSC_ZAI_QTY) = 0
    End If
                                            
                                            
                                            'BU�݌�
    If IsNumeric(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)) Then
        STOCK(Row, colBU_ZAI_QTY) = Format(CLng(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)), "#0")
                                            
    Else
        STOCK(Row, colBU_ZAI_QTY) = 0
    End If
                                            
                                            
                                            '�o�n�r�݌�
    STOCK(Row, colPOS_ZAIKO) = Format(CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)), "#0")
                                            '�W���I�݌�
    If Len(Trim(StrConv(STOCKREC.ST_SOKO, vbUnicode))) <> 0 Then
        Edit = StrConv(STOCKREC.ST_SOKO, vbUnicode) & "-"
        Edit = Edit & StrConv(STOCKREC.ST_RETU, vbUnicode) & "-"
        Edit = Edit & StrConv(STOCKREC.ST_REN, vbUnicode) & "-"
        Edit = Edit & StrConv(STOCKREC.ST_DAN, vbUnicode) & " "
        Num_Edit = Format(CLng(StrConv(STOCKREC.ST_ZAIKO, vbUnicode)), "#0")
        If Len(Num_Edit) < 6 Then
            Num_Edit = Space(6 - Len(Num_Edit)) & Num_Edit
        End If
        Edit = Edit & Num_Edit
        STOCK(Row, colST_Zaiko) = Edit
    End If
                                            '�ʒu���݌ɂP
    If Len(Trim(StrConv(STOCKREC.EE1_LOCATION, vbUnicode))) <> 0 Then
        Edit = Left(StrConv(STOCKREC.EE1_LOCATION, vbUnicode), 2) & "-"
        Edit = Edit & Mid(StrConv(STOCKREC.EE1_LOCATION, vbUnicode), 3, 2) & "-"
        Edit = Edit & Mid(StrConv(STOCKREC.EE1_LOCATION, vbUnicode), 5, 2) & "-"
        Edit = Edit & Mid(StrConv(STOCKREC.EE1_LOCATION, vbUnicode), 7, 2) & " "
        Num_Edit = Format(CLng(StrConv(STOCKREC.EE1_ZAIKO, vbUnicode)), "#0")
        If Len(Num_Edit) < 6 Then
            Num_Edit = Space(6 - Len(Num_Edit)) & Num_Edit
        End If
        Edit = Edit & Num_Edit
        STOCK(Row, colEE1_ZAIKO) = Edit
    End If
                                            '�ʒu���݌ɂQ
    If Len(Trim(StrConv(STOCKREC.EE2_LOCATION, vbUnicode))) <> 0 Then
        Edit = Left(StrConv(STOCKREC.EE2_LOCATION, vbUnicode), 2) & "-"
        Edit = Edit & Mid(StrConv(STOCKREC.EE2_LOCATION, vbUnicode), 3, 2) & "-"
        Edit = Edit & Mid(StrConv(STOCKREC.EE2_LOCATION, vbUnicode), 5, 2) & "-"
        Edit = Edit & Mid(StrConv(STOCKREC.EE2_LOCATION, vbUnicode), 7, 2) & " "
        Num_Edit = Format(CLng(StrConv(STOCKREC.EE2_ZAIKO, vbUnicode)), "#0")
        If Len(Num_Edit) < 6 Then
            Num_Edit = Space(6 - Len(Num_Edit)) & Num_Edit
        End If
        Edit = Edit & Num_Edit
        STOCK(Row, colEE2_ZAIKO) = Edit
    End If
                                            '�ʒu���݌ɂR
    If Len(Trim(StrConv(STOCKREC.EE3_LOCATION, vbUnicode))) <> 0 Then
        Edit = Left(StrConv(STOCKREC.EE3_LOCATION, vbUnicode), 2) & "-"
        Edit = Edit & Mid(StrConv(STOCKREC.EE3_LOCATION, vbUnicode), 3, 2) & "-"
        Edit = Edit & Mid(StrConv(STOCKREC.EE3_LOCATION, vbUnicode), 5, 2) & "-"
        Edit = Edit & Mid(StrConv(STOCKREC.EE3_LOCATION, vbUnicode), 7, 2)
        Num_Edit = Format(CLng(StrConv(STOCKREC.EE3_ZAIKO, vbUnicode)), "#0")
        If Len(Num_Edit) < 6 Then
            Num_Edit = Space(6 - Len(Num_Edit)) & Num_Edit
        End If
        Edit = Edit & Num_Edit
        STOCK(Row, colEE3_ZAIKO) = Edit
    End If
                                            '���̑��݌�
    STOCK(Row, colETC_ZAIKO) = Format(CLng(StrConv(STOCKREC.ETC_ZAIKO, vbUnicode)), "#")
                                            '�`�F�b�N�}�[�N
    Select Case StrConv(STOCKREC.CHECK_MARK, vbUnicode)
        Case " ", "0"
            STOCK(Row, colCHECK_MARK_OK) = False
            STOCK(Row, colCHECK_MARK_NG) = False
        Case "1"
            STOCK(Row, colCHECK_MARK_OK) = True
            STOCK(Row, colCHECK_MARK_NG) = False
            OK_CNT = OK_CNT + 1         '2008.10.21
        Case "2"
            STOCK(Row, colCHECK_MARK_OK) = False
            STOCK(Row, colCHECK_MARK_NG) = True
            NG_CNT = NG_CNT + 1         '2008.10.21

    End Select

    STOCK(Row, colPRINT_YMD) = Left(StrConv(STOCKREC.PRINT_YMD, vbUnicode), 4) & "/" & Mid(StrConv(STOCKREC.PRINT_YMD, vbUnicode), 5, 2) & "/" & Right(StrConv(STOCKREC.PRINT_YMD, vbUnicode), 2)
    
    If Len(Trim(StrConv(STOCKREC.INPUT_YMD, vbUnicode))) <> 0 Then
        STOCK(Row, colINPUT_YMD) = Left(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 4) & "/" & Mid(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 5, 2) & "/" & Right(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 2)
    Else
        STOCK(Row, colINPUT_YMD) = ""
    End If


    STOCK(Row, colSAI_QTY) = Format(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode)), "#0")



    '���ƕ�
    STOCK(Row, colJGYOBU) = stJGYOBU
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    
'�w�b�_�����łr�n�q�s�ǉ�   2008.10.21
    
Dim i   As Integer
    
    If Sort_Tbl(ColIndex) = 0 Then
        Sort_Tbl(ColIndex) = 1
    Else
        If Sort_Tbl(ColIndex) = 1 Then
            Sort_Tbl(ColIndex) = 0
        End If
    
    End If
    
    If Sort_Tbl(ColIndex) = 0 Or Sort_Tbl(ColIndex) = 1 Then
                    
                    
        If ColIndex = colHOST_ZAIKO Or ColIndex = colPPSC_ZAI_QTY Or ColIndex = colBU_ZAI_QTY Or ColIndex = colPOS_ZAIKO Or ColIndex = colETC_ZAIKO Then
                    
            STOCK.QuickSort Min_Row, STOCK.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_NUMBER
            
        Else
                
            STOCK.QuickSort Min_Row, STOCK.UpperBound(1), ColIndex, Sort_Tbl(ColIndex), XTYPE_STRING
        End If
        
        For i = Min_Row To STOCK.UpperBound(1)
            STOCK(i, colNo) = i
        Next i
        
        
        Set TDBGrid1.Array = STOCK
        
        TDBGrid1.ReBind
        TDBGrid1.Update
        TDBGrid1.MoveFirst


    End If

End Sub

Private Sub TDBGrid1_LostFocus()
    
Dim i   As Integer
    
    If Not Data_Flg Then
        Exit Sub
    End If
    
    Set TDBGrid1.Array = STOCK
    TDBGrid1.Refresh
    
    TDBGrid1.Update
    
    
    For i = 1 To STOCK.UpperBound(1)
        If STOCK(i, colCHECK_MARK_OK) And _
            STOCK(i, colCHECK_MARK_NG) Then
            
            MsgBox "OK/NG��I�����ĉ�����"
            
            STOCK(i, colCHECK_MARK_OK) = False
            STOCK(i, colCHECK_MARK_NG) = False
            
            TDBGrid1.Refresh
        
        End If
    Next i
    

End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Not Data_Flg Then
        Exit Sub
    End If

    
    If LastCol = colCHECK_MARK_OK Or _
        LastCol = colCHECK_MARK_NG Then
    
        If STOCK(LastRow, colCHECK_MARK_OK) And _
            STOCK(LastRow, colCHECK_MARK_NG) Then
            MsgBox "OK/NG��I�����ĉ�����"
        
            STOCK(LastRow, colCHECK_MARK_OK) = False
            STOCK(LastRow, colCHECK_MARK_NG) = False
        
            Set TDBGrid1.Array = STOCK
            TDBGrid1.Refresh
    
            TDBGrid1.Update
        
        End If
    
    End If

End Sub

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   �I�����f�[�^�X�V����
'----------------------------------------------------------------------------
Dim sts         As Integer
Dim ans         As Integer
Dim i           As Integer
Dim Stock_Max   As Long


    Update_Proc = True

    Call Input_Lock

    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�I�������́@�X�V�J�n", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)




    For i = 1 To STOCK.UpperBound(1)
        Call UniCode_Conv(K0_STOCK.JGYOBU, STOCK(i, colJGYOBU))
        Call UniCode_Conv(K0_STOCK.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
        Call UniCode_Conv(K0_STOCK.HIN_GAI, STOCK(i, colHin_Gai))
        
        Do
            sts = BTRV(BtOpGetEqual, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                Case BtErrKeyNotFound
                    '�l�����Ȃ��������p��
                    Exit Do
                Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                    Beep
                    ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                    If ans = vbCancel Then
                        Call File_Error(sts, BtOpGetEqual, "�I���f�[�^")
                        
                        Exit Function
                    End If
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "�I���f�[�^")
                    Exit Function
            End Select
        
        Loop
             
        If sts <> BtErrKeyNotFound Then
            Select Case StrConv(STOCKREC.CHECK_MARK, vbUnicode)
                Case " "
                    If STOCK(i, colCHECK_MARK_OK) Then
                        Call UniCode_Conv(STOCKREC.INPUT_YMD, Format(Now, "YYYYMMDD"))
                        Call UniCode_Conv(STOCKREC.CHECK_MARK, "1")
                    Else
                        If STOCK(i, colCHECK_MARK_NG) Then
                            Call UniCode_Conv(STOCKREC.INPUT_YMD, Format(Now, "YYYYMMDD"))
                            Call UniCode_Conv(STOCKREC.CHECK_MARK, "2")
                        Else
                            Call UniCode_Conv(STOCKREC.CHECK_MARK, "0")
                        End If
                    End If
                Case "0"
                    If STOCK(i, colCHECK_MARK_OK) Then
                        Call UniCode_Conv(STOCKREC.INPUT_YMD, Format(Now, "YYYYMMDD"))
                        Call UniCode_Conv(STOCKREC.CHECK_MARK, "1")
                    Else
                        If STOCK(i, colCHECK_MARK_NG) Then
                            Call UniCode_Conv(STOCKREC.INPUT_YMD, Format(Now, "YYYYMMDD"))
                            Call UniCode_Conv(STOCKREC.CHECK_MARK, "2")
                        End If
                    End If
                Case "1"
                    If STOCK(i, colCHECK_MARK_OK) Then
                    Else
                        If STOCK(i, colCHECK_MARK_NG) Then
                            Call UniCode_Conv(STOCKREC.INPUT_YMD, Format(Now, "YYYYMMDD"))
                            Call UniCode_Conv(STOCKREC.CHECK_MARK, "2")
                        
                        Else
                            Call UniCode_Conv(STOCKREC.INPUT_YMD, "")
                            Call UniCode_Conv(STOCKREC.CHECK_MARK, "0")
                        End If
                    End If
                Case "2"
                    If STOCK(i, colCHECK_MARK_OK) Then
                        Call UniCode_Conv(STOCKREC.INPUT_YMD, Format(Now, "YYYYMMDD"))
                        Call UniCode_Conv(STOCKREC.CHECK_MARK, "1")
                    Else
                        If STOCK(i, colCHECK_MARK_NG) Then
                        Else
                            Call UniCode_Conv(STOCKREC.INPUT_YMD, "")
                            Call UniCode_Conv(STOCKREC.CHECK_MARK, "0")
                        End If
                    End If
            End Select
        
                        
            If STOCK(i, colCHECK_MARK_NG) Then
                                                
                If CLng(STOCK(i, colSAI_QTY)) < 0 Then
                                
                    Call UniCode_Conv(STOCKREC.SAI_QTY, Format(CLng(STOCK(i, colSAI_QTY)), "00000000"))
                Else
                    Call UniCode_Conv(STOCKREC.SAI_QTY, Format(CLng(STOCK(i, colSAI_QTY)), "000000000"))
                End If
            Else
            
                Call UniCode_Conv(STOCKREC.SAI_QTY, "000000000")
            
            End If
            
            Do
                
                sts = BTRV(BtOpUpdate, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Call File_Error(sts, BtOpUpdate, "�I���f�[�^")
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "�I���f�[�^")
                        Exit Function
                End Select
            
            Loop
        
        End If
    
    
    
    Next i

    Call Input_UnLock
    
    
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�I�������́@�X�V�I��", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
    
    
    
    Update_Proc = False
    

End Function

Private Function Report_Proc() As Integer
'----------------------------------------------------------------------------
'                   �񍐏��쐬����
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim com             As Integer
Dim ans             As Integer

Dim Save_Soko       As String * 2

Dim FileNo_OK       As Integer
Dim FileNo_NG       As Integer
Dim FileName_OK     As String
Dim FileName_NG     As String

Dim OK_NO           As Integer
Dim NG_NO           As Integer


Dim c               As String * 128
Dim Soko_No         As String * 2

Dim Data_Mode       As Integer
Dim Skip_Flg        As Boolean

Dim Fsw             As Integer
Dim i               As Integer


    Report_Proc = True
                                
                                
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�I�������́@�񍐏��쐬�J�n", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
                                
                                
    If Last_JGYOBU = "*" Then
                                    '�����p���̊m�F
        Call UniCode_Conv(K5_STOCK.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
        Call UniCode_Conv(K5_STOCK.ST_SOKO, "")
        Call UniCode_Conv(K5_STOCK.CHECK_MARK, "")
        
        sts = BTRV(BtOpGetGreaterEqual, STOCK_POS, STOCKREC, Len(STOCKREC), K2_STOCK, Len(K2_STOCK), 2)
        Select Case sts
            Case BtNoErr
                
                If StrConv(STOCKREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                    Beep
                    MsgBox "�Ώۃf�[�^���L��܂���B"
                    Report_Proc = False
                    Exit Function
                End If
                                            
            
                If StrConv(STOCKREC.CHECK_MARK, vbUnicode) = " " Or StrConv(STOCKREC.CHECK_MARK, vbUnicode) = "0" Then
                    Beep
                    ans = MsgBox("�����͂̃f�[�^���L��܂��B�������f�[�^���c���܂����H" & vbCrLf & "�������f�[�^���c���u�͂��v�A�������f�[�^�������u�������v�A�����L�����Z���u�L�����Z���v", vbYesNoCancel + vbDefaultButton3, "�m�F����")
            
                    Select Case ans
                        Case vbCancel
                            Report_Proc = False
                            Exit Function
                                                                        
                        Case vbYes
                            Data_Mode = 1           '�f�[�^���c��
                        Case vbNo
                            Data_Mode = 2           '�f�[�^����
                    End Select
            
                End If
            Case BtErrEOF
                
                Beep
                MsgBox "�Ώۃf�[�^���L��܂���B"
                Report_Proc = False
                Exit Function
            
            Case Else
                
                Call File_Error(sts, BtOpGetGreaterEqual, "�I�����f�[�^")
                Exit Function
        
        End Select
        
                                            'OK�f�[�^ Open
        FileName_OK = Stock_OK_DATA
        sts = InStr(1, Trim(FileName_OK), ".") - 1
        
        
        FileName_OK = Left(Trim(FileName_OK), sts) & "_" & Format(Now, "YYYYMMDDHHMMSS") & Right(Trim(FileName_OK), Len(Trim(FileName_OK)) - sts)
        
        On Error GoTo Error_Proc
        
        FileNo_OK = FreeFile
        Open (FileName_OK) For Output As FileNo_OK
                                            'NG�f�[�^ Open
        FileName_NG = Stock_NG_DATA
        sts = InStr(1, Trim(FileName_NG), ".") - 1
        FileName_NG = Left(Trim(FileName_NG), sts) & "_" & Format(Now, "YYYYMMDDHHMMSS") & Right(Trim(FileName_NG), Len(Trim(FileName_NG)) - sts)
        
        FileNo_NG = FreeFile
        Open (FileName_NG) For Output As FileNo_NG
                                            
        On Error GoTo 0
                                            
                                            '�g�����U�N�V�����J�n
        sts = BTRV(BtOpBeginConcurrentTransaction, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
            Exit Function
        End If
        
        Call UniCode_Conv(K4_STOCK.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
        Call UniCode_Conv(K4_STOCK.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K4_STOCK.ST_RETU, "")
        Call UniCode_Conv(K4_STOCK.ST_REN, "")
        Call UniCode_Conv(K4_STOCK.ST_DAN, "")
        Call UniCode_Conv(K4_STOCK.HIN_GAI, "")
        
        com = BtOpGetGreater
        
        Fsw = True
        
        Do
            DoEvents
            
            Do
                sts = BTRV(com, STOCK_POS, STOCKREC, Len(STOCKREC), K4_STOCK, Len(K4_STOCK), 4)
                Select Case sts
                    Case BtNoErr
                        If StrConv(STOCKREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                            
                            sts = BtErrEOF
                        
                        End If
                        If StrConv(STOCKREC.ST_SOKO, vbUnicode) <> Trim(Text(ptxSOKO).Text) Then
                                    
                        
                            sts = BtErrEOF
                        
                        End If
                        
                        
                        
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Call File_Error(sts, BtOpGetEqual, "�I���f�[�^")
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�I���f�[�^")
                        GoTo Abort_Tran
                End Select
            Loop
        
            If sts = BtErrEOF Then
                Exit Do
            End If
                                            
            
            Skip_Flg = False
            If Data_Mode = 1 Then
                If StrConv(STOCKREC.CHECK_MARK, vbUnicode) = " " Or StrConv(STOCKREC.CHECK_MARK, vbUnicode) = "0" Then
                    Skip_Flg = True
                End If
            End If
                        
                        
            
            If Skip_Flg Then
            Else
                If Fsw Then
                
                    Write #FileNo_OK, "��", "�i��", "�i��", "�I��", "BU�݌�", "PPSC�݌�", "POS�݌�", "��", " �{��", "����", "�ƍ����t"
                    Write #FileNo_NG, "��", "�i��", "�i��", "�I��", "BU�݌�", "PPSC�݌�", "POS�݌�", "��", " �{��", "����", "�ƍ����t"
                
                                
                
                
                    Fsw = False
                
                End If
                                                
                                                '�f�[�^�̐U�蕪��
                Select Case StrConv(STOCKREC.CHECK_MARK, vbUnicode)
                    Case " ", "0"       '�������@OR ������
                    '------------------- �Ȃɂ����Ȃ�
                    Case "1"            'OK
                    '------------------- �񍐏��쐬
                            
                            
                        OK_NO = OK_NO + 1
                        Write #FileNo_OK, Format(OK_NO, "#0"),                              'No
                        Write #FileNo_OK, Trim(StrConv(STOCKREC.HIN_GAI, vbUnicode)),       '�i��
                            
                            
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(STOCKREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(STOCKREC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(STOCKREC.HIN_GAI, vbUnicode))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�I���f�[�^")
                                GoTo Abort_Tran
                        End Select
                        Write #FileNo_OK, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)),      '�i��
                        
                        Write #FileNo_OK, Trim(StrConv(STOCKREC.ST_SOKO, vbUnicode)) & _
                                            Trim(StrConv(STOCKREC.ST_RETU, vbUnicode)) & _
                                            Trim(StrConv(STOCKREC.ST_REN, vbUnicode)) & _
                                            Trim(StrConv(STOCKREC.ST_DAN, vbUnicode)),      '�W���I��
                                            
                                                                                            'BU�݌�
                        If IsNumeric(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)) Then
                            Write #FileNo_OK, Format(CLng(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)), "#0"),
                        Else
                            Write #FileNo_OK, 0,
                        End If
                                                                                            
                            
                                                                                            'PPSC�݌�
                        If IsNumeric(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)) Then
                            Write #FileNo_OK, Format(CLng(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)), "#0"),
                        Else
                            Write #FileNo_OK, 0,
                        End If
                                                                                            
                                                                                            'POS�݌�
                        Write #FileNo_OK, Format(CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)), "#0"),
                                                                                            '��
                        Write #FileNo_OK, Format(Abs(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode))), "#0"),
                            
                        If CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode)) = 0 Then
                        
                            Write #FileNo_OK, , ,
                        
                        Else
                           If CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode)) > 0 Then
                                                                                            '�{
                                Write #FileNo_OK, Format(Abs(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode))), "#0"), ,
                            
                           Else
                                                                                            '�|
                                Write #FileNo_OK, , Format(Abs(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode))), "#0"),
                            End If
                        End If
                   
                        If Trim(StrConv(STOCKREC.INPUT_YMD, vbUnicode)) <> "" Then
                            Write #FileNo_OK, Left(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 2)
                        Else
                            Write #FileNo_OK, ,
                    
                        End If
                    
                    
                    Case "2"            'NG
                        
                        NG_NO = NG_NO + 1
                        Write #FileNo_NG, Format(NG_NO, "#0"),                              'No
                        Write #FileNo_NG, Trim(StrConv(STOCKREC.HIN_GAI, vbUnicode)),       '�i��
                            
                            
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(STOCKREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(STOCKREC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(STOCKREC.HIN_GAI, vbUnicode))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�I���f�[�^")
                                GoTo Abort_Tran
                        End Select
                        Write #FileNo_NG, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)),      '�i��
                        
                        Write #FileNo_NG, Trim(StrConv(STOCKREC.ST_SOKO, vbUnicode)) & _
                                            Trim(StrConv(STOCKREC.ST_RETU, vbUnicode)) & _
                                            Trim(StrConv(STOCKREC.ST_REN, vbUnicode)) & _
                                            Trim(StrConv(STOCKREC.ST_DAN, vbUnicode)),      '�W���I��
                            
                                                                                            'BU�݌�
                        If IsNumeric(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)) Then
                            Write #FileNo_NG, Format(CLng(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)), "#0"),
                        Else
                            Write #FileNo_NG, 0,
                        End If
                                                                                            'PPSC�݌�
                        If IsNumeric(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)) Then
                            Write #FileNo_NG, Format(CLng(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)), "#0"),
                        Else
                            Write #FileNo_NG, 0,
                        End If
                                                                                            
                                                                                            'POS�݌�
                        Write #FileNo_NG, Format(CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)), "#0"),
                                                                                            '��
                        Write #FileNo_NG, Format(Abs(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode))), "#0"),
                            
                        If CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode)) = 0 Then
                        
                            Write #FileNo_NG, , ,
                        
                        Else
                           If CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode)) > 0 Then
                                                                                            '�{
                                Write #FileNo_NG, Format(Abs(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode))), "#0"), ,
                            
                           Else
                                                                                            '�|
                                Write #FileNo_NG, , Format(Abs(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode))), "#0"),
                            End If
                        End If
                   
                        If Trim(StrConv(STOCKREC.INPUT_YMD, vbUnicode)) <> "" Then
                            Write #FileNo_NG, Left(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 2)
                        Else
                            Write #FileNo_NG, , ,
                    
                        End If
                
                
                
                End Select
            
            
                Do
                    sts = BTRV(BtOpDelete, STOCK_POS, STOCKREC, Len(STOCKREC), K1_STOCK, Len(K1_STOCK), 1)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound
                            '�ُ킾�������p��
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Call File_Error(sts, BtOpDelete, "�I���f�[�^")
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, BtOpDelete, "�I���f�[�^")
                            GoTo Abort_Tran
                    End Select
                
                Loop
            
            End If
        
            com = BtOpGetNext
        Loop
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
                                
    Else
                                    '�����p���̊m�F
        Call UniCode_Conv(K2_STOCK.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K2_STOCK.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
        Call UniCode_Conv(K2_STOCK.ST_SOKO, "")
        Call UniCode_Conv(K2_STOCK.CHECK_MARK, "")
        
        sts = BTRV(BtOpGetGreaterEqual, STOCK_POS, STOCKREC, Len(STOCKREC), K2_STOCK, Len(K2_STOCK), 2)
        Select Case sts
            Case BtNoErr
                
                If StrConv(STOCKREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(STOCKREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                    Beep
                    MsgBox "�Ώۃf�[�^���L��܂���B"
                    Report_Proc = False
                    Exit Function
                End If
                                            
    '            If StrConv(STOCKREC.CHECK_MARK, vbUnicode) = " " Then
    '                Beep
    '                ans = MsgBox("�������̃f�[�^���L��܂��B�������p�����܂����H", vbYesNo + vbDefaultButton2, "�m�F����")
    '                If ans = vbNo Then
    '                    Report_Proc = False
    '                    Exit Function
    '                End If
    '            End If
    '
    '            If StrConv(STOCKREC.CHECK_MARK, vbUnicode) = "0" Then
    '                Beep
    '                ans = MsgBox("�����͂̃f�[�^���L��܂��B�������p�����܂����H", vbYesNo + vbDefaultButton2, "�m�F����")
    '                If ans = vbNo Then
    '                    Report_Proc = False
    '                    Exit Function
    '                End If
    '            End If
            
                If StrConv(STOCKREC.CHECK_MARK, vbUnicode) = " " Or StrConv(STOCKREC.CHECK_MARK, vbUnicode) = "0" Then
                    Beep
                    ans = MsgBox("�����͂̃f�[�^���L��܂��B�������f�[�^���c���܂����H" & vbCrLf & "�������f�[�^���c���u�͂��v�A�������f�[�^�������u�������v�A�����L�����Z���u�L�����Z���v", vbYesNoCancel + vbDefaultButton3, "�m�F����")
            
                    Select Case ans
                        Case vbCancel
                            Report_Proc = False
                            Exit Function
                                                                        
                        Case vbYes
                            Data_Mode = 1           '�f�[�^���c��
                        Case vbNo
                            Data_Mode = 2           '�f�[�^����
                    End Select
            
                End If
            Case BtErrEOF
                
                Beep
                MsgBox "�Ώۃf�[�^���L��܂���B"
                Report_Proc = False
                Exit Function
            
            Case Else
                
                Call File_Error(sts, BtOpGetGreaterEqual, "�I�����f�[�^")
                Exit Function
        
        End Select
        
                                            'OK�f�[�^ Open
        FileName_OK = Stock_OK_DATA
        sts = InStr(1, Trim(FileName_OK), ".") - 1
        FileName_OK = Left(Trim(FileName_OK), sts) & "_" & Last_JGYOBU & "_" & Format(Now, "YYYYMMDDHHMMSS") & Right(Trim(FileName_OK), Len(Trim(FileName_OK)) - sts)
        
        On Error GoTo Error_Proc
        
        FileNo_OK = FreeFile
        Open (FileName_OK) For Output As FileNo_OK
                                            'NG�f�[�^ Open
        FileName_NG = Stock_NG_DATA
        sts = InStr(1, Trim(FileName_NG), ".") - 1
        FileName_NG = Left(Trim(FileName_NG), sts) & "_" & Last_JGYOBU & "_" & Format(Now, "YYYYMMDDHHMMSS") & Right(Trim(FileName_NG), Len(Trim(FileName_NG)) - sts)
        
        FileNo_NG = FreeFile
        Open (FileName_NG) For Output As FileNo_NG
                                            
        On Error GoTo 0
                                            
                                            '�g�����U�N�V�����J�n
        sts = BTRV(BtOpBeginConcurrentTransaction, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
        If sts <> BtNoErr Then
            Call File_Error(sts, BtOpBeginConcurrentTransaction, "")
            Exit Function
        End If
        
        Call UniCode_Conv(K1_STOCK.JGYOBU, Last_JGYOBU)
        Call UniCode_Conv(K1_STOCK.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
        Call UniCode_Conv(K1_STOCK.ST_SOKO, Text(ptxSOKO).Text)
        Call UniCode_Conv(K1_STOCK.ST_RETU, "")
        Call UniCode_Conv(K1_STOCK.ST_REN, "")
        Call UniCode_Conv(K1_STOCK.ST_DAN, "")
        Call UniCode_Conv(K1_STOCK.HIN_GAI, "")
        
        com = BtOpGetGreater
        
        Fsw = True
        
        Do
            DoEvents
            
            Do
                sts = BTRV(com, STOCK_POS, STOCKREC, Len(STOCKREC), K1_STOCK, Len(K1_STOCK), 1)
                Select Case sts
                    Case BtNoErr
                        If StrConv(STOCKREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                            StrConv(STOCKREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                            
                            sts = BtErrEOF
                        
                        End If
                        If StrConv(STOCKREC.ST_SOKO, vbUnicode) <> Trim(Text(ptxSOKO).Text) Then
                                    
                        
                            sts = BtErrEOF
                        
                        End If
                        
                        
                        
                        Exit Do
                    Case BtErrEOF
                        Exit Do
                    Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                        Beep
                        ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                        If ans = vbCancel Then
                            Call File_Error(sts, BtOpGetEqual, "�I���f�[�^")
                            GoTo Abort_Tran
                        End If
                    Case Else
                        Call File_Error(sts, BtOpGetEqual, "�I���f�[�^")
                        GoTo Abort_Tran
                End Select
            Loop
        
            If sts = BtErrEOF Then
                Exit Do
            End If
                                            
            
            Skip_Flg = False
            If Data_Mode = 1 Then
                If StrConv(STOCKREC.CHECK_MARK, vbUnicode) = " " Or StrConv(STOCKREC.CHECK_MARK, vbUnicode) = "0" Then
                    Skip_Flg = True
                End If
            End If
                        
                        
            
            If Skip_Flg Then
            Else
                If Fsw Then
    ''                Save_Soko = StrConv(STOCKREC.ST_SOKO, vbUnicode)
    ''                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
    ''                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    ''                Select Case sts
    ''                    Case BtNoErr
    ''
    ''
    ''
    ''                    Case BtErrKeyNotFound
    ''                        Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
    ''                    Case Else
    ''                        Call File_Error(sts, BtOpGetEqual, "�I���f�[�^")
    ''                        GoTo Abort_Tran
    ''                End Select
                                                
                                                                                
    ''                Write #FileNo_OK, "�z�ƍ��񍐏��i" & StrConv(SOKOREC.SOKO_NAME, vbUnicode), ")"
    ''                Write #FileNo_OK, "���t", "���i�R�[�h", "���_�݌�", "������", "�݌ɒ�����", "�I��", "�݌ɒ������R"
                                                
    ''                Write #FileNo_NG, "�z�s�ƍ��񍐏��i" & StrConv(SOKOREC.SOKO_NAME, vbUnicode), ")"
    ''                Write #FileNo_NG, "���t", "���i�R�[�h", "���_�݌�", "�I��", "���ِ�"
                
                    Write #FileNo_OK, "��", "�i��", "�i��", "�I��", "BU�݌�", "PPSC�݌�", "POS�݌�", "��", " �{��", "����", "�ƍ����t"
                    Write #FileNo_NG, "��", "�i��", "�i��", "�I��", "BU�݌�", "PPSC�݌�", "POS�݌�", "��", " �{��", "����", "�ƍ����t"
                
                                
                
                
                    Fsw = False
                
                End If
                                                
    ''            If Save_Soko <> StrConv(STOCKREC.ST_SOKO, vbUnicode) Then
    ''                Save_Soko = StrConv(STOCKREC.ST_SOKO, vbUnicode)
    ''                Call UniCode_Conv(K0_SOKO.Soko_No, Save_Soko)
    ''                sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
    ''                Select Case sts
    ''                    Case BtNoErr
    ''
    ''
    ''
    ''                    Case BtErrKeyNotFound
    ''                        Call UniCode_Conv(SOKOREC.SOKO_NAME, "")
    ''                    Case Else
    ''                        Call File_Error(sts, BtOpGetEqual, "�I���f�[�^")
    ''                        GoTo Abort_Tran
    ''                End Select
    ''
    ''
    ''                Write #FileNo_OK, "�z�ƍ��񍐏��i" & StrConv(SOKOREC.SOKO_NAME, vbUnicode), ")"
    ''                Write #FileNo_OK, "���t", "���i�R�[�h", "���_�݌�", "������", "�݌ɒ�����", "�I��", "�݌ɒ������R"
    ''
    ''                Write #FileNo_NG, "�z�s�ƍ��񍐏��i" & StrConv(SOKOREC.SOKO_NAME, vbUnicode), ")"
    ''                Write #FileNo_NG, "���t", "���i�R�[�h", "���_�݌�", "�I��", "���ِ�"
    ''
    ''            End If
                                                '�f�[�^�̐U�蕪��
                Select Case StrConv(STOCKREC.CHECK_MARK, vbUnicode)
                    Case " ", "0"       '�������@OR ������
                    '------------------- �Ȃɂ����Ȃ�
                    Case "1"            'OK
                    '------------------- �񍐏��쐬
    ''                    Write #FileNo_OK, StrConv(STOCKREC.INPUT_YMD, vbUnicode),
    ''                    Write #FileNo_OK, StrConv(STOCKREC.HIN_GAI, vbUnicode),
    ''                    Write #FileNo_OK, Format(CLng(StrConv(STOCKREC.HOST_ZAIKO, vbUnicode)), "#0"),
    ''                    Write #FileNo_OK, Format(CLng(StrConv(STOCKREC.HOST_ZAIKO, vbUnicode)), "#0"),
    ''                    Write #FileNo_OK, Format(0),
    ''
    ''
    ''                    If GetIni("SOKO_NO", StrConv(STOCKREC.ST_SOKO, vbUnicode), "SYS", c) Then
    ''                        Soko_No = StrConv(STOCKREC.ST_SOKO, vbUnicode)
    ''                    Else
    ''                        Soko_No = Trim(c)
    ''                    End If
    ''
    ''
    ''                    Write #FileNo_OK, Soko_No & "-" & _
    ''                            StrConv(STOCKREC.ST_RETU, vbUnicode) & "-" & _
    ''                            StrConv(STOCKREC.ST_REN, vbUnicode) & "-" & _
    ''                            StrConv(STOCKREC.ST_DAN, vbUnicode)
                            
                            
                        OK_NO = OK_NO + 1
                        Write #FileNo_OK, Format(OK_NO, "#0"),                              'No
                        Write #FileNo_OK, Trim(StrConv(STOCKREC.HIN_GAI, vbUnicode)),       '�i��
                            
                            
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(STOCKREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(STOCKREC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(STOCKREC.HIN_GAI, vbUnicode))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�I���f�[�^")
                                GoTo Abort_Tran
                        End Select
                        Write #FileNo_OK, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)),      '�i��
                        
                        Write #FileNo_OK, Trim(StrConv(STOCKREC.ST_SOKO, vbUnicode)) & _
                                            Trim(StrConv(STOCKREC.ST_RETU, vbUnicode)) & _
                                            Trim(StrConv(STOCKREC.ST_REN, vbUnicode)) & _
                                            Trim(StrConv(STOCKREC.ST_DAN, vbUnicode)),      '�W���I��
                                            
                                                                                            'BU�݌�
                        If IsNumeric(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)) Then
                            Write #FileNo_OK, Format(CLng(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)), "#0"),
                        Else
                            Write #FileNo_OK, 0,
                        End If
                                                                                            
                            
                                                                                            'PPSC�݌�
                        If IsNumeric(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)) Then
                            Write #FileNo_OK, Format(CLng(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)), "#0"),
                        Else
                            Write #FileNo_OK, 0,
                        End If
                                                                                            
                                                                                            'POS�݌�
                        Write #FileNo_OK, Format(CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)), "#0"),
                                                                                            '��
                        Write #FileNo_OK, Format(Abs(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode))), "#0"),
                            
                        If CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode)) = 0 Then
                        
                            Write #FileNo_OK, , ,
                        
                        Else
                           If CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode)) > 0 Then
                                                                                            '�{
                                Write #FileNo_OK, Format(Abs(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode))), "#0"), ,
                            
                           Else
                                                                                            '�|
                                Write #FileNo_OK, , Format(Abs(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode))), "#0"),
                            End If
                        End If
                   
                        If Trim(StrConv(STOCKREC.INPUT_YMD, vbUnicode)) <> "" Then
                            Write #FileNo_OK, Left(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 2)
                        Else
                            Write #FileNo_OK, ,
                    
                        End If
                    
                    
                    Case "2"            'NG
                        
    ''                    Write #FileNo_NG, StrConv(STOCKREC.INPUT_YMD, vbUnicode),
    ''                    Write #FileNo_NG, StrConv(STOCKREC.HIN_GAI, vbUnicode),
    ''                    Write #FileNo_NG, Format(CLng(StrConv(STOCKREC.HOST_ZAIKO, vbUnicode)), "#0"),
    ''
    ''                    If GetIni("SOKO_NO", StrConv(STOCKREC.ST_SOKO, vbUnicode), "SYS", c) Then
    ''                        Soko_No = StrConv(STOCKREC.ST_SOKO, vbUnicode)
    ''                    Else
    ''                        Soko_No = Trim(c)
    ''                    End If
    ''
    ''
    ''                    Write #FileNo_NG, Soko_No & "-" & _
    ''                            StrConv(STOCKREC.ST_RETU, vbUnicode) & "-" & _
    ''                            StrConv(STOCKREC.ST_REN, vbUnicode) & "-" & _
    ''                            StrConv(STOCKREC.ST_DAN, vbUnicode),
    ''
    ''
    ''                    Write #FileNo_NG, Format(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode)), "#0")
                
                
                
                        NG_NO = NG_NO + 1
                        Write #FileNo_NG, Format(NG_NO, "#0"),                              'No
                        Write #FileNo_NG, Trim(StrConv(STOCKREC.HIN_GAI, vbUnicode)),       '�i��
                            
                            
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(STOCKREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(STOCKREC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(STOCKREC.HIN_GAI, vbUnicode))
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                                Call UniCode_Conv(ITEMREC.HIN_NAME, "")
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "�I���f�[�^")
                                GoTo Abort_Tran
                        End Select
                        Write #FileNo_NG, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)),      '�i��
                        
                        Write #FileNo_NG, Trim(StrConv(STOCKREC.ST_SOKO, vbUnicode)) & _
                                            Trim(StrConv(STOCKREC.ST_RETU, vbUnicode)) & _
                                            Trim(StrConv(STOCKREC.ST_REN, vbUnicode)) & _
                                            Trim(StrConv(STOCKREC.ST_DAN, vbUnicode)),      '�W���I��
                            
                                                                                            'BU�݌�
                        If IsNumeric(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)) Then
                            Write #FileNo_NG, Format(CLng(StrConv(STOCKREC.BU_ZAI_QTY, vbUnicode)), "#0"),
                        Else
                            Write #FileNo_NG, 0,
                        End If
                                                                                            'PPSC�݌�
                        If IsNumeric(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)) Then
                            Write #FileNo_NG, Format(CLng(StrConv(STOCKREC.PPSC_ZAI_QTY, vbUnicode)), "#0"),
                        Else
                            Write #FileNo_NG, 0,
                        End If
                                                                                            
                                                                                            'POS�݌�
                        Write #FileNo_NG, Format(CLng(StrConv(STOCKREC.POS_ZAIKO, vbUnicode)), "#0"),
                                                                                            '��
                        Write #FileNo_NG, Format(Abs(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode))), "#0"),
                            
                        If CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode)) = 0 Then
                        
                            Write #FileNo_NG, , ,
                        
                        Else
                           If CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode)) > 0 Then
                                                                                            '�{
                                Write #FileNo_NG, Format(Abs(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode))), "#0"), ,
                            
                           Else
                                                                                            '�|
                                Write #FileNo_NG, , Format(Abs(CLng(StrConv(STOCKREC.SAI_QTY, vbUnicode))), "#0"),
                            End If
                        End If
                   
                        If Trim(StrConv(STOCKREC.INPUT_YMD, vbUnicode)) <> "" Then
                            Write #FileNo_NG, Left(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 4) & "/" & _
                                                Mid(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 5, 2) & "/" & _
                                                Right(StrConv(STOCKREC.INPUT_YMD, vbUnicode), 2)
                        Else
                            Write #FileNo_NG, , ,
                    
                        End If
                
                
                
                End Select
            
            
                Do
                    sts = BTRV(BtOpDelete, STOCK_POS, STOCKREC, Len(STOCKREC), K1_STOCK, Len(K1_STOCK), 1)
                    Select Case sts
                        Case BtNoErr
                            Exit Do
                        Case BtErrKeyNotFound
                            '�ُ킾�������p��
                            Exit Do
                        Case BtErrFILE_INUSE, BtErrRECORD_INUSE
                            Beep
                            ans = MsgBox("���[���Ńf�[�^�g�p���ł��B<STOCKTAKING.DAT>", vbRetryCancel + vbQuestion, "�m�F����")
                            If ans = vbCancel Then
                                Call File_Error(sts, BtOpDelete, "�I���f�[�^")
                                GoTo Abort_Tran
                            End If
                        Case Else
                            Call File_Error(sts, BtOpDelete, "�I���f�[�^")
                            GoTo Abort_Tran
                    End Select
                
                Loop
            
            End If
        
            com = BtOpGetNext
        Loop
    End If
    
    
End_Tran:
                                        
    Close #FileNo_OK
    Close #FileNo_NG
                                        
                                        '�g�����U�N�V�����I��
    sts = BTRV(BtOpEndTransaction, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpEndTransaction, "")
        GoTo Abort_Tran
    End If
    
    Call Input_UnLock
    
    Beep
    MsgBox "�u" & FileName_OK & "�v�u" & FileName_NG & "�͐���ɏo�͂���܂����B"
    
    
    
    
    '�X�e�[�^�X�E�B���h�E���쐬����
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "�I�������́@�񍐏��쐬�I��", Me.hwnd, 0)
    '�Ō�̗v�f��-1�ɂ����
    '�e�E�B���h�E�̑S�̂̕��̎c��̕���
    '�����I�Ɋ��蓖�Ă�
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)
    
    
    
    
    Report_Proc = False
    
    Exit Function

Abort_Tran:
    
    
    Close #FileNo_OK
    Close #FileNo_NG
    
    sts = BTRV(BtOpAbortTransaction, STOCK_POS, STOCKREC, Len(STOCKREC), K0_STOCK, Len(K0_STOCK), 0)
    If sts <> BtNoErr Then
        Call File_Error(sts, BtOpAbortTransaction, "")
    End If

    Call Input_UnLock
    Exit Function
    
Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox FileName_OK & "���g�p���ł��B"
        Call Input_UnLock         '��ʍ��ڃ��b�N����
        Report_Proc = False
    Else
        MsgBox "Err.Number" & Err.Number
        Report_Proc = True
    End If


End Function


Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim i   As Integer
Dim sts As Integer
    
Dim c   As String * 128
    
    If KeyCode <> vbKeyReturn Then Exit Sub
        
            
    Select Case Index
    
        Case ptxSOKO
    
            Text(ptxSOKO).Text = StrConv(Text(ptxSOKO).Text, vbUpperCase)
            Call UniCode_Conv(K0_SOKO.Soko_No, Text(ptxSOKO).Text)
            sts = BTRV(BtOpGetEqual, SOKO_POS, SOKOREC, Len(SOKOREC), K0_SOKO, Len(K0_SOKO), 0)
            Select Case sts
                Case BtNoErr

'                    If GetIni("F107010", "ZENKAI_LOC" & Trim(Text(ptxSOKO).Text), "SYS", c) Then
                    If GetIni("F107010", "ZENKAI_LOC" & Trim(Text(ptxSOKO).Text), "F107010", c) Then
                        
                        lblZEN_LOC.Caption = ""
                    Else
                        lblZEN_LOC.Caption = RTrim(c)
                    End If
                    
                
                Case BtErrKeyNotFound
                
                    Beep
                    MsgBox ("�q�ɖ��o�^�ł��B")
                    Text(ptxSOKO).SetFocus
                    Exit Sub
                
                Case Else
                    Call File_Error(sts, BtOpGetGreaterEqual, "�q��Ͻ�")
                    Exit Sub
            
            End Select
    
    
            If List_Disp_Proc() Then
                Unload Me
            End If
    
    End Select
        
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i


End Sub
