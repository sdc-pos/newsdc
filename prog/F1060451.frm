VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form F1060451 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "�I���x�����s"
   ClientHeight    =   6315
   ClientLeft      =   2010
   ClientTop       =   2640
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   11295
   StartUpPosition =   2  '��ʂ̒���
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   7
      Left            =   8160
      MaxLength       =   2
      TabIndex        =   24
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   6
      Left            =   7560
      MaxLength       =   2
      TabIndex        =   23
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   5
      Left            =   6840
      MaxLength       =   2
      TabIndex        =   22
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   4
      Left            =   6120
      MaxLength       =   2
      TabIndex        =   21
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   3
      Left            =   5520
      MaxLength       =   2
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   2
      Left            =   4800
      MaxLength       =   2
      TabIndex        =   2
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   4080
      MaxLength       =   2
      TabIndex        =   1
      Top             =   2160
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   3360
      MaxLength       =   2
      TabIndex        =   0
      Top             =   2160
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
      Left            =   10320
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   10
      Left            =   9480
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   9
      Left            =   8640
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "��  ��"
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   7
      Left            =   6480
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   6
      Left            =   5640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   5
      Left            =   4800
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   4
      Left            =   3960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
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
      Index           =   2
      Left            =   1800
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
      Index           =   1
      Left            =   960
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
      Index           =   0
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�|"
      Height          =   255
      Index           =   7
      Left            =   6480
      TabIndex        =   27
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�|"
      Height          =   255
      Index           =   6
      Left            =   7200
      TabIndex        =   26
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�|"
      Height          =   255
      Index           =   5
      Left            =   7920
      TabIndex        =   25
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   4
      Left            =   5880
      TabIndex        =   20
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�|"
      Height          =   255
      Index           =   3
      Left            =   5160
      TabIndex        =   19
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�|"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   18
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�|"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   17
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "���s�I�Ԕ͈�"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   16
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "F1060451"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Const ptxS_Soko_No% = 0
Private Const ptxS_Retu% = 1
Private Const ptxS_Ren% = 2
Private Const ptxS_Dan% = 3

Private Const ptxE_Soko_No% = 4
Private Const ptxE_Retu% = 5
Private Const ptxE_Ren% = 6
Private Const ptxE_Dan% = 7

Private Const Text_Max% = 7

Dim Pri_Name    As Printer


Private Sub Clear_Field()
Dim i As Integer

    For i = 0 To Text_Max
        Text(i).Text = ""
    Next i
End Sub


Private Function Print_Proc() As Integer
'�������

Dim lPrinterHandl   As Long   '���������ق��擾
Dim sts             As Integer
Dim com             As Integer

Dim sEditWK         As String       '�ҏWܰ�
Dim sJis            As String       '�����ϊ�������
    
    
    Print_Proc = True
    
    Call Input_Lock
    
    
'   ����J�n����
    PrinterDriver_Start "�I���x�����s", lPrinterHandl

    


    Call UniCode_Conv(K0_TANA.Soko_No, Text(ptxS_Soko_No).Text)
    Call UniCode_Conv(K0_TANA.Retu, Text(ptxS_Retu).Text)
    Call UniCode_Conv(K0_TANA.Ren, Text(ptxS_Ren).Text)
    Call UniCode_Conv(K0_TANA.Dan, Text(ptxS_Dan).Text)

    com = BtOpGetGreaterEqual


    Do
    
        sts = BTRV(com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
            
                If (StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)) > _
                    (Text(ptxE_Soko_No).Text & Text(ptxE_Retu).Text & Text(ptxE_Ren).Text & Text(ptxE_Dan).Text) Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�I�}�X�^")
                Exit Function
        End Select
    
'       STX�w��
        sEditWK = Chr(&H2)
'       �ް����M�J�n�w��
        sEditWK = sEditWK & Chr(&H1B) & "A"
    
    
        '2007.03.02
        sEditWK = sEditWK & Chr(&H1B) & "A3V+000H+000"
    
    
'       ���ʒu(150�ޯ�),�c�ʒu(15�ޯ�),�����߯�(8�ޯ�),���{��(1�{),�c�{��(2�{)
        
'''        sJis = Kanji_Conv("H", StrConv(StrConv(TANAREC.Soko_No, vbUnicode), vbWide) & "�|" & _
'''                                StrConv(StrConv(TANAREC.Retu, vbUnicode), vbWide) & "�|" & _
'''                                StrConv(StrConv(TANAREC.Ren, vbUnicode), vbWide) & "�|" & _
'''                                StrConv(StrConv(TANAREC.Dan, vbUnicode), vbWide))


'''        sEditWK = sEditWK & Chr(&H1B) & "H0150" & Chr(&H1B) & "V0015" & Chr(&H1B) & "P32" & Chr(&H1B) & "L0102"
'''        sEditWK = sEditWK & Chr(&H1B) & "K2H" & sJis
    
    
    
'        sEditWK = sEditWK & Chr(&H1B) & "H0151" & Chr(&H1B) & "V0015" & Chr(&H1B) & "P32" & Chr(&H1B) & "L0102"
'        sEditWK = sEditWK & Chr(&H1B) & "K2H" & sJis
        
'        sEditWK = sEditWK & Chr(&H1B) & "H0152" & Chr(&H1B) & "V0015" & Chr(&H1B) & "P32" & Chr(&H1B) & "L0102"
'        sEditWK = sEditWK & Chr(&H1B) & "K2H" & sJis
    
'        sEditWK = sEditWK & Chr(&H1B) & "H0153" & Chr(&H1B) & "V0015" & Chr(&H1B) & "P32" & Chr(&H1B) & "L0102"
'        sEditWK = sEditWK & Chr(&H1B) & "K2H" & sJis
    
'        sEditWK = sEditWK & Chr(&H1B) & "H0154" & Chr(&H1B) & "V0015" & Chr(&H1B) & "P32" & Chr(&H1B) & "L0102"
'        sEditWK = sEditWK & Chr(&H1B) & "K2H" & sJis
    
'        sEditWK = sEditWK & Chr(&H1B) & "H0155" & Chr(&H1B) & "V0015" & Chr(&H1B) & "P32" & Chr(&H1B) & "L0102"
'        sEditWK = sEditWK & Chr(&H1B) & "K2H" & sJis
    
    
'2007.03.02        sEditWK = sEditWK & Chr(&H1B) & "H0050" & Chr(&H1B) & "V0005" & Chr(&H1B) & "P32"
        sEditWK = sEditWK & Chr(&H1B) & "H0150" & Chr(&H1B) & "V0005" & Chr(&H1B) & "P32"
        sEditWK = sEditWK & Chr(&H1B) & "$A9999990" & Chr(&H1B) & "$=" & StrConv(TANAREC.Soko_No, vbUnicode) & "-" & _
                                                                        StrConv(TANAREC.Retu, vbUnicode) & "-" & _
                                                                        StrConv(TANAREC.Ren, vbUnicode) & "-" & _
                                                                        StrConv(TANAREC.Dan, vbUnicode)
    

'       ���ʒu(460�ޯ�),�c�ʒu(470�ޯ�),���{��(1�{),�c�{��(1�{)
'2007.03.02        sEditWK = sEditWK & Chr(&H1B) & "H0230" & Chr(&H1B) & "V0080" & Chr(&H1B) & "L0101"
        sEditWK = sEditWK & Chr(&H1B) & "H0200" & Chr(&H1B) & "V0080" & Chr(&H1B) & "L0101"
        sEditWK = sEditWK & Chr(&H1B) & "D103080" & "*/" & StrConv(TANAREC.Soko_No, vbUnicode) & _
                                                        StrConv(TANAREC.Retu, vbUnicode) & _
                                                        StrConv(TANAREC.Ren, vbUnicode) & _
                                                        StrConv(TANAREC.Dan, vbUnicode) & "*"
    
'2007.03.02        sEditWK = sEditWK & Chr(&H1B) & "H0350" & Chr(&H1B) & "V0170" & Chr(&H1B) & "L0101"
        sEditWK = sEditWK & Chr(&H1B) & "H0300" & Chr(&H1B) & "V0170" & Chr(&H1B) & "L0101"
        sEditWK = sEditWK & Chr(&H1B) & "X21," & "*/" & StrConv(TANAREC.Soko_No, vbUnicode) & _
                                                        StrConv(TANAREC.Retu, vbUnicode) & _
                                                        StrConv(TANAREC.Ren, vbUnicode) & _
                                                        StrConv(TANAREC.Dan, vbUnicode) & "*"
    
    
    
    
'       �w�薇��
        sEditWK = sEditWK & Chr(&H1B) & "Q1"
'       ���2007.03.22
        sEditWK = sEditWK & Chr(&H1B) & "CT0"

    
'       �ް����M�I���w��
        sEditWK = sEditWK & Chr(&H1B) & "Z"

'       ETX�w��
        sEditWK = sEditWK & Chr(&H3)
    
'       �ް����M
        PrinterDriver_Write lPrinterHandl, sEditWK
    
    
    
    
    
    
        com = BtOpGetNext
    
    Loop




    '����I������
    PrinterDriver_End lPrinterHandl








    Call Input_UnLock
    
    Print_Proc = False


End Function

Private Sub Command_Click(Index As Integer)

Dim sts         As Integer
Dim i           As Integer
Dim Tana_Cnt    As Long
Dim Yn          As Integer
    
    
    
    Select Case Index
        
        
        Case 8                              '�u���s�v
            

            For i = ptxS_Soko_No To ptxE_Dan
            
            
                If IsNumeric(Text(i).Text) Then
                    Text(i).Text = Format(CInt(Text(i).Text), "00")
                End If
            
            
                Select Case i
                
                    Case ptxE_Soko_No To ptxE_Dan
                        If Trim(Text(i).Text) = "" Then
                            Text(i).Text = "zz"
                        End If
                
                End Select
            
            
            
            Next i
            
            
            
            
            
            
            If Text(ptxS_Soko_No).Text & Text(ptxS_Retu).Text & Text(ptxS_Ren).Text & Text(ptxS_Dan).Text > _
                Text(ptxE_Soko_No).Text & Text(ptxE_Retu).Text & Text(ptxE_Ren).Text & Text(ptxE_Dan).Text Then
        
            End If
        
        
            Tana_Cnt = Print_Cnt_Proc()
            If Tana_Cnt = True Then
                Unload Me
            End If
        
            Yn = MsgBox("�I���x���́u" & StrConv(Format(Tana_Cnt, "#,##0"), vbWide) & "�v�����s����܂��B�X�����ł����H", vbYesNo, "�m�F����")
        
            If Yn = vbYes Then
                
                CommonDialog1.CancelError = True
                On Error GoTo ErrHandler
                
                CommonDialog1.ShowPrinter
        
        
                If Print_Proc() Then
                    Unload Me
                End If
        
        
        
            End If
        
        
        Case 11                             '�u�I���v
            Unload Me
        Case Else
            Beep
    End Select
    
    Exit Sub
    
ErrHandler:
    
    
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
    

    Show
                                '���O�t�@�C������荞��
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        End
    End If
    
                                '�I�}�X�^�n�o�d�m
    If TANA_Open(BtOpenNomal) Then
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�����𒆎~���ĉ������B"
        Unload Me
    End If
                                
    For Each Pri_Name In Printers
        If Pri_Name.DeviceName = Printer.DeviceName Then
            Exit For
        End If
    Next
                                
                                
                                
                                '��ʏ����ݒ�
    Call Clear_Field
    
    Text(ptxS_Soko_No).SetFocus
    
    End Sub

Private Sub Form_Unload(CANCEL As Integer)

Dim sts         As Integer
Dim Wk_Printer  As Printer
                                            
    For Each Wk_Printer In Printers
        If RTrim(Wk_Printer.DeviceName) = RTrim(Pri_Name.DeviceName) Then
            SetWindowsDefaultPrinter Wk_Printer.DeviceName, Wk_Printer.DriverName, Wk_Printer.Port
            Exit For
        End If
    Next
                                            
                                            '�I�}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�I�}�X�^")
            Beep
            MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
        End If
    End If
    
    sts = BTRV(BtOpReset, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "�I�}�X�^")
        Beep
        MsgBox "�V�X�e���ُ킪�������܂����B�������I�����ĉ������B", vbOKOnly
    End If

    Set F1060451 = Nothing


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

    
    If IsNumeric(Text(Index).Text) Then
        Text(Index).Text = Format(CInt(Text(Index).Text), "00")
    End If


    Select Case Index
    
        Case ptxE_Soko_No To ptxE_Dan
            If Trim(Text(Index).Text) = "" Then
                Text(Index).Text = "zz"
            End If
    
    End Select
    
    
    
        
    For i = Index + 1 To Text_Max
        If Text(i).TabStop Then
            Text(i).SetFocus
            Exit For
        End If
    Next i
End Sub


Private Function Print_Cnt_Proc() As Long
'��������̃J�E���g
Dim sts         As Integer
Dim com         As Integer

Dim Tana_Cnt    As Long


    Print_Cnt_Proc = True

    Tana_Cnt = 0

    Call UniCode_Conv(K0_TANA.Soko_No, Text(ptxS_Soko_No).Text)
    Call UniCode_Conv(K0_TANA.Retu, Text(ptxS_Retu).Text)
    Call UniCode_Conv(K0_TANA.Ren, Text(ptxS_Ren).Text)
    Call UniCode_Conv(K0_TANA.Dan, Text(ptxS_Dan).Text)

    com = BtOpGetGreaterEqual


    Do
    
        sts = BTRV(com, TANA_POS, TANAREC, Len(TANAREC), K0_TANA, Len(K0_TANA), 0)
        Select Case sts
            Case BtNoErr
            
                If (StrConv(TANAREC.Soko_No, vbUnicode) & StrConv(TANAREC.Retu, vbUnicode) & StrConv(TANAREC.Ren, vbUnicode) & StrConv(TANAREC.Dan, vbUnicode)) > _
                    (Text(ptxE_Soko_No).Text & Text(ptxE_Retu).Text & Text(ptxE_Ren).Text & Text(ptxE_Dan).Text) Then
                    Exit Do
                End If
            
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, BtOpGetEqual, "�I�}�X�^")
                Exit Function
        End Select
    
        Tana_Cnt = Tana_Cnt + 1
    
        com = BtOpGetNext
    
    Loop

    Print_Cnt_Proc = Tana_Cnt

End Function

Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�i�C�x���g�擾�s�j
'----------------------------------------------------------------------------

    F1060451.MousePointer = vbHourglass

    Call Ctrl_Lock(F1060451)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1060451)


    F1060451.MousePointer = vbDefault

End Sub


Private Function isWindowsNT() As Boolean
  isWindowsNT = IIf(GetVersion() And &H80000000, False, True)
End Function
Private Sub SetWindowsDefaultPrinter(ByVal DeviceName As String, ByVal DriverName As String, ByVal Port As String)
  Dim param As String
  param = DeviceName & "," & DriverName & "," & Port
  WriteProfileString "windows", "device", param
  If isWindowsNT Then
    'Windows NT/2000
    SendMessage HWND_BROADCAST, WM_WININICHANGE, 0&, ByVal 0&
  Else
    'Windows 95/98/Me
    SendMessage HWND_BROADCAST, WM_WININICHANGE, 0&, ByVal "windows"
  End If
'  Printer.EndDoc
End Sub

