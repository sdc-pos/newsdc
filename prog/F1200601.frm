VERSION 5.00
Begin VB.Form F1200601 
   BackColor       =   &H00FFFFFF&
   Caption         =   "���P�ʏo�א��W�v�f�[�^�쐬"
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
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   1
      Left            =   6240
      MaxLength       =   2
      TabIndex        =   2
      Top             =   2520
      Width           =   375
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   5400
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   5400
      MaxLength       =   4
      TabIndex        =   1
      Top             =   2520
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
      TabIndex        =   14
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
      Index           =   9
      Left            =   8640
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
      Index           =   8
      Left            =   7800
      TabIndex        =   11
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
      Index           =   6
      Left            =   5640
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
      Index           =   5
      Left            =   4800
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
      Index           =   4
      Left            =   3960
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
      Index           =   3
      Left            =   2640
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
      Index           =   2
      Left            =   1800
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
      Index           =   1
      Left            =   960
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
      Index           =   0
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��"
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   19
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�N"
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   18
      Top             =   2640
      Width           =   375
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
      TabIndex        =   17
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   33
      Left            =   3720
      TabIndex        =   16
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�w��N��"
      Height          =   255
      Index           =   0
      Left            =   3720
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1200601"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxYY% = 0                    '�w��N���@�N
Private Const ptxMM% = 1                    '�w��N���@��

Private Const Text_Max% = 1                 '��ʍ��ڕʍő���ޯ��

Private Const pcmbNAIGAI% = 0               '�����O

Dim ShipmentCnt_DATA    As String           '���P�ʏo�׏W�v�f�[�^�t���p�X
Private Function OUTPUT_Proc() As Integer
'----------------------------------------------------------------------------
'                  �b�r�u�f�[�^�o�͏���
'----------------------------------------------------------------------------
    
Dim sts             As Integer
Dim com             As Integer
Dim com_IDO         As Integer
Dim Ret             As Integer
    

Dim FileNo          As Integer
Dim FileName        As String


Dim Shipment_Cnt    As Long

Dim c               As String * 128
Dim Soko_No         As String * 2


    OUTPUT_Proc = True
'���s���̓C�x���g�擾�s��
    Call Input_Lock         '��ʍ��ڃ��b�N



    FileNo = FreeFile
    FileName = ShipmentCnt_DATA
    
    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & "_" & Last_JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)
    
    On Error GoTo Error_Proc
    
    Open (FileName) For Output As FileNo


    Write #FileNo, "�o�א��W�v�f�[�^(" & Text(ptxYY).Text & "�N" & Text(ptxMM).Text & "����)"
    Write #FileNo, "�i�ԁi�O���j", "�W���I��", "�o�׉�"

    
    com = BtOpGetFirst

    Do
        DoEvents
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K6_ITEM, Len(K6_ITEM), 6)

        Select Case sts
            Case BtNoErr
            
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "�i�ڃ}�X�^")
                Exit Function
        End Select
        '�݌Ɉړ������o�׉񐔃J�E���g
        Call UniCode_Conv(K1_IDO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K1_IDO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K1_IDO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
        Call UniCode_Conv(K1_IDO.JITU_DT, Text(ptxYY).Text & Text(ptxMM).Text & "01")
        Call UniCode_Conv(K1_IDO.JITU_TM, "")
                                                
        Shipment_Cnt = 0
                                                
        com_IDO = BtOpGetGreater
                                                
        Do
            DoEvents
            
            sts = BTRV(com_IDO, IDO_POS, IDOREC, Len(IDOREC), K1_IDO, Len(K1_IDO), 1)

            Select Case sts
                Case BtNoErr
            
                    If StrConv(IDOREC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                        StrConv(IDOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                        Trim(StrConv(IDOREC.HIN_GAI, vbUnicode)) <> Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
                        Exit Do
                    End If
            
                    If StrConv(IDOREC.JITU_DT, vbUnicode) > (Text(ptxYY).Text & Text(ptxMM).Text & "31") Then
                        Exit Do
                    End If
                Case BtErrEOF
                    Exit Do
                Case Else
                    Call File_Error(sts, com_IDO, "�݌Ɉړ���")
                    Exit Function
            End Select
        
            Select Case Left(StrConv(IDOREC.RIRK_ID, vbUnicode), 1)
                '�o�׏����Ȃ�񐔃J�E���g
                Case ACT_SYUKA_KEI, ACT_SYUKA_HYO, ACT_SYUKA_GAI
                    
                    Shipment_Cnt = Shipment_Cnt + 1
            End Select
            
            
            com_IDO = BtOpGetNext
        
        Loop
                                                
                                                
                                                
                                                '�i�ԁi�O���j
        Write #FileNo, StrConv(ITEMREC.HIN_GAI, vbUnicode),
                                                '�W���I��
        If GetIni("SOKO_NO", StrConv(ITEMREC.ST_SOKO, vbUnicode), "SYS", c) Then
            Soko_No = StrConv(ITEMREC.ST_SOKO, vbUnicode)
        Else
            Soko_No = Trim(c)
        End If
        
        
        Write #FileNo, Soko_No & "-" & _
                        StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & _
                        StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & _
                        StrConv(ITEMREC.ST_DAN, vbUnicode),
                                                '�o�׉�
        Write #FileNo, Format(Shipment_Cnt, "#0")



        com = BtOpGetNext
    Loop

    Close #FileNo
    
    Call Input_UnLock         '��ʍ��ڃ��b�N����
    Beep
    MsgBox "�u" & FileName & "�v�͐���ɏo�͂���܂����B"

    OUTPUT_Proc = False
    
    Exit Function


Error_Proc:

    If Err.Number = 70 Then
        Beep
        MsgBox FileName & "���g�p���ł��B"
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

    F1200601.MousePointer = vbHourglass

    Call Ctrl_Lock(F1200601)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F1200601)


    F1200601.MousePointer = vbDefault

End Sub

Private Function Err_Chk() As Integer
'----------------------------------------------------------------------------
'                   ���̓G���[�`�F�b�N����
'----------------------------------------------------------------------------
Dim i   As Integer
    
    Err_Chk = True

    
    If Not IsNumeric(Text(ptxYY).Text) Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxYY).SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(Text(ptxMM).Text) Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxMM).SetFocus
        Exit Function
    Else
        Text(ptxMM).Text = Format(CInt(Text(ptxMM).Text), "00")
        
        If Text(ptxMM).Text < "01" Or Text(ptxMM).Text > "12" Then
            Beep
            MsgBox "���͂������ڂ̓G���[�ł��B"
            Text(ptxMM).SetFocus
            Exit Function
        End If
    
    End If
    
    Err_Chk = False

End Function

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   �R���{�{�b�N�X���́i�j�����c�������j����
'----------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
        
    Select Case Index
        Case pcmbNAIGAI        '�����敪
            Text(ptxYY).SetFocus
    End Select

End Sub

Private Sub Command_Click(Index As Integer)
Dim ans As Integer
        
    Select Case Index
        Case 7                              '�f�[�^�o��
            If Err_Chk() Then
                Exit Sub
            End If
        
            Beep
            ans = MsgBox("�u���P�ʏo�א��W�v�f�[�^�v�f�[�^�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
            
            If ans = vbYes Then
                
                
                If OUTPUT_Proc() Then
                    Unload Me
                End If
            End If
            Combo(pcmbNAIGAI).SetFocus
        Case 11                             '�I��
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
                                '���P�ʏo�א��W�v�f�[�^�t�@�C������荞��
    If GetIni("FILE", "ShipmentCnt_DATA", "SYS", c) Then
        Beep
        MsgBox "���P�ʏo�א��W�v�f�[�^�쐬�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    ShipmentCnt_DATA = Trim(c)
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).Code = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).Code = Last_JGYOBU Then
            F1200601.Caption = "���P�ʏo�א��W�v�f�[�^�쐬�i" + RTrim(JGYOBU_T(i).NAME) + ")"
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
                                
                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ɉړ����n�o�d�m
    If IDO_Open(BtOpenNomal) Then
        Unload Me
    End If
    
    Show
                                
                                '��ʏ����ݒ�
    Combo(pcmbNAIGAI).AddItem NAIGAI1 & "   " & NAIGAI_NAI
    Combo(pcmbNAIGAI).AddItem NAIGAI2 & "   " & NAIGAI_GAI
    Combo(pcmbNAIGAI).ListIndex = 0
    
    Combo(pcmbNAIGAI).SetFocus
    
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
                                            '�i�ڃ}�X�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�i�ڃ}�X�^")
        End If
    End If
                                            '�݌Ɉړ����b�k�n�r�d
    sts = BTRV(BtOpClose, IDO_POS, IDOREC, Len(IDOREC), K0_IDO, Len(K0_IDO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ɉړ���")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1200601 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    '���j���[���I���v��
    If JGYOBU_T(Index).Code = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T) - 1
        If JGYOBU_T(i).Code = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F1200601.Caption = "���P�ʏo�א��W�v�f�[�^�쐬�i" + RTrim(JGYOBU_T(Index).NAME) + ")"
    Last_JGYOBU = JGYOBU_T(Index).Code
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
                
        
        
    For i = Index + 1 To Text_Max
        If Text(i).Enabled And Text(i).Visible And Text(i).TabStop Then
                Text(i).SetFocus
                Exit For
        End If
    Next i

End Sub
