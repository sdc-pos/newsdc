VERSION 5.00
Begin VB.Form F2010301 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�i�ԕʍ݌Ƀf�[�^�o��"
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
      Left            =   6120
      MaxLength       =   13
      TabIndex        =   14
      Top             =   1800
      Width           =   1695
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   3840
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   12
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  '�̌Œ�
      Index           =   0
      Left            =   3840
      MaxLength       =   13
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
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
      TabIndex        =   18
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�`"
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   17
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����O"
      Height          =   255
      Index           =   33
      Left            =   2160
      TabIndex        =   16
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�i�ԁi�O���j"
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   15
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Menu MainMenu 
      Caption         =   "���ƕ�"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F2010301"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ptxS_HIN_GAI% = 0             '�J�n�@�i��
Private Const ptxE_HIN_GAI% = 1             '�I���@�i��

Private Const Text_Max% = 1                 '��ʍ��ڕʍő���ޯ��

Private Const pcmbNAIGAI% = 0               '�����O

Dim HINZAI_DATA As String                   '�i�ԕʍ݌Ƀf�[�^�t���p�X
'Private Const Last_Update_Day$ = "[F201030]2015.08.20 14:30"
Private Const Last_Update_Day$ = "[F201030]2019.11.06 10:00 �i��trim�Ή�"



Private Function OUTPUT_Proc(Mode As Integer) As Integer
    
Dim sts                     As Integer
Dim ZAIKO_com               As Integer
Dim ITEM_com                As Integer

Dim SUMI_ALL_ZAIKO_QTY      As Long
Dim MI_ALL_ZAIKO_QTY        As Long
Dim ALL_ZAIKO_QTY           As Long

Dim SUMI_ST_ZAIKO_QTY       As Long
Dim MI_ST_ZAIKO_QTY         As Long
Dim ST_ZAIKO_QTY            As Long


Dim LOC_ZAIKO_QTY           As Long

Dim SAVE_LOC                As String * 8

Dim Ret                     As Integer
Dim FileNo                  As Integer
Dim FileName                As String

Dim c                       As String * 128
Dim Soko_No                 As String * 2

    OUTPUT_Proc = True
'���s���̓C�x���g�擾�s��
    Call Input_Lock         '��ʍ��ڃ��b�N

    FileNo = FreeFile
    FileName = HINZAI_DATA
    
    Ret = InStr(1, Trim(FileName), ".") - 1
    FileName = Left(Trim(FileName), Ret) & "_" & Last_JGYOBU & Right(Trim(FileName), Len(Trim(FileName)) - Ret)
    
    On Error GoTo Error_Proc
    Open (FileName) For Output As FileNo

    Write #FileNo, _
    "�i�ԁi�O���j", _
    "�i��", _
    "���݌�", _
    "�W���I��", _
    "�W���I��" & vbCrLf & "�݌�", _
    "�ʒu" & vbCrLf & "�I��1", "�ʒu" & vbCrLf & "�݌�1", "�ʒu" & vbCrLf & "�I��2", "�ʒu" & vbCrLf & "�݌�2", "�ʒu" & vbCrLf & "�I��3", "�ʒu" & vbCrLf & "�݌�3", "�ʒu" & vbCrLf & "�I��4", "�ʒu" & vbCrLf & "�݌�4", _
    "�ʒu" & vbCrLf & "�I��5", "�ʒu" & vbCrLf & "�݌�5", "�ʒu" & vbCrLf & "�I��6", "�ʒu" & vbCrLf & "�݌�6", "�ʒu" & vbCrLf & "�I��7", "�ʒu" & vbCrLf & "�݌�7", "�ʒu" & vbCrLf & "�I��8", "�ʒu" & vbCrLf & "�݌�8", _
    "�ʒu" & vbCrLf & "�I��9", "�ʒu" & vbCrLf & "�݌�9", "�ʒu" & vbCrLf & "�I��10", "�ʒu" & vbCrLf & "�݌�10", "�ʒu" & vbCrLf & "�I��11", "�ʒu" & vbCrLf & "�݌�11", "�ʒu" & vbCrLf & "�I��12", "�ʒu" & vbCrLf & "�݌�12", _
    "�ʒu" & vbCrLf & "�I��13", "�ʒu" & vbCrLf & "�݌�13", "�ʒu" & vbCrLf & "�I��14", "�ʒu" & vbCrLf & "�݌�14", "�ʒu" & vbCrLf & "�I��15", "�ʒu" & vbCrLf & "�݌�15", "�ʒu" & vbCrLf & "�I��16", "�ʒu" & vbCrLf & "�݌�16", _
    "�ʒu" & vbCrLf & "�I��17", "�ʒu" & vbCrLf & "�݌�17", "�ʒu" & vbCrLf & "�I��18", "�ʒu" & vbCrLf & "�݌�18", "�ʒu" & vbCrLf & "�I��19", "�ʒu" & vbCrLf & "�݌�19", "�ʒu" & vbCrLf & "�I��20", "�ʒu" & vbCrLf & "�݌�20"
    

    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, Right(Combo(pcmbNAIGAI).Text, 1))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, Text(ptxS_HIN_GAI).Text)

    ITEM_com = BtOpGetGreaterEqual

    Do
        DoEvents
        
        sts = BTRV(ITEM_com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)

        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) <> Last_JGYOBU Or _
                    StrConv(ITEMREC.NAIGAI, vbUnicode) <> Right(Combo(pcmbNAIGAI).Text, 1) Or _
                    RTrim(StrConv(ITEMREC.HIN_GAI, vbUnicode)) > Text(ptxE_HIN_GAI).Text Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, ITEM_com, "�i�ڃ}�X�^")
                Exit Function
        End Select

        
        If Zaiko_Syukei_Proc(SUMI_ALL_ZAIKO_QTY, MI_ALL_ZAIKO_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode)) Then
            Exit Function
        End If

        ALL_ZAIKO_QTY = SUMI_ALL_ZAIKO_QTY + MI_ALL_ZAIKO_QTY

        If Mode = 1 And ALL_ZAIKO_QTY = 0 Then
        Else

            Write #FileNo, Trim(StrConv(ITEMREC.HIN_GAI, vbUnicode)),   '2019/11/06 trim�Ή�
            Write #FileNo, Trim(StrConv(ITEMREC.HIN_NAME, vbUnicode)),  '2019/11/06 trim�Ή�

            Write #FileNo, Format(ALL_ZAIKO_QTY, "#0"),

            SAVE_LOC = ""


            If ALL_ZAIKO_QTY = 0 Then
            Else
                                                    '�W���I�ԕ�
                If Len(Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode))) = 0 Then
                    ST_ZAIKO_QTY = 0
                    Write #FileNo, , ,
                Else
                    If Zaiko_Syukei_Proc(SUMI_ST_ZAIKO_QTY, MI_ST_ZAIKO_QTY, StrConv(ITEMREC.JGYOBU, vbUnicode), StrConv(ITEMREC.NAIGAI, vbUnicode), StrConv(ITEMREC.HIN_GAI, vbUnicode), StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                        Exit Function
                    End If
            
                    ST_ZAIKO_QTY = SUMI_ST_ZAIKO_QTY + MI_ST_ZAIKO_QTY
            
                    If Len(Trim(StrConv(ITEMREC.ST_SOKO, vbUnicode))) = 0 Then
                        Write #FileNo, ,
                    Else
                        If GetIni("SOKO_NO", StrConv(ITEMREC.ST_SOKO, vbUnicode), "SYS", c) Then
                            Soko_No = StrConv(ITEMREC.ST_SOKO, vbUnicode)
                        Else
                            Soko_No = Trim(c)
                        End If
                        Write #FileNo, Soko_No & "-" & StrConv(ITEMREC.ST_RETU, vbUnicode) & "-" & StrConv(ITEMREC.ST_REN, vbUnicode) & "-" & StrConv(ITEMREC.ST_DAN, vbUnicode),
                    End If
                    Write #FileNo, Format(ST_ZAIKO_QTY, "#0"),
                End If
            
                If ALL_ZAIKO_QTY = ST_ZAIKO_QTY Then
                Else
            
                    Call UniCode_Conv(K4_ZAIKO.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                    Call UniCode_Conv(K4_ZAIKO.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                    Call UniCode_Conv(K4_ZAIKO.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                    Call UniCode_Conv(K4_ZAIKO.Soko_No, "")
                    Call UniCode_Conv(K4_ZAIKO.Retu, "")
                    Call UniCode_Conv(K4_ZAIKO.Ren, "")
                    Call UniCode_Conv(K4_ZAIKO.Dan, "")
    
                    LOC_ZAIKO_QTY = 0
                
                    ZAIKO_com = BtOpGetGreater

                    Do
    
                        sts = BTRV(ZAIKO_com, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K4_ZAIKO, Len(K4_ZAIKO), 4)

                        Select Case sts
                            Case BtNoErr
                                If StrConv(ZAIKOREC.JGYOBU, vbUnicode) <> StrConv(ITEMREC.JGYOBU, vbUnicode) Or _
                                    StrConv(ZAIKOREC.NAIGAI, vbUnicode) <> StrConv(ITEMREC.NAIGAI, vbUnicode) Or _
                                    StrConv(ZAIKOREC.HIN_GAI, vbUnicode) <> StrConv(ITEMREC.HIN_GAI, vbUnicode) Then
                                    Exit Do
                                End If
                            Case BtErrEOF
                                Exit Do
                            Case Else
                                Call File_Error(sts, ZAIKO_com, "�݌Ƀf�[�^")
                                Exit Function
                        End Select

                        If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)) = _
                            (StrConv(ITEMREC.ST_SOKO, vbUnicode) & StrConv(ITEMREC.ST_RETU, vbUnicode) & StrConv(ITEMREC.ST_REN, vbUnicode) & StrConv(ITEMREC.ST_DAN, vbUnicode)) Then
                
                            ZAIKO_com = BtOpGetGreater

                    
                        Else
                            If Len(Trim(SAVE_LOC)) = 0 Then
                                SAVE_LOC = StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)

                            End If
                        
                            If (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode)) <> _
                                SAVE_LOC Then
                                
                                If GetIni("SOKO_NO", Left(SAVE_LOC, 2), "SYS", c) Then
                                    Soko_No = Left(SAVE_LOC, 2)
                                Else
                                    Soko_No = Trim(c)
                                End If
                                
                                
                                Write #FileNo, Soko_No & "-" & Mid(SAVE_LOC, 3, 2) & "-" & Mid(SAVE_LOC, 5, 2) & "-" & Right(SAVE_LOC, 2),
                                Write #FileNo, Format(LOC_ZAIKO_QTY, "#0"),
                            
                                SAVE_LOC = (StrConv(ZAIKOREC.Soko_No, vbUnicode) & StrConv(ZAIKOREC.Retu, vbUnicode) & StrConv(ZAIKOREC.Ren, vbUnicode) & StrConv(ZAIKOREC.Dan, vbUnicode))
                                LOC_ZAIKO_QTY = 0
                            End If
                    
                    
                            LOC_ZAIKO_QTY = LOC_ZAIKO_QTY + CLng(StrConv(ZAIKOREC.YUKO_Z_QTY, vbUnicode))
                        
                            ZAIKO_com = BtOpGetNext
                        End If
                    
                
                    Loop
                
                    If Len(Trim(SAVE_LOC)) = 0 Then
                    Else
                        If GetIni("SOKO_NO", Left(SAVE_LOC, 2), "SYS", c) Then
                            Soko_No = Left(SAVE_LOC, 2)
                        Else
                            Soko_No = Trim(c)
                        End If
                        Write #FileNo, Soko_No & "-" & Mid(SAVE_LOC, 3, 2) & "-" & Mid(SAVE_LOC, 5, 2) & "-" & Right(SAVE_LOC, 2),
                        Write #FileNo, Format(LOC_ZAIKO_QTY, "#0"),
            
                    End If
                End If
            End If
        
            Write #FileNo,
        
        
        End If
        
        ITEM_com = BtOpGetNext
    
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

    F2010301.MousePointer = vbHourglass

    Call Ctrl_Lock(F2010301)


End Sub
Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   ��ʍ��ڃ��b�N�����i�C�x���g�擾�j
'----------------------------------------------------------------------------

    Call Ctrl_UnLock(F2010301)


    F2010301.MousePointer = vbDefault

End Sub

                                            '�G���[�`�F�b�N
Private Function Err_Chk() As Integer
                                            
                                            
                                            
    Err_Chk = True

'�i��(�O��)
    If Len(Text(ptxE_HIN_GAI).Text) = 0 Then
        Text(ptxE_HIN_GAI).Text = String(Len(StrConv(ZAIKOREC.HIN_GAI, vbUnicode)), "z")
    End If

    If Text(ptxS_HIN_GAI).Text > Text(ptxE_HIN_GAI).Text Then
        Beep
        MsgBox "���͂������ڂ̓G���[�ł��B"
        Text(ptxS_HIN_GAI).SetFocus
        Exit Function
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
            Text(ptxS_HIN_GAI).SetFocus
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
            ans = MsgBox("�u�i�ԕʍ݌Ƀf�[�^�v�o�͂��܂����H", vbYesNo + vbQuestion, "�m�F����")
                
            If ans = vbYes Then
                Beep
                ans = MsgBox("�݌ɂȂ��̕i�Ԃ��o�͂��܂����H", vbYesNo + vbQuestion + vbDefaultButton2, "�m�F����")
                    
                If ans = vbYes Then
                    If OUTPUT_Proc(0) Then
                        Unload Me
                    End If
                Else
                    If OUTPUT_Proc(1) Then
                        Unload Me
                    End If
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
Dim i As Integer
Dim c As String * 128
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
        MsgBox "���O�t�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    LOG_F = Trim(c)
                                '�݌Ƀt�@�C������荞��
    If GetIni("FILE", "HINZAI_DATA", "SYS", c) Then
        Beep
        MsgBox "�i�ԕʍ݌Ƀt�@�C�����̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If
    HINZAI_DATA = Trim(c)
                                '���ƕ���荞��
    If JGYOB_TB_Set Then
        Beep
        MsgBox "���ƕ��̊l���Ɏ��s���܂����B�����𒆎~���ĉ������B"
        End
    End If

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.08.20
'    For i = 0 To UBound(JGYOBU_T) - 1
'        If JGYOBU_T(i).CODE = " " Then
'            Unload SubMenu(i)
'            Exit For
'        End If
'
'        Load SubMenu(i + 1)
'        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)
'
'        If JGYOBU_T(i).CODE = Last_JGYOBU Then
'            F2010301.Caption = "�i�ԕʍ݌Ƀf�[�^�o�́i" + RTrim(JGYOBU_T(i).NAME) + ")"
'            SubMenu(i).Checked = True
'            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
'            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
''            LabJIGYO.BorderStyle = 1
'        Else
'            SubMenu(i).Checked = False
'        End If
'    Next i


    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Unload SubMenu(i)
            Exit For
        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F2010301.Caption = "�i�ԕʍ݌Ƀf�[�^�o�́i" + RTrim(JGYOBU_T(i).NAME) + ")" & Last_Update_Day

            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>    2015.08.20


                                '�i�ڃ}�X�^�n�o�d�m
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '�݌Ƀf�[�^�n�o�d�m
    If ZAIKO_Open(BtOpenNomal) Then
        Unload Me
    End If
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
                                            '�݌Ƀf�[�^�b�k�n�r�d
    sts = BTRV(BtOpClose, ZAIKO_POS, ZAIKOREC, Len(ZAIKOREC), K0_ZAIKO, Len(K0_ZAIKO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "�݌Ƀf�[�^")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F2010301 = Nothing

    End
End Sub

Private Sub SubMenu_Click(Index As Integer)

'>>>>>>>>>>>>>>>>>>>>>> 2015.08.20
'Dim i As Integer
'                                    '���j���[���I���v��
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If
'
'    For i = 0 To UBound(JGYOBU_T) - 1
'        If JGYOBU_T(i).CODE = " " Then
'            Exit For
'        End If
'        SubMenu(i).Checked = False
'    Next i
'                                    '���ƕ��؂�ւ�
'    F2010301.Caption = "�i�ԕʍ݌Ƀf�[�^�o�́i" + RTrim(JGYOBU_T(Index).NAME) + ")"
'    Last_JGYOBU = JGYOBU_T(Index).CODE
'    SubMenu(Index).Checked = True
'
'    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
'    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)



Dim i As Integer
                                    '���j���[���I���v��
    If JGYOBU_T(Index).CODE = " " Then
        Unload Me
    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '���ƕ��؂�ւ�
    F2010301.Caption = "�i�ԕʍ݌Ɉꗗ�\����i" + RTrim(JGYOBU_T(Index).NAME) + ")" & Last_Update_Day
    Last_JGYOBU = JGYOBU_T(Index).CODE
    SubMenu(Index).Checked = True

    LabJIGYO.Caption = RTrim(JGYOBU_T(Index).NAME)
    LabJIGYO.ForeColor = QBColor(JGYOBU_T(Index).COLOR)
'>>>>>>>>>>>>>>>>>>>>>> 2015.08.20

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


