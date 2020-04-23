VERSION 5.00
Begin VB.Form ITEM_DATA_SET1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "品目マスタ復旧処理（ITEM_DATA_SET 2010.08.15 12:30)"
   ClientHeight    =   8550
   ClientLeft      =   2325
   ClientTop       =   2625
   ClientWidth     =   10095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   10095
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   4725
      TabIndex        =   16
      Top             =   2040
      Width           =   2265
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   2250
      TabIndex        =   15
      Top             =   2040
      Width           =   2265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "チェック＆更新"
      Height          =   435
      Index           =   1
      Left            =   7380
      TabIndex        =   14
      Top             =   1920
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終了"
      Height          =   435
      Index           =   2
      Left            =   7290
      TabIndex        =   13
      Top             =   3480
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "チェック"
      Height          =   435
      Index           =   0
      Left            =   7290
      TabIndex        =   10
      Top             =   1080
      Width           =   1860
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   4725
      TabIndex        =   9
      Top             =   1620
      Width           =   2265
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   2250
      TabIndex        =   8
      Top             =   1620
      Width           =   2265
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   5940
      TabIndex        =   25
      Top             =   5760
      Width           =   1410
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   7
      Left            =   4455
      TabIndex        =   24
      Top             =   5760
      Width           =   1410
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   6
      Left            =   2925
      TabIndex        =   23
      Top             =   5760
      Width           =   1410
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   5
      Left            =   5940
      TabIndex        =   22
      Top             =   5280
      Width           =   1410
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   4
      Left            =   4455
      TabIndex        =   21
      Top             =   5280
      Width           =   1410
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   2925
      TabIndex        =   20
      Top             =   5280
      Width           =   1410
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   5940
      TabIndex        =   19
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "A：B"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   1665
      TabIndex        =   18
      Top             =   2040
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "A：C"
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   1665
      TabIndex        =   17
      Top             =   1620
      Width           =   555
   End
   Begin VB.Label Label3 
      Height          =   315
      Index           =   1
      Left            =   4095
      TabIndex        =   12
      Top             =   7260
      Width           =   2445
   End
   Begin VB.Label Label3 
      Height          =   315
      Index           =   0
      Left            =   1530
      TabIndex        =   11
      Top             =   7260
      Width           =   2445
   End
   Begin VB.Label Label2 
      Caption         =   "更新件数"
      Height          =   315
      Index           =   2
      Left            =   5940
      TabIndex        =   7
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "対象件数"
      Height          =   315
      Index           =   1
      Left            =   4455
      TabIndex        =   6
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label2 
      Caption         =   "読込み件数"
      Height          =   315
      Index           =   0
      Left            =   2925
      TabIndex        =   5
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   2925
      TabIndex        =   4
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "品目マスタ"
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   1350
      TabIndex        =   3
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Cnt 
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   4455
      TabIndex        =   2
      Top             =   4800
      Width           =   1410
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   3300
      Width           =   240
   End
   Begin VB.Label MsgLab 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "データ抽出処理"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   3360
   End
End
Attribute VB_Name = "ITEM_DATA_SET1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Function Update_Proc(Mode As Integer) As Integer

Dim sts             As Integer
Dim Upd_com         As Integer
Dim com             As Integer
Dim ans             As Integer
Dim Count           As Long
Dim sel_count       As Long
Dim upd_count       As Long

Dim DISP_INTERVAL   As Long


Dim c               As String * 128
Dim FullPath        As String



Dim Start_Now       As String







    Update_Proc = True




    Label3(0).Caption = Format(Now)
    Label3(1).Caption = ""

'---------------------------------------------  受入履歴データのコンバート
    MsgLab(1) = "品目マスタ抽出処理中！！"
    Me.MousePointer = vbHourglass
    Count = 0
    sel_count = 0
    upd_count = 0
    DISP_INTERVAL = 0
    Cnt(0).Caption = Format(Count, "#0")
    Cnt(1).Caption = Format(Count, "#0")
    Cnt(2).Caption = Format(Count, "#0")
                                        
    Start_Now = Format(Now, "YYYYMMDDHHMMSS")
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
    sts = BTRV(BtOpClose, ITEM_UP_POS, ITEM_UPREC, Len(ITEM_UPREC), K0_ITEM_UP, Len(K0_ITEM_UP), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目ﾏｽﾀ")
        End If
    End If
                
                
    sts = GetIni("FILE", ITEM_UP_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_UP]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)
    On Error Resume Next
    Kill (FullPath)
    On Error GoTo 0
                
                
    If ITEM_UP_Open(BtOpenNomal) Then
        Unload Me
    End If
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
    sts = BTRV(BtOpClose, ITEM_A_POS, ITEM_AREC, Len(ITEM_AREC), K0_ITEM_A, Len(K0_ITEM_A), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目ﾏｽﾀ")
        End If
    End If
                
                
    sts = GetIni("FILE", ITEM_A_ID, "SYS", c)
    If sts <> False Then
        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_A]読み込みエラー ")
        Exit Function
    End If

    FullPath = RTrim(c)
    On Error Resume Next
    Kill (FullPath)
    On Error GoTo 0
                
                
    If ITEM_A_Open(BtOpenNomal) Then
        Unload Me
    End If
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
                                        
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) = "S" Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        
        
        Count = Count + 1
        Cnt(0).Caption = Format(Count, "#0")
        
        
        If StrConv(ITEMREC.UPD_DATETIME, vbUnicode) >= Trim(Text1(0).Text) And _
            StrConv(ITEMREC.UPD_DATETIME, vbUnicode) <= Trim(Text1(1).Text) Then
                
                sel_count = sel_count + 1
                Cnt(1).Caption = Format(sel_count, "#0")
        
'            If StrConv(ITEMREC.L_PAPER, vbUnicode) <= "0" And _
'                StrConv(ITEMREC.L_PLASTIC, vbUnicode) <= "0" Then
        
'                sel_count = sel_count + 1
'                Cnt(1).Caption = Format(sel_count, "#0")
                    
                    
                        Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, StrConv(ITEMREC.L_PAPER, vbUnicode))
                        Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, StrConv(ITEMREC.L_PLASTIC, vbUnicode))
                    
                    
                        Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")
                        Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")
                    
                    
                        Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")
                        Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")
                        Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")
                        Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")
                    
                    
                        Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, StrConv(ITEMREC.L_PAPER, vbUnicode))
                        Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, StrConv(ITEMREC.L_PLASTIC, vbUnicode))
                    
                    
                    
                    
                    
                    
                    
                sts = BTRV(BtOpInsert, ITEM_A_POS, ITEMREC, Len(ITEMREC), K0_ITEM_A, Len(K0_ITEM_A), 0)
                If sts Then
                    Call File_Error(sts, BtOpInsert, "品目マスタ")
                    Exit Function
                End If
                
 '           End If
                
        End If
                
                
        com = BtOpGetNext
    Loop
                
    MsgLab(1) = "品目マスタマッチング処理中！！"
                
                
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, ITEM_A_POS, ITEM_AREC, Len(ITEM_AREC), K0_ITEM_A, Len(K0_ITEM_A), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        
        
        
        Call UniCode_Conv(K0_ITEM_C.JGYOBU, StrConv(ITEM_AREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM_C.NAIGAI, StrConv(ITEM_AREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM_C.HIN_GAI, StrConv(ITEM_AREC.HIN_GAI, vbUnicode))
        
        
If Trim(StrConv(ITEM_AREC.HIN_GAI, vbUnicode)) = "APB01H413-CU" Then
Debug.Print
End If
        
        
        sts = BTRV(BtOpGetEqual, ITEM_C_POS, ITEM_CREC, Len(ITEM_CREC), K0_ITEM_C, Len(K0_ITEM_C), 0)
        Select Case sts
            Case BtNoErr
            
            
                If (StrConv(ITEM_CREC.L_PAPER, vbUnicode) = "0" Or StrConv(ITEM_CREC.L_PAPER, vbUnicode) = "1") And _
                    (StrConv(ITEM_CREC.L_PLASTIC, vbUnicode) = "0" Or StrConv(ITEM_CREC.L_PLASTIC, vbUnicode) = "1") Then
            
            
            
                    If StrConv(ITEM_AREC.L_PAPER, vbUnicode) = "1" Or StrConv(ITEM_AREC.L_PLASTIC, vbUnicode) = "1" Then
            
                    Else
                        If StrConv(ITEM_CREC.L_PAPER, vbUnicode) <> StrConv(ITEM_AREC.L_PAPER, vbUnicode) Or _
                            StrConv(ITEM_CREC.L_PLASTIC, vbUnicode) <> StrConv(ITEM_AREC.L_PLASTIC, vbUnicode) Then
                        
                        
                            Call UniCode_Conv(ITEM_AREC.BEF_1_L_PAPER, StrConv(ITEM_AREC.L_PAPER, vbUnicode))
                            Call UniCode_Conv(ITEM_AREC.BEF_1_L_PLASTIC, StrConv(ITEM_AREC.L_PLASTIC, vbUnicode))
                        
                        
                            Call UniCode_Conv(ITEM_AREC.BEF_2_L_PAPER, StrConv(ITEM_CREC.L_PAPER, vbUnicode))
                            Call UniCode_Conv(ITEM_AREC.BEF_2_L_PLASTIC, StrConv(ITEM_CREC.L_PLASTIC, vbUnicode))
                        
                        
                            Call UniCode_Conv(ITEM_AREC.BEF_3_L_PAPER, "")
                            Call UniCode_Conv(ITEM_AREC.BEF_3_L_PLASTIC, "")
                            Call UniCode_Conv(ITEM_AREC.BEF_4_L_PAPER, "")
                            Call UniCode_Conv(ITEM_AREC.BEF_4_L_PLASTIC, "")
                        
                        
                            Call UniCode_Conv(ITEM_AREC.BEF_LAST_L_PAPER, StrConv(ITEM_CREC.L_PAPER, vbUnicode))
                            Call UniCode_Conv(ITEM_AREC.BEF_LAST_L_PLASTIC, StrConv(ITEM_CREC.L_PLASTIC, vbUnicode))
                        
                        
                            sts = BTRV(BtOpInsert, ITEM_UP_POS, ITEM_AREC, Len(ITEM_AREC), K0_ITEM_UP, Len(K0_ITEM_UP), 0)
                            If sts Then
                                Call File_Error(sts, BtOpInsert, "品目マスタ")
                                Exit Function
                            End If
                        
                        
                            upd_count = upd_count + 1
                            Cnt(2).Caption = Format(upd_count, "#0")
                                        
                        
                                    
                        
                        
                        
                        End If
                    End If
                End If
            
            
            
            
            
                Call UniCode_Conv(ITEM_AREC.BEF_2_L_PAPER, StrConv(ITEM_CREC.L_PAPER, vbUnicode))
                Call UniCode_Conv(ITEM_AREC.BEF_2_L_PLASTIC, StrConv(ITEM_CREC.L_PLASTIC, vbUnicode))
            
            
            
            
                sts = BTRV(BtOpUpdate, ITEM_A_POS, ITEM_AREC, Len(ITEM_AREC), K0_ITEM_A, Len(K0_ITEM_A), 0)
                If sts Then
                    Call File_Error(sts, BtOpInsert, "品目マスタ")
                    Exit Function
                End If
            
            
            
            
            
            Case BtErrKeyNotFound
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
        
        
        
        
                
                
        com = BtOpGetNext
    Loop
                
                
                
'    sts = BTRV(BtOpClose, ITEM_A_POS, ITEM_AREC, Len(ITEM_AREC), K0_ITEM_A, Len(K0_ITEM_A), 0)
'    If sts Then
'        If sts <> BtErrNoOpen Then
'            Call File_Error(sts, BtOpClose, "品目ﾏｽﾀ")
'        End If
'    End If


'    sts = GetIni("FILE", ITEM_A_ID, "SYS", c)
'    If sts <> False Then
'        Call LOG_OUT(LOG_F, "SYS.INI [ITEM_A]読み込みエラー ")
'        Exit Function
'    End If

'    FullPath = RTrim(c)
'    On Error Resume Next
'    Kill (FullPath)
'    On Error GoTo 0


'    If ITEM_A_Open(BtOpenNomal) Then
'        Unload Me
'    End If
                
                
                
                
    Count = 0
    sel_count = 0
    upd_count = 0
    DISP_INTERVAL = 0
    Cnt(3).Caption = Format(Count, "#0")
    Cnt(4).Caption = Format(Count, "#0")
    Cnt(5).Caption = Format(Count, "#0")
                
     MsgLab(1) = "品目マスタ抽出処理中！！"
               
                
                
                
'-------------------------------------------------------
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
        Select Case sts
            Case BtNoErr
                If StrConv(ITEMREC.JGYOBU, vbUnicode) = "S" Then
                    Exit Do
                End If
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        
        Count = Count + 1
        Cnt(3).Caption = Format(Count, "#0")
If Trim(StrConv(ITEM_AREC.HIN_GAI, vbUnicode)) = "APB01H413-CU" Then
Debug.Print
End If
        
        
        If StrConv(ITEMREC.UPD_DATETIME, vbUnicode) >= Trim(Text1(2).Text) And _
            StrConv(ITEMREC.UPD_DATETIME, vbUnicode) <= Trim(Text1(3).Text) Then
            
                sel_count = sel_count + 1
                Cnt(4).Caption = Format(sel_count, "#0")
            
            
'            If StrConv(ITEMREC.L_PAPER, vbUnicode) <= "0" And _
'                StrConv(ITEMREC.L_PLASTIC, vbUnicode) <= "0" Then
    
'                sel_count = sel_count + 1
'                Cnt(4).Caption = Format(sel_count, "#0")
                    
                    
                    
                Call UniCode_Conv(K0_ITEM_A.JGYOBU, StrConv(ITEMREC.JGYOBU, vbUnicode))
                Call UniCode_Conv(K0_ITEM_A.NAIGAI, StrConv(ITEMREC.NAIGAI, vbUnicode))
                Call UniCode_Conv(K0_ITEM_A.HIN_GAI, StrConv(ITEMREC.HIN_GAI, vbUnicode))
                
                sts = BTRV(BtOpGetEqual, ITEM_A_POS, ITEM_AREC, Len(ITEM_AREC), K0_ITEM_A, Len(K0_ITEM_A), 0)
                Select Case sts
                    
                    
                    Case BtNoErr
                            
                            
                            
                            
                            
                    Case BtErrKeyNotFound
                            
                            
                                Call UniCode_Conv(ITEMREC.BEF_1_L_PAPER, StrConv(ITEMREC.L_PAPER, vbUnicode))
                                Call UniCode_Conv(ITEMREC.BEF_1_L_PLASTIC, StrConv(ITEMREC.L_PLASTIC, vbUnicode))
                            
                            
                                Call UniCode_Conv(ITEMREC.BEF_2_L_PAPER, "")
                                Call UniCode_Conv(ITEMREC.BEF_2_L_PLASTIC, "")
                            
                            
                                Call UniCode_Conv(ITEMREC.BEF_3_L_PAPER, "")
                                Call UniCode_Conv(ITEMREC.BEF_3_L_PLASTIC, "")
                                Call UniCode_Conv(ITEMREC.BEF_4_L_PAPER, "")
                                Call UniCode_Conv(ITEMREC.BEF_4_L_PLASTIC, "")
                            
                            
                                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PAPER, StrConv(ITEMREC.L_PAPER, vbUnicode))
                                Call UniCode_Conv(ITEMREC.BEF_LAST_L_PLASTIC, StrConv(ITEMREC.L_PLASTIC, vbUnicode))
                            
                            
                            
                            
                            
                        sts = BTRV(BtOpInsert, ITEM_A_POS, ITEMREC, Len(ITEMREC), K0_ITEM_A, Len(K0_ITEM_A), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrDuplicates
                            
                            Case Else
                                Call File_Error(sts, BtOpInsert, "品目マスタ")
                                Exit Function
                        End Select
                    
                    
        End Select
            
'            End If
            
        End If
                
                
        com = BtOpGetNext
    Loop
                
    MsgLab(1) = "品目マスタマッチング処理中！！"
                
                
    com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, ITEM_A_POS, ITEM_AREC, Len(ITEM_AREC), K0_ITEM_A, Len(K0_ITEM_A), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        
If Trim(StrConv(ITEM_AREC.HIN_GAI, vbUnicode)) = "APB01H413-CU" Then
Debug.Print
End If
        
        
        Call UniCode_Conv(K0_ITEM_B.JGYOBU, StrConv(ITEM_AREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM_B.NAIGAI, StrConv(ITEM_AREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM_B.HIN_GAI, StrConv(ITEM_AREC.HIN_GAI, vbUnicode))
        
        sts = BTRV(BtOpGetEqual, ITEM_B_POS, ITEM_BREC, Len(ITEM_BREC), K0_ITEM_B, Len(K0_ITEM_B), 0)
        Select Case sts
            Case BtNoErr
            
            
                
                If (StrConv(ITEM_BREC.L_PAPER, vbUnicode) = "0" Or StrConv(ITEM_BREC.L_PAPER, vbUnicode) = "1") And _
                    (StrConv(ITEM_BREC.L_PLASTIC, vbUnicode) = "0" Or StrConv(ITEM_BREC.L_PLASTIC, vbUnicode) = "1") Then
                
                    If StrConv(ITEM_AREC.L_PAPER, vbUnicode) = "1" Or StrConv(ITEM_AREC.L_PLASTIC, vbUnicode) = "1" Then
            
                    Else
                        If StrConv(ITEM_BREC.L_PAPER, vbUnicode) <> StrConv(ITEM_AREC.L_PAPER, vbUnicode) Or _
                            StrConv(ITEM_BREC.L_PLASTIC, vbUnicode) <> StrConv(ITEM_AREC.L_PLASTIC, vbUnicode) Then
                        
                            Call UniCode_Conv(K0_ITEM_UP.JGYOBU, StrConv(ITEM_AREC.JGYOBU, vbUnicode))
                            Call UniCode_Conv(K0_ITEM_UP.NAIGAI, StrConv(ITEM_AREC.NAIGAI, vbUnicode))
                            Call UniCode_Conv(K0_ITEM_UP.HIN_GAI, StrConv(ITEM_AREC.HIN_GAI, vbUnicode))
                            
                            sts = BTRV(BtOpGetEqual, ITEM_UP_POS, ITEM_UPREC, Len(ITEM_UPREC), K0_ITEM_UP, Len(K0_ITEM_UP), 0)
                            Select Case sts
                                Case BtNoErr
                                
                                    Upd_com = BtOpUpdate
                                
                                Case BtErrKeyNotFound
                                    Upd_com = BtOpInsert
                                
                                Case Else
                                    Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                    Exit Function
                            End Select
                            
                            
                            
                            
                            
                            If Upd_com = BtOpInsert Then
                            
                            
                                Call UniCode_Conv(ITEM_AREC.BEF_1_L_PAPER, StrConv(ITEM_AREC.L_PAPER, vbUnicode))
                                Call UniCode_Conv(ITEM_AREC.BEF_1_L_PLASTIC, StrConv(ITEM_AREC.L_PLASTIC, vbUnicode))
                        
                        
                                Call UniCode_Conv(ITEM_AREC.BEF_2_L_PAPER, StrConv(ITEM_AREC.BEF_2_L_PAPER, vbUnicode))
                                Call UniCode_Conv(ITEM_AREC.BEF_2_L_PLASTIC, StrConv(ITEM_AREC.BEF_2_L_PLASTIC, vbUnicode))
                        
                        
                                Call UniCode_Conv(ITEM_AREC.BEF_4_L_PAPER, "")
                                Call UniCode_Conv(ITEM_AREC.BEF_4_L_PLASTIC, "")
                        
                        
                            Else
                                Call UniCode_Conv(ITEM_AREC.BEF_1_L_PAPER, StrConv(ITEM_UPREC.BEF_1_L_PAPER, vbUnicode))
                                Call UniCode_Conv(ITEM_AREC.BEF_1_L_PLASTIC, StrConv(ITEM_UPREC.BEF_1_L_PLASTIC, vbUnicode))
                        
                        
                                Call UniCode_Conv(ITEM_AREC.BEF_2_L_PAPER, StrConv(ITEM_UPREC.BEF_2_L_PAPER, vbUnicode))
                                Call UniCode_Conv(ITEM_AREC.BEF_2_L_PLASTIC, StrConv(ITEM_UPREC.BEF_2_L_PLASTIC, vbUnicode))
                            
                            
                            End If
                        
                        
                            Call UniCode_Conv(ITEM_AREC.BEF_3_L_PAPER, StrConv(ITEM_BREC.L_PAPER, vbUnicode))
                            Call UniCode_Conv(ITEM_AREC.BEF_3_L_PLASTIC, StrConv(ITEM_BREC.L_PLASTIC, vbUnicode))
                        
                        
                            Call UniCode_Conv(ITEM_AREC.BEF_LAST_L_PAPER, StrConv(ITEM_BREC.L_PAPER, vbUnicode))
                            Call UniCode_Conv(ITEM_AREC.BEF_LAST_L_PLASTIC, StrConv(ITEM_BREC.L_PLASTIC, vbUnicode))
                        
                        
                        
                            sts = BTRV(Upd_com, ITEM_UP_POS, ITEM_AREC, Len(ITEM_AREC), K0_ITEM_UP, Len(K0_ITEM_UP), 0)
                            If sts Then
                                Call File_Error(sts, BtOpInsert, "品目マスタ")
                                Exit Function
                            End If
                            
                            upd_count = upd_count + 1
                        
                        
                            Cnt(5).Caption = Format(upd_count, "#0")
                        
                        
                        
                        End If
                    
                    
                    
                    
                        
                
                    End If
                End If
            
            
            
                Call UniCode_Conv(ITEM_AREC.BEF_3_L_PAPER, StrConv(ITEM_BREC.L_PAPER, vbUnicode))
                Call UniCode_Conv(ITEM_AREC.BEF_3_L_PLASTIC, StrConv(ITEM_BREC.L_PLASTIC, vbUnicode))
            
If Trim(StrConv(ITEM_BREC.HIN_GAI, vbUnicode)) = "APB01H413-CU" Then
    Debug.Print
End If
            
            
            
            
            
            
            
                sts = BTRV(BtOpUpdate, ITEM_A_POS, ITEM_AREC, Len(ITEM_AREC), K0_ITEM_A, Len(K0_ITEM_A), 0)
                If sts Then
                    Call File_Error(sts, BtOpInsert, "品目マスタ")
                    Exit Function
                End If
            
            
            
            
            
            
            
            Case BtErrKeyNotFound
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
        
        
        
        
                
                
        com = BtOpGetNext
    Loop


    Count = 0
    sel_count = 0
    upd_count = 0

    Cnt(6).Caption = Format(Count, "#0")
    Cnt(7).Caption = Format(Count, "#0")
    Cnt(8).Caption = Format(Count, "#0")


    MsgLab(1) = "実行ログ出力処理中！！"

    Call LOG_OUT(Start_Now & "item.CSV", "," & "事業部" & "," & _
                        "内外" & "," & _
                        "品外" & "," & _
                        "現　紙" & "," & _
                        "現　プラ" & "," & _
                        "2010/7/23 紙" & "," & _
                        "2010/7/23 プラ" & "," & _
                        "2009/4/28 紙" & "," & _
                        "2009/4/28 プラ" & "," & _
                        "新　紙" & "," & _
                        "新　プラ" & "," & "更新" & "," & "更新担当" & "," & "更新日付")



    com = BtOpGetFirst
    Do
        
        DoEvents
        
        
        sts = BTRV(com, ITEM_A_POS, ITEM_AREC, Len(ITEM_AREC), K0_ITEM_A, Len(K0_ITEM_A), 0)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "品目マスタ")
                Exit Function
        End Select
        
        Count = Count + 1
        Cnt(6).Caption = Format(Count, "#0")
        
        
        Call UniCode_Conv(K0_ITEM_UP.JGYOBU, StrConv(ITEM_AREC.JGYOBU, vbUnicode))
        Call UniCode_Conv(K0_ITEM_UP.NAIGAI, StrConv(ITEM_AREC.NAIGAI, vbUnicode))
        Call UniCode_Conv(K0_ITEM_UP.HIN_GAI, StrConv(ITEM_AREC.HIN_GAI, vbUnicode))
        
        sts = BTRV(BtOpGetEqual, ITEM_UP_POS, ITEM_UPREC, Len(ITEM_UPREC), K0_ITEM_UP, Len(K0_ITEM_UP), 0)
        Select Case sts
            Case BtNoErr
            
            
            
            
                If StrConv(ITEM_UPREC.BEF_1_L_PAPER, vbUnicode) <> StrConv(ITEM_UPREC.BEF_LAST_L_PAPER, vbUnicode) Or _
                    StrConv(ITEM_UPREC.BEF_1_L_PLASTIC, vbUnicode) <> StrConv(ITEM_UPREC.BEF_LAST_L_PLASTIC, vbUnicode) Then
                
                        Call UniCode_Conv(ITEMREC.UPD_TANTO, "CONV2")
                        Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                
                    If Mode = 1 Then
                    
                    
                        Call UniCode_Conv(K0_ITEM.JGYOBU, StrConv(ITEM_UPREC.JGYOBU, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(ITEM_UPREC.NAIGAI, vbUnicode))
                        Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(ITEM_UPREC.HIN_GAI, vbUnicode))
                    
                        sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Exit Function
                        End Select
                
                        Call UniCode_Conv(ITEMREC.L_PAPER, StrConv(ITEM_UPREC.BEF_LAST_L_PAPER, vbUnicode))
                        Call UniCode_Conv(ITEMREC.L_PLASTIC, StrConv(ITEM_UPREC.BEF_LAST_L_PLASTIC, vbUnicode))
                
                        Call UniCode_Conv(ITEMREC.UPD_TANTO, "CONV2")
                        Call UniCode_Conv(ITEMREC.UPD_DATETIME, Format(Now, "YYYYMMDDHHMMSS"))
                
                
                        sts = BTRV(BtOpUpdate, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
                        Select Case sts
                            Case BtNoErr
                            Case BtErrKeyNotFound
                            
                            Case Else
                                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                                Exit Function
                        End Select
                
                    
                    
                        Call LOG_OUT(Start_Now & "item.CSV", "," & StrConv(ITEM_UPREC.JGYOBU, vbUnicode) & "," & _
                                    StrConv(ITEM_UPREC.NAIGAI, vbUnicode) & "," & _
                                    StrConv(ITEM_UPREC.HIN_GAI, vbUnicode) & "," & _
                                    StrConv(ITEM_AREC.BEF_1_L_PAPER, vbUnicode) & "," & _
                                    StrConv(ITEM_AREC.BEF_1_L_PLASTIC, vbUnicode) & "," & _
                                    Trim(StrConv(ITEM_AREC.BEF_2_L_PAPER, vbUnicode)) & "," & _
                                    Trim(StrConv(ITEM_AREC.BEF_2_L_PLASTIC, vbUnicode)) & "," & _
                                    Trim(StrConv(ITEM_AREC.BEF_3_L_PAPER, vbUnicode)) & "," & _
                                    Trim(StrConv(ITEM_AREC.BEF_3_L_PLASTIC, vbUnicode)) & "," & _
                                    Trim(StrConv(ITEM_AREC.BEF_LAST_L_PAPER, vbUnicode)) & "," & _
                                    Trim(StrConv(ITEM_AREC.BEF_LAST_L_PLASTIC, vbUnicode)) & "," & "*" & "," & StrConv(ITEMREC.UPD_TANTO, vbUnicode) & "," & Mid(StrConv(ITEMREC.UPD_DATETIME, vbUnicode), 1, 8) & " " & Mid(StrConv(ITEMREC.UPD_DATETIME, vbUnicode), 9, 6))
                    
                    Else
                    
                    
                        Call LOG_OUT(Start_Now & "item.CSV", "," & StrConv(ITEM_UPREC.JGYOBU, vbUnicode) & "," & _
                                    StrConv(ITEM_UPREC.NAIGAI, vbUnicode) & "," & _
                                    StrConv(ITEM_UPREC.HIN_GAI, vbUnicode) & "," & _
                                    StrConv(ITEM_AREC.BEF_1_L_PAPER, vbUnicode) & "," & _
                                    StrConv(ITEM_AREC.BEF_1_L_PLASTIC, vbUnicode) & "," & _
                                    Trim(StrConv(ITEM_AREC.BEF_2_L_PAPER, vbUnicode)) & "," & _
                                    Trim(StrConv(ITEM_AREC.BEF_2_L_PLASTIC, vbUnicode)) & "," & _
                                    Trim(StrConv(ITEM_AREC.BEF_3_L_PAPER, vbUnicode)) & "," & _
                                    Trim(StrConv(ITEM_AREC.BEF_3_L_PLASTIC, vbUnicode)) & "," & _
                                    Trim(StrConv(ITEM_AREC.BEF_LAST_L_PAPER, vbUnicode)) & "," & _
                                    Trim(StrConv(ITEM_AREC.BEF_LAST_L_PLASTIC, vbUnicode)) & "," & "*" & "," & StrConv(ITEM_UPREC.UPD_TANTO, vbUnicode) & "," & Mid(StrConv(ITEM_UPREC.UPD_DATETIME, vbUnicode), 1, 8) & " " & Mid(StrConv(ITEMREC.UPD_DATETIME, vbUnicode), 9, 6))
                    
                    
                    
                    End If
                
                
                
                
        
        
                
                
                
                
                    upd_count = upd_count + 1
                    Cnt(8).Caption = Format(upd_count, "#0")
               
                
                
                
                End If
            
            
            
            
            
            
            
            
            
            
            
            
            
            Case BtErrKeyNotFound
            
            
            
            
            
            
                    Call LOG_OUT(Start_Now & "item.CSV", "," & StrConv(ITEM_UPREC.JGYOBU, vbUnicode) & "," & _
                                StrConv(ITEM_AREC.NAIGAI, vbUnicode) & "," & _
                                StrConv(ITEM_AREC.HIN_GAI, vbUnicode) & "," & _
                                StrConv(ITEM_AREC.BEF_1_L_PAPER, vbUnicode) & "," & _
                                StrConv(ITEM_AREC.BEF_1_L_PLASTIC, vbUnicode) & "," & _
                                Trim(StrConv(ITEM_AREC.BEF_2_L_PAPER, vbUnicode)) & "," & _
                                Trim(StrConv(ITEM_AREC.BEF_2_L_PLASTIC, vbUnicode)) & "," & _
                                Trim(StrConv(ITEM_AREC.BEF_3_L_PAPER, vbUnicode)) & "," & _
                                Trim(StrConv(ITEM_AREC.BEF_3_L_PLASTIC, vbUnicode)) & "," & _
                                Trim(StrConv(ITEM_AREC.BEF_LAST_L_PAPER, vbUnicode)) & "," & _
                                Trim(StrConv(ITEM_AREC.BEF_LAST_L_PLASTIC, vbUnicode)) & "," & "," & StrConv(ITEM_AREC.UPD_TANTO, vbUnicode) & "," & Mid(StrConv(ITEM_AREC.UPD_DATETIME, vbUnicode), 1, 8) & " " & Mid(StrConv(ITEM_AREC.UPD_DATETIME, vbUnicode), 9, 6))
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            Case Else
                Call File_Error(sts, BtOpGetEqual, "品目マスタ")
                Exit Function
        End Select
        
        
        
        
        
        
            sel_count = sel_count + 1
            Cnt(7).Caption = Format(sel_count, "#0")
                
        com = BtOpGetNext
    Loop

                
                
    








    Cnt(6).Caption = Format(Count, "#0")

    Label3(1).Caption = Format(Now)
    Me.MousePointer = vbDefault

'---------------------------------------------  終了
Update_End:
    
    Update_Proc = False

End Function




Private Sub Command1_Click(Index As Integer)
    
Dim ans As Integer
    
Dim sts As Integer
    
Dim FullPath    As String
Dim c           As String * 128


    
    
    
    
    
    Select Case Index
        Case 0
            ans = MsgBox("「チェック」実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                
                
                
                If Update_Proc(0) Then
                    Unload Me
                End If
            
            
            
            
                MsgBox "終了しました"
            
            End If


        Case 1

            ans = MsgBox("「チェック＆更新」実行しますか？", vbYesNo + vbQuestion, "確認入力")
            If ans = vbYes Then
                
                
                
                
                If Update_Proc(1) Then
                    Unload Me
                End If
            
            
            
            
                MsgBox "終了しました"
            
            End If




        Case 2
            
            Unload Me
    
    End Select

End Sub

Private Sub Form_Activate()

Dim ans As Integer
                                
                                
    Text1(0).Text = "20100716164000"
    Text1(1).Text = "20101231235959"
                                
    Text1(2).Text = "20100716164000"
    Text1(3).Text = "20101231235959"
                                
                                
                                

End Sub

Private Sub Form_DblClick()
    PrintForm
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim c As String * 128
Dim sts As Integer
    
    If App.PrevInstance Then
        Beep
        MsgBox "同一プログラム実行中です。"
        End
    End If
   
    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)
    
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
                    
    If ITEM_A_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
    If ITEM_B_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
    If ITEM_C_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
                    
    If ITEM_UP_Open(BtOpenNomal) Then
        Unload Me
    End If
                    
                    
                    
                    
                    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目ﾏｽﾀ")
        End If
    End If
    
    
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set ITEM_DATA_SET1 = Nothing

    End
End Sub

