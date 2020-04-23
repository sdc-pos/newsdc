VERSION 5.00
Begin VB.Form F1030801 
   BackColor       =   &H00FFFFFF&
   Caption         =   "出荷予定メンテナンス"
   ClientHeight    =   8610
   ClientLeft      =   2130
   ClientTop       =   2715
   ClientWidth     =   14370
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
   MaxButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   14370
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   0
      Left            =   945
      MaxLength       =   8
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox Text 
      Height          =   360
      Index           =   1
      Left            =   3360
      MaxLength       =   8
      TabIndex        =   44
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "全件対象"
      Height          =   495
      Left            =   8400
      TabIndex        =   43
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtID_No 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1560
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   1
      Left            =   4560
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   5
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   6
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   7
      Left            =   6000
      MaxLength       =   2
      TabIndex        =   8
      Top             =   1320
      Width           =   375
   End
   Begin VB.ListBox List1 
      Height          =   4140
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2880
      Width           =   13335
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   8
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox Text 
      Alignment       =   1  '右揃え
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   11
      Left            =   10320
      MaxLength       =   6
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   10
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   25
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2040
      Width           =   3735
   End
   Begin VB.TextBox Text 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Index           =   9
      Left            =   3840
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ComboBox Combo 
      Height          =   360
      Index           =   0
      Left            =   1320
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "終  了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "表　示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "削 除"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command 
      Caption         =   "更  新"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   1785
      TabIndex        =   46
      Top             =   120
      Visible         =   0   'False
      Width           =   3915
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "担当者"
      Height          =   255
      Index           =   4
      Left            =   105
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID№"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   42
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "済数"
      Height          =   255
      Index           =   13
      Left            =   10680
      TabIndex        =   40
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "予定数"
      Height          =   255
      Index           =   12
      Left            =   9600
      TabIndex        =   39
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "伝票日付"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   38
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "伝票№"
      Height          =   255
      Index           =   10
      Left            =   1440
      TabIndex        =   37
      Top             =   2520
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   10920
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "＝＝＞"
      Height          =   255
      Index           =   9
      Left            =   3480
      TabIndex        =   36
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   8
      Left            =   5040
      TabIndex        =   35
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   7
      Left            =   5760
      TabIndex        =   34
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label LabJIGYO 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
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
      TabIndex        =   33
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "向け先"
      Height          =   255
      Index           =   17
      Left            =   2520
      TabIndex        =   32
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   31
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "/"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   30
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "伝票日付"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "注文区分"
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   28
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品  名"
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   27
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label 
      BackColor       =   &H00FFFFFF&
      Caption         =   "品  番"
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   26
      Top             =   2520
      Width           =   735
   End
   Begin VB.Menu MainMenu 
      Caption         =   "事業部"
      Begin VB.Menu SubMenu 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "F1030801"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim P_ID As String * 8
Dim WS_ID As String * 3



Private Const Text_Max% = 11
Private Const ptxTANTO_CODE% = 0        '担当者コード           2007.11.06
Private Const ptxMUKE_CODE% = 1         '向け先（コード入力用）
Private Const ptxM_YY% = 2
Private Const ptxM_MM% = 3
Private Const ptxM_DD% = 4
Private Const ptxC_YY% = 5
Private Const ptxC_MM% = 6
Private Const ptxC_DD% = 7
Private Const ptxNO% = 8
Private Const ptxITEM% = 9
Private Const ptxI_NM% = 10
Private Const ptxQTY% = 11

Private Const pcmbC_KBN% = 0
Private Const pcmbMUKE_CODE% = 1

Private Upd_Mode    As Integer


Private YOIN_CANCEL As String * 2       '出荷ｷｬﾝｾﾙの要因        2007.11.06
Private MENU_NO     As String * 2       '実績ログ出力用ﾒﾆｭｰ№   2007.11.06

'Private Const LAST_UPDATE_DAY$ = "[F103080]2010.11.22 08:30"
Private Const LAST_UPDATE_DAY$ = "[F103080]2020.04.14 11:30 出荷予定数6桁実行時エラー修正"

                                    '画面初期状態を設定する
Private Sub Clear_Field(Start_Pos As Integer)
Dim i As Integer
    
    For i = Start_Pos To Text_Max
        Text(i).Text = ""
    Next i
    
    txtID_No.Text = ""      '2007.11.02
End Sub
Private Function P_Off() As Integer

Dim sts As Integer
Dim com As Integer
Dim yn As Integer
    
    P_Off = True
    
    Call UniCode_Conv(K4_Y_SYU.WEL_ID, WS_ID)
    Call UniCode_Conv(K4_Y_SYU.PRG_ID, P_ID)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrKeyNotFound
                P_Off = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                yn = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                
                If yn = vbCancel Then
                    P_Off = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定データ")
                Exit Function
        End Select
    Loop
        
    Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
    Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
        
'    Select Case StrConv(Y_SYUREC.KAN_KBN, vbUnicode)
'        Case KAN_L_SOFF_POFF_KOFF
'            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_SOFF_POFF_KOFF)
'        Case KAN_L_SING_POFF_KOFF
'            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_SING_POFF_KOFF)
'        Case KAN_L_SOFF_PON_KOFF
'            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_SOFF_PON_KOFF)
'        Case KAN_L_SING_PON_KOFF
'            Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_SING_PON_KOFF)
'    End Select
        
    Do
        sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
        Select Case sts
            Case BtNoErr
                Exit Do
                
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                yn = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                
                If yn = vbCancel Then
                    
                    sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "出荷予定データ")
                        Exit Function
                    End If
                    P_Off = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定データ")
                Exit Function
        End Select
    Loop

    P_Off = False
End Function

Private Function Item_Dsp() As Integer
                                    '各項目を表示する
Dim sts         As Integer
Dim yn          As Integer
Dim i           As Integer
Dim Qty         As Long
Dim ans         As Integer
    
    Item_Dsp = True
    
        
    sts = P_Off
    Select Case sts
        Case False
        Case SYS_ERR
            Exit Function
        Case SYS_CANCEL
            List1.SetFocus
            Item_Dsp = SYS_CANCEL
            Exit Function
    End Select
    
                                                '出荷予定読み込み
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
'    Call UniCode_Conv(K0_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbC_KBN).Text, 1))2004.04.08
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, Mid(List1.List(List1.ListIndex), 19, 12))
    
    
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                If Len(Trim(StrConv(Y_SYUREC.WEL_ID, vbUnicode))) <> 0 And _
                    StrConv(Y_SYUREC.WEL_ID, vbUnicode) <> WS_ID Then
                    
                    sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "出荷予定データ")
                        Exit Function
                    End If
                    Beep
                    yn = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If yn = vbCancel Then
                        List1.SetFocus
                        Item_Dsp = SYS_CANCEL
                        Exit Function
                    End If
                Else
                    Exit Do
                End If
            Case BtErrKeyNotFound
                Beep
                MsgBox "データ内容が変更されています。最新内容を表示します。"
                If List_Dsp() Then
                    Exit Function
                End If
                List1.SetFocus
                Item_Dsp = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                yn = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If yn = vbCancel Then
                    sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K4_Y_SYU, Len(K4_Y_SYU), 4)
                    If sts Then
                        Call File_Error(sts, BtOpUnlock, "出荷予定データ")
                        Exit Function
                    End If
                    List1.SetFocus
                    Item_Dsp = SYS_CANCEL
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定データ")
                Exit Function
        End Select
    Loop
                                        '完了のチェック
    If Upd_Mode = 1 Then
    Else
        If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then
            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            If sts Then
                Call File_Error(sts, BtOpUnlock, "出荷予定データ")
                Exit Function
            End If
                
            Beep
            MsgBox "指定の伝票は出庫完了の為、変更できません。"
            List1.SetFocus
            List1.ListIndex = 0
            Item_Dsp = False
            Exit Function
        End If
    End If
    
    If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> 0 Then
        Beep
        yn = MsgBox("指定の伝票は作業中です。処理を継続しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")
        If yn = vbNo Then
            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            If sts Then
                Call File_Error(sts, BtOpUnlock, "出荷予定データ")
                Item_Dsp = SYS_ERR
                Exit Function
            End If
            
            List1.SetFocus
            List1.ListIndex = 0
            Item_Dsp = False
            Exit Function
        End If
    End If
        
        
    If Len(Trim(StrConv(Y_SYUREC.PRINT_YMD, vbUnicode))) <> 0 Then
        Beep
        yn = MsgBox("指定の伝票は出庫表印刷済です。処理を継続しますか？", vbYesNo + vbQuestion + vbDefaultButton2, "確認入力")
        If yn = vbNo Then
            sts = BTRV(BtOpUnlock, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
            If sts Then
                Call File_Error(sts, BtOpUnlock, "出荷予定データ")
                Item_Dsp = SYS_ERR
                Exit Function
            End If
                
            List1.SetFocus
            List1.ListIndex = 0
            Item_Dsp = False
            Exit Function
            
        End If
    End If
        
    
    Text(ptxM_YY).Text = Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 1, 4)
    Text(ptxM_MM).Text = Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 5, 2)
    Text(ptxM_DD).Text = Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 7, 2)

    Text(ptxC_YY).Text = Text(ptxM_YY)
    Text(ptxC_MM).Text = Text(ptxM_MM)
    Text(ptxC_DD).Text = Text(ptxM_DD)

    Text(ptxNO).Text = StrConv(Y_SYUREC.DEN_NO, vbUnicode)
    txtID_No.Text = StrConv(Y_SYUREC.ID_NO, vbUnicode)
    
    Text(ptxITEM).Text = StrConv(Y_SYUREC.HIN_NO, vbUnicode)
                                                        '品目マスタ読込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "品目マスタ")
            Item_Dsp = SYS_ERR
            Exit Function
    End Select
    
    Text(ptxI_NM) = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    Text(ptxQTY) = Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#0")
    
'    Text(ptxHS_C_KBN) = StrConv(Y_SYUREC.CYU_KBN, vbUnicode)
    
    Call UniCode_Conv(Y_SYUREC.PRG_ID, P_ID)
    Call UniCode_Conv(Y_SYUREC.WEL_ID, WS_ID)
    
    Do
        sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                Exit Do
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                yn = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If yn = vbCancel Then
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpUpdate, "出荷予定データ")
                Item_Dsp = SYS_ERR
                Exit Function
        End Select
    Loop
    Text(ptxC_YY).SetFocus
    
    Item_Dsp = False
    
    
End Function

Private Function Err_Chk(Mode As Integer) As Integer
                                    
Dim sts As Integer
Dim i   As Integer
                                    '入力項目のエラーチェック

    Err_Chk = True
                                    
    '担当者     2007.11.06
'    Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(ptxTANTO_CODE).Text)
'    sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
'    Select Case sts
'        Case BtNoErr
'            Label(15).Caption = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
'        Case BtErrKeyNotFound
'            Label(15).Caption = ""
'            Beep
'            MsgBox "入力した項目はエラーです。(担当者コード)"
'            Text(ptxTANTO_CODE).SetFocus
'
'            Exit Function
'
'        Case Else
'            Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
'            Err_Chk = SYS_ERR
'            Exit Function
'    End Select
                                    
                                    
                                    
                                    
                                    
                                    '出荷予定読み込み
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
'    Call UniCode_Conv(K0_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbC_KBN).Text, 1))2004.04.08
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, txtID_No)
    sts = BTRV(BtOpGetEqual, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    Select Case sts
        Case BtNoErr
            If StrConv(Y_SYUREC.WEL_ID, vbUnicode) = WS_ID And _
                StrConv(Y_SYUREC.PRG_ID, vbUnicode) = P_ID Then
            Else
                Beep
                MsgBox "更新対象の出荷予定が確定していません。"
                List1.SetFocus
                Exit Function
            End If
                                    '削除時のチェック
            If Mode = 1 Then
                If Upd_Mode = 0 Then
                    If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) <> 0 Then
                        Beep
                        MsgBox "出庫実績が有るので削除できません。"
                        List1.SetFocus
                    Else
                        Err_Chk = False
                        Exit Function
                    End If
                End If
            End If
        Case BtErrKeyNotFound
            Beep
            MsgBox "更新対象の出荷予定が確定していません。"
            List1.SetFocus
            Exit Function
        Case Else
            Call File_Error(sts, BtOpGetEqual, "出荷予定データ")
            Err_Chk = SYS_ERR
            Exit Function
    End Select


    For i = ptxC_YY To ptxC_DD
        If Trim(Text(i).Text) = "" Then
            Text(i).Text = "0"
        End If
    
        If Not IsNumeric(Text(i).Text) Then
            Beep
            MsgBox "入力した項目はエラーです。"
            Text(i).SetFocus
            Exit Function
        Else
            Text(i).Text = Right(Format(CInt(Text(i).Text), "0000"), Text(i).MaxLength)
        End If
    
    Next i

    If Not IsDate(Text(ptxC_YY).Text & "/" & Text(ptxC_MM).Text & "/" & Text(ptxC_DD).Text) Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxC_YY).SetFocus
        Exit Function
    End If
    If (Text(ptxC_YY).Text & "/" & Text(ptxC_MM).Text & "/" & Text(ptxC_DD).Text) < Format(Date, "YYYY/MM/DD") Then
        Beep
        MsgBox "入力した項目はエラー（＜本日）です。"
        Text(ptxC_YY).SetFocus
        Exit Function
    End If

    If Not IsNumeric(Text(ptxQTY).Text) Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxQTY).SetFocus
        Exit Function
    Else
        Text(ptxQTY).Text = Format(CLng(Text(ptxQTY).Text), "#0")
    End If

    If CLng(Text(ptxQTY).Text) <= 0 Then
        Beep
        MsgBox "入力した項目はエラーです。"
        Text(ptxQTY).SetFocus
        Exit Function
    End If
'2009.04.14                                    '数量変更のチェック
'    If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) > CLng(Text(ptxQTY).Text) Then
'        Beep
'        MsgBox "出庫実績未満への数量変更はできません。"
'        Text(ptxQTY).SetFocus
'        Exit Function
'    End If
    
    Err_Chk = False
    
End Function

                                            '出荷予定の訂正／削除
Private Function Update_Proc(Mode As Integer) As Integer

Dim sts As Integer
Dim ans As Integer
                                            
    Update_Proc = True
    
                                                '出荷予定読み込み
    Call UniCode_Conv(K0_Y_SYU.JGYOBU, Last_JGYOBU)
'    Call UniCode_Conv(K0_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbC_KBN).Text, 1))2004.04.08
    Call UniCode_Conv(K0_Y_SYU.KEY_ID_NO, txtID_No)
    Do
        sts = BTRV(BtOpGetEqual + BtSNoWait, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
        Select Case sts
            Case BtNoErr
                
                Exit Do
            Case BtErrKeyNotFound
                Beep
                MsgBox "更新対象の出荷予定が確定していません。"
                List1.SetFocus
                Update_Proc = False
                Exit Function
            Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                Beep
                ans = MsgBox("他端末でデータ使用中です。<Y_SYUKA.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                If ans = vbCancel Then
                    List1.SetFocus
                    Update_Proc = False
                    Exit Function
                End If
            Case Else
                Call File_Error(sts, BtOpGetEqual, "出荷予定データ")
                Update_Proc = SYS_ERR
                Exit Function
        End Select
    Loop
    
    Select Case Mode
        Case 0                              '訂正
                                            
            If SYUKA_LOG_ON Then
                Call SYUKA_LOG_OUT_PROC("UPD", "BEF")
            End If
                                            
                                            '出荷予定データ更新
            Call UniCode_Conv(Y_SYUREC.SURYO, Format(CLng(Text(ptxQTY).Text), "0000000"))
            Call UniCode_Conv(Y_SYUREC.SYUKA_YMD, _
                            Text(ptxC_YY).Text & Text(ptxC_MM).Text & Text(ptxC_DD).Text)
            Call UniCode_Conv(Y_SYUREC.WEL_ID, "")
            Call UniCode_Conv(Y_SYUREC.PRG_ID, "")
            
            If CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)) = CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)) Then
                                            '出庫完了
                Call UniCode_Conv(Y_SYUREC.KAN_KBN, KAN_KBN_FIN)
            
            End If
            
            Do
                sts = BTRV(BtOpUpdate, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            List1.SetFocus
                            Update_Proc = False
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpUpdate, "出荷予定データ")
                        Update_Proc = SYS_ERR
                        Exit Function
                End Select
            Loop
            
            If SYUKA_LOG_ON Then
                Call SYUKA_LOG_OUT_PROC("UPD", "AFT")
            End If
            
            List1.RemoveItem List1.ListIndex
            Call List_Edit
        Case 1                              '削除
                                            '出荷予定データ削除
            Do
                sts = BTRV(BtOpDelete, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
                Select Case sts
                    Case BtNoErr
                        Exit Do
                    Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                        Beep
                        ans = MsgBox("他端末でデータ使用中です。", vbRetryCancel + vbQuestion, "確認入力")
                        If ans = vbCancel Then
                            Update_Proc = False
                            List1.SetFocus
                            Exit Function
                        End If
                    Case Else
                        Call File_Error(sts, BtOpDelete, "出荷予定データ")
                        Update_Proc = SYS_ERR
                        Exit Function
                End Select
            Loop
            
            
            '作業ログ出力 2007.11.06
            If MENU_NO <> "" Then
                If P_SAGYO_LOG_OUTPUT_PROC(WS_ID, _
                                            WS_ID, _
                                            StrConv(Y_SYUREC.JGYOBU, vbUnicode), _
                                            StrConv(Y_SYUREC.NAIGAI, vbUnicode), _
                                            MENU_NO, _
                                            YOIN_CANCEL, _
                                            StrConv(Y_SYUREC.HIN_NO, vbUnicode), _
                                            0, 0, , , _
                                            StrConv(Y_SYUREC.ID_NO, vbUnicode), _
                                            StrConv(Y_SYUREC.MUKE_CODE, vbUnicode), _
                                            StrConv(Y_SYUREC.SS_CODE, vbUnicode)) Then
        
        
                    Exit Function
                End If
            End If
            
            
            
            
            
            If SYUKA_LOG_ON Then
                Call SYUKA_LOG_OUT_PROC("DEL", "AFT")
            End If
            
            List1.RemoveItem List1.ListIndex
        
    End Select
    Update_Proc = False
    
End Function

Private Function MTS_Set_Proc() As Integer

Dim sts     As Integer
Dim com     As Integer
    
Dim Edit    As String
    
    MTS_Set_Proc = True
    
'    Call Input_Lock
    
    
    Combo(pcmbMUKE_CODE).Clear
    
    com = BtOpGetFirst
    Do
        sts = BTRV(com, MTS_POS, MTSREC, Len(MTSREC), K1_MTS, Len(K1_MTS), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "向け先マスタ")
                Exit Function
        End Select
        
        Edit = StrConv(MTSREC.MUKE_DNAME, vbUnicode) & "   "
        Edit = Edit & StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode)
        
        
        Combo(pcmbMUKE_CODE).AddItem Edit
    
        com = BtOpGetNext
    
    Loop

    If Combo(pcmbMUKE_CODE).ListCount <= 0 Then
    Else
        Combo(pcmbMUKE_CODE).ListIndex = 0
    End If

'    Call Input_UnLock

    MTS_Set_Proc = False
End Function

Private Function List_Dsp() As Integer
Dim sts As Integer
Dim com As Integer
Dim yn As Integer
Dim WS01 As Integer
Dim W_Str As String

    List_Dsp = True

    List1.Clear
    
    
                                                    '事業部
    Call UniCode_Conv(K2_Y_SYU.JGYOBU, Last_JGYOBU)
                                                    '注文区分
    Call UniCode_Conv(K2_Y_SYU.KEY_CYU_KBN, Right(Combo(pcmbC_KBN).Text, 1))
                                                    '向け先
    Call UniCode_Conv(K2_Y_SYU.KEY_MUKE_CODE, Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8))
    Call UniCode_Conv(K2_Y_SYU.KEY_SS_CODE, Right(Combo(pcmbMUKE_CODE).Text, 8))
    
    
    com = BtOpGetGreaterEqual
    Do
        sts = BTRV(com, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K2_Y_SYU, Len(K2_Y_SYU), 2)
        Select Case sts
            Case BtNoErr
                If StrConv(Y_SYUREC.JGYOBU, vbUnicode) <> Last_JGYOBU Then
                    Exit Do
                End If
            
            Case BtErrEOF
                Exit Do
            Case Else
                Call File_Error(sts, com, "出荷予定データ")
                Exit Function
        End Select
        
        If StrConv(Y_SYUREC.KEY_CYU_KBN, vbUnicode) <> Right(Combo(pcmbC_KBN).Text, 1) Then
            Exit Do
        End If
        
        If StrConv(Y_SYUREC.KEY_MUKE_CODE, vbUnicode) <> Left(Right(Combo(pcmbMUKE_CODE).Text, 16), 8) Or _
           StrConv(Y_SYUREC.SS_CODE, vbUnicode) <> Right(Combo(pcmbMUKE_CODE).Text, 8) Then
            Exit Do
        End If
        
        
        If Upd_Mode = 0 Then
            If Len(Trim(StrConv(Y_SYUREC.KAN_YMD, vbUnicode))) <> 0 Then
            Else
                If List_Edit() Then
                    Exit Function
                End If
            End If
        Else
            If List_Edit() Then
                Exit Function
            End If
        End If
        
        com = BtOpGetNext
    
    Loop

    List_Dsp = False

End Function
Private Function List_Edit() As Integer
Dim sts     As Integer
Dim com     As Integer
Dim yn      As Integer
Dim WS01    As Integer
Dim Edit    As String
Dim RetBuf  As String
    
Dim strWork As String
    
    List_Edit = True
    
    Edit = ""
    Edit = Edit & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 1, 4) & "/" & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 5, 2) & "/" & Mid(StrConv(Y_SYUREC.SYUKA_YMD, vbUnicode), 7, 2) & " "
    Edit = Edit & Left(StrConv(Y_SYUREC.DEN_NO, vbUnicode), 6) & "(" & StrConv(Y_SYUREC.ID_NO, vbUnicode) & ") "
    Edit = Edit & Left(StrConv(Y_SYUREC.HIN_NO, vbUnicode), 13) & "  "
                                                        '品目マスタ読込み
    Call UniCode_Conv(K0_ITEM.JGYOBU, Last_JGYOBU)
    Call UniCode_Conv(K0_ITEM.NAIGAI, StrConv(Y_SYUREC.NAIGAI, vbUnicode))
    Call UniCode_Conv(K0_ITEM.HIN_GAI, StrConv(Y_SYUREC.HIN_NO, vbUnicode))
        
    sts = BTRV(BtOpGetEqual, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrKeyNotFound
            Call UniCode_Conv(ITEMREC.HIN_NAME, "")
        Case Else
            Call File_Error(sts, BtOpGetEqual, "出荷予定データ")
            Exit Function
    End Select
    
    
'    Edit = Edit & Left(StrConv(ITEMREC.HIN_NAME, vbUnicode), 25) & " "
    strWork = StrConv(ITEMREC.HIN_NAME, vbUnicode)
    
If LenB(strWork) <> 80 Then
    Edit = Edit & StrConv(LeftB(strWork, 30), vbWide)
Else
    Edit = Edit & Left(strWork, 30)
End If
    
''    Edit = Edit & LeftB(strWork, 50)
            
            
    RetBuf = Format(CLng(StrConv(Y_SYUREC.SURYO, vbUnicode)), "#,##0")
'    RetBuf = Space(6 - Len(RetBuf)) & Trim(RetBuf) & "("
    RetBuf = Right(Space(6) & Trim(RetBuf), 7) & "("                  '2020/04/14 桁数6桁(カンマ含むと7桁)で実行時エラーになる為、修正
    
    Edit = Edit & RetBuf
    RetBuf = Format(CLng(StrConv(Y_SYUREC.JITU_SURYO, vbUnicode)), "#,##0")
'    RetBuf = Space(6 - Len(RetBuf)) & Trim(RetBuf) & ")"
    RetBuf = Right(Space(6) & Trim(RetBuf), 7) & ")"                  '2020/04/14 桁数6桁(カンマ含むと7桁)で実行時エラーになる為、修正
    Edit = Edit & RetBuf
    
    
    List1.AddItem Edit

    List_Edit = False

End Function

Private Sub Combo_Click(Index As Integer)
        
    Select Case Index
        Case pcmbC_KBN
            '向け先設定
            If MTS_Set_Proc() Then
                Unload Me
            End If
    
            Combo(pcmbC_KBN).SetFocus
    End Select


End Sub

Private Sub Combo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim sts As Integer
  
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        
        Case pcmbC_KBN
            '向け先設定
            If MTS_Set_Proc() Then
                Unload Me
            End If
        
            Combo(pcmbMUKE_CODE).SetFocus
        
        Case pcmbMUKE_CODE
            '出荷予定表示
            If List_Dsp() Then
                Unload Me
            End If
    
    End Select
End Sub



Private Sub Command_Click(Index As Integer)
Dim yn As Integer
Dim sts As Integer

    Select Case Index
        Case 0
                                            'エラーチェック
            sts = Err_Chk(0)
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            Beep
            yn = MsgBox("更新しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                sts = Update_Proc(0)
                Select Case sts
                    Case True, False
                    Case SYS_ERR
                        Unload Me
                End Select
            Else
                sts = P_Off()
                Select Case sts
                    Case False
                    Case SYS_ERR
                        Unload Me
                    Case SYS_CANCEL
                        List1.SetFocus
                        Exit Sub
                    End Select
            End If
            
            Call Clear_Field(1)
            
            If List1.ListCount <> 0 Then
                List1.SetFocus
            Else
                Combo(pcmbC_KBN).SetFocus
            End If
        Case 3
                                            'エラーチェック
            sts = Err_Chk(1)
            Select Case sts
                Case False
                Case True
                    Exit Sub
                Case SYS_ERR
                    Unload Me
            End Select
            
            Beep
            yn = MsgBox("削除しますか？", vbYesNo + vbQuestion, "確認入力")
            If yn = vbYes Then
                If Update_Proc(1) Then
                    Unload Me
                End If
            End If
            
            Call Clear_Field(1)
            
            If List1.ListCount <> 0 Then
                List1.SetFocus
            Else
                Combo(pcmbC_KBN).SetFocus
            End If
        Case 7
            If List_Dsp() Then
            End If
        Case 11
            Unload Me
        Case Else
            Beep
    End Select
    
End Sub


Private Sub Command1_Click()
    
    If Command(0).Enabled = False Then
'        Command(0).Enabled = True
        Upd_Mode = 0
    Else
'        Command(0).Enabled = False
        Upd_Mode = 1
    End If

    If List_Dsp() Then
        Unload Me
    End If

End Sub

'Private Sub Form_DblClick() 2020/04/14 コメントアウト
'    PrintForm
'End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
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
Dim com     As Integer
Dim sts     As Integer
Dim yn      As Integer
Dim RetBuf  As String * 255

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
                                '出荷ログファイル名取り込み
    If SYUKA_LOGF_GET_PROC() Then
        Beep
        MsgBox "出荷ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    
'---------------------------------------------- '出荷ｷｬﾝｾﾙ要因の獲得    2007.11.06
    If GetIni(App.EXEName, "YOIN", "SYS", c) Then
        Beep
        MsgBox "出荷ｷｬﾝｾﾙ要因の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    YOIN_CANCEL = RTrim(c)
'---------------------------------------------- 'ﾒﾆｭｰ№の獲得    2007.11.06
    If GetIni(App.EXEName, "MENU_NO", "SYS", c) Then
        MENU_NO = ""
    Else
        MENU_NO = RTrim(c)
    End If
    
        
    
    
    
    P_ID = StrConv(App.EXEName, vbUpperCase)
    
    If GetComputerNameA(RetBuf, 255) <> 0 Then
        WS_ID = Trim(Left(RetBuf, InStr(RetBuf, vbNullChar) - 1))
    Else
        WS_ID = "???"
    End If
                                '事業部取り込み
    If JGYOB_TB_Set Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If

    For i = 0 To UBound(JGYOBU_T)
'        If JGYOBU_T(i).Code = " " Then
'            Unload SubMenu(i)
'            Exit For
'        End If

        Load SubMenu(i + 1)
        SubMenu(i).Caption = RTrim(JGYOBU_T(i).NAME)

        If JGYOBU_T(i).CODE = Last_JGYOBU Then
            F1030801.Caption = "出荷予定メンテナンス（" + RTrim(JGYOBU_T(i).NAME) + ") " & LAST_UPDATE_DAY
            SubMenu(i).Checked = True
            LabJIGYO.Caption = RTrim(JGYOBU_T(i).NAME)
            LabJIGYO.ForeColor = QBColor(JGYOBU_T(i).COLOR)
'            LabJIGYO.BorderStyle = 1
        Else
            SubMenu(i).Checked = False
        End If
    Next i
    Unload SubMenu(i)


                                '品目マスタＯＰＥＮ
    If ITEM_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '向け先管理マスタＯＰＥＮ
    If MTS_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '担当者マスタＯＰＥＮ   2007.11.06
    If TANTO_Open(BtOpenNomal) Then
        Unload Me
    End If
                                '出荷予定データファイルＯＰＥＮ
    If Y_SYU_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
'---------------------------------------------- '作業実績ﾛｸﾞＯＰＥＮ    2007.11.06
    If P_SAGYO_LOG_Open(BtOpenNomal) Then
        Unload Me
    End If
                                
                                
                                
                                
                                
                                
    Upd_Mode = 0
                                '画面初期設定
                                    
                                    '注文区分
    Combo(pcmbC_KBN).Clear
    Combo(pcmbC_KBN).AddItem CYU_KBN_1 & "   " & CYU_KBN_TUK
    Combo(pcmbC_KBN).AddItem CYU_KBN_2 & "   " & CYU_KBN_SPO
    Combo(pcmbC_KBN).AddItem CYU_KBN_3 & "   " & CYU_KBN_HJU
'    Combo(pcmbC_KBN).AddItem CYU_KBN_4
    Combo(pcmbC_KBN).AddItem CYU_KBN_E & "   " & CYU_KBN_BOU
    Combo(pcmbC_KBN).AddItem CYU_KBN_T & "   " & CYU_KBN_KIN
    
    Combo(pcmbC_KBN).ListIndex = 0
    
    
    
    
    If MTS_Set_Proc() Then
        Unload Me
    End If
    
    Call Clear_Field(0)
        
    Combo(pcmbC_KBN).SetFocus

'    Text(ptxTANTO_CODE).SetFocus    '2007.11.06
    End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim sts As Integer
    
    sts = P_Off
                                            '品目マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目マスタ")
        End If
    End If
                                            '向け先管理マスタＣＬＯＳＥ
    sts = BTRV(BtOpClose, MTS_POS, MTSREC, Len(MTSREC), K0_MTS, Len(K0_MTS), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "向け先管理マスタ")
        End If
    End If
                                            '担当者マスタＣＬＯＳＥ 2007.11.06
    sts = BTRV(BtOpClose, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "担当者マスタ")
        End If
    End If
                                            '出荷予定データファイルＣＬＯＳＥ
    sts = BTRV(BtOpClose, Y_SYU_POS, Y_SYUREC, Len(Y_SYUREC), K0_Y_SYU, Len(K0_Y_SYU), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "出荷予定データファイル")
        End If
    End If
                                            '作業ログファイルＣＬＯＳＥ 2007.11.06
    sts = BTRV(BtOpClose, P_SAGYO_LOG_POS, P_SAGYO_LOG_REC, Len(P_SAGYO_LOG_REC), K0_P_SAGYO_LOG, Len(K0_P_SAGYO_LOG), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "作業ログファイル")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_POS, ITEMREC, Len(ITEMREC), K0_ITEM, Len(K0_ITEM), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If
    
    Set F1030801 = Nothing
    
    End
End Sub
Private Sub List1_DblClick()
Dim sts As Integer

    sts = Item_Dsp()
    Select Case sts
        Case False
        Case SYS_CANCEL
        Case Else
            Unload Me
    End Select
    
End Sub


Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)

Dim sts As Integer
    
    If List1.ListCount = 0 Then
        Exit Sub
    End If
        
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    sts = Item_Dsp()
    
    Select Case sts
        Case False
        Case SYS_CANCEL
        Case Else
            Unload Me
    End Select

End Sub


Private Sub SubMenu_Click(Index As Integer)
Dim i As Integer
                                    'メニューより終了要求
'    If JGYOBU_T(Index).CODE = " " Then
'        Unload Me
'    End If

    For i = 0 To UBound(JGYOBU_T)
        If JGYOBU_T(i).CODE = " " Then
            Exit For
        End If
        SubMenu(i).Checked = False
    Next i
                                    '事業部切り替え
    F1030801.Caption = "出荷予定メンテナンス（" + RTrim(JGYOBU_T(Index).NAME) + ") " & LAST_UPDATE_DAY
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
Dim RetBuf As String
Dim i As Integer
Dim sts As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    Select Case Index
        
        Case ptxTANTO_CODE  '2007.11.06
            
            
            Call UniCode_Conv(K0_TANTO.TANTO_CODE, Text(Index).Text)
            sts = BTRV(BtOpGetEqual, TANTO_POS, TANTOREC, Len(TANTOREC), K0_TANTO, Len(K0_TANTO), 0)
            Select Case sts
                Case BtNoErr
                    Label(15).Caption = Trim(StrConv(TANTOREC.TANTO_NAME, vbUnicode))
                Case BtErrKeyNotFound
                    Label(15).Caption = ""
                    Beep
                    MsgBox "入力した項目はエラーです。(向け先コード)"
                    Exit Sub
                                        
                Case Else
                    Call File_Error(sts, BtOpGetEqual, "担当者マスタ")
                    Unload Me
            End Select
        
        
            Combo(pcmbC_KBN).SetFocus
            Exit Sub
        
        Case ptxMUKE_CODE
            Call UniCode_Conv(K2_MTS.MUKE_CODE, Text(Index).Text)
            sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K2_MTS, Len(K2_MTS), 2)
            Select Case sts
                Case BtNoErr
                    If Len(Trim(StrConv(MTSREC.SS_CODE, vbUnicode))) <> 0 Then
                        Beep
                        MsgBox "入力した項目はエラーです。(向け先コード)"
                        Exit Sub
                    End If
                                
                Case BtErrKeyNotFound
                                
                    Call UniCode_Conv(K3_MTS.SS_CODE, Text(Index).Text)
                                                        
                    sts = BTRV(BtOpGetEqual, MTS_POS, MTSREC, Len(MTSREC), K3_MTS, Len(K3_MTS), 3)
                    Select Case sts
                        Case BtNoErr
                                        
                        Case BtErrKeyNotFound
                            Beep
                            MsgBox "入力した項目はエラーです。(向け先コード)"
                            Exit Sub
                                        
                        Case Else
                            Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                            Unload Me
                    End Select

                Case Else
                    Call File_Error(sts, BtOpGetEqual, "向け先管理マスタ")
                    Unload Me
            End Select


            For i = 0 To Combo(pcmbMUKE_CODE).ListCount - 1 '向け先
    
                If Right(Combo(pcmbMUKE_CODE).List(i), 16) = StrConv(MTSREC.MUKE_CODE, vbUnicode) & StrConv(MTSREC.SS_CODE, vbUnicode) Then
                    Combo(pcmbMUKE_CODE).ListIndex = i
                    Exit For
                End If
            
    
            Next

            Combo(pcmbMUKE_CODE).SetFocus
        Case Else
        
    End Select
    For i = Index + 1 To Text_Max
        If Text(i).Visible And Text(i).Enabled And Text(i).TabStop Then
            Text(i).SetFocus
            Call Text_GotFocus(Index)
            Exit Sub
        End If
    Next i
End Sub
