VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form SEM00501 
   Caption         =   "[請求システム]品名カテゴリーマスタ登録"
   ClientHeight    =   9975
   ClientLeft      =   2025
   ClientTop       =   -4470
   ClientWidth     =   12510
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
   OLEDropMode     =   1  '手動
   ScaleHeight     =   9975
   ScaleWidth      =   12510
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox Text1 
      Appearance      =   0  'ﾌﾗｯﾄ
      Height          =   375
      Left            =   1680
      OLEDragMode     =   1  '自動
      OLEDropMode     =   1  '手動
      TabIndex        =   5
      Top             =   1200
      Width           =   6975
   End
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   7575
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   13361
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "№"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "BU"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "品名ｶﾃｺﾞﾘｰｺｰﾄﾞ"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "品名ｶﾃｺﾞﾘｰ名称"
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "見積ﾛｯﾄ数"
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "前後工数   　 (秒/ﾛｯﾄ)"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "前後工数　　　(秒/個)"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   714
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1005"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=900"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=714"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=609"
      Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=0"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=2540"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2434"
      Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=0"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(3).Width=5530"
      Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=5424"
      Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(19)=   "Column(4).Width=2408"
      Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2302"
      Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=2"
      Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(24)=   "Column(5).Width=2408"
      Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=2302"
      Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=2"
      Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(29)=   "Column(6).Width=2408"
      Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=2302"
      Splits(0)._ColumnProps(32)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageFooterFont=   "Size=12,Charset=128,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=ＭＳ ゴシック"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   2
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      OLEDropMode     =   1
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=1200,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(5)   =   ":id=0,.fontname=ＭＳ ゴシック"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(8)   =   ":id=1,.fontname=ＭＳ ゴシック"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(12)  =   ":id=2,.fontname=ＭＳ ゴシック"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(15)  =   ":id=3,.fontname=ＭＳ ゴシック"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=67,.parent=1,.bgcolor=&HFFFF00&"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=76,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=68,.parent=2,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(27)  =   ":id=68,.underline=0,.strikethrough=0,.charset=128"
      _StyleDefs(28)  =   ":id=68,.fontname=ＭＳ ゴシック"
      _StyleDefs(29)  =   "Splits(0).FooterStyle:id=69,.parent=3"
      _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=70,.parent=5"
      _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=72,.parent=6"
      _StyleDefs(32)  =   "Splits(0).EditorStyle:id=71,.parent=7"
      _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=74,.parent=9,.bgcolor=&HFFFF00&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=75,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=82,.parent=67,.alignment=3"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=68"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=69"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=71"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=94,.parent=67,.alignment=0"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=91,.parent=68"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=92,.parent=69"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=93,.parent=71"
      _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=98,.parent=67,.alignment=0"
      _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=95,.parent=68"
      _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=96,.parent=69"
      _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=97,.parent=71"
      _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=20,.parent=67"
      _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=17,.parent=68"
      _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=18,.parent=69"
      _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=19,.parent=71"
      _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=102,.parent=67,.alignment=1"
      _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=99,.parent=68"
      _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=100,.parent=69"
      _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=101,.parent=71"
      _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=114,.parent=67,.alignment=1"
      _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=111,.parent=68"
      _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=112,.parent=69"
      _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=113,.parent=71"
      _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=16,.parent=67,.alignment=1"
      _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=13,.parent=68"
      _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=14,.parent=69"
      _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=15,.parent=71"
      _StyleDefs(66)  =   "Named:id=33:Normal"
      _StyleDefs(67)  =   ":id=33,.parent=0"
      _StyleDefs(68)  =   "Named:id=34:Heading"
      _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   ":id=34,.wraptext=-1"
      _StyleDefs(71)  =   "Named:id=35:Footing"
      _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   "Named:id=36:Selected"
      _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=37:Caption"
      _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(77)  =   "Named:id=38:HighlightRow"
      _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=39:EvenRow"
      _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(81)  =   "Named:id=40:OddRow"
      _StyleDefs(82)  =   ":id=40,.parent=33"
      _StyleDefs(83)  =   "Named:id=41:RecordSelector"
      _StyleDefs(84)  =   ":id=41,.parent=34"
      _StyleDefs(85)  =   "Named:id=42:FilterBar"
      _StyleDefs(86)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "終 了"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2016
      TabIndex        =   2
      ToolTipText     =   "処理を終了します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "登 録"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   408
      TabIndex        =   1
      ToolTipText     =   "入荷予定データを登録します"
      Top             =   0
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "読　込"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblDisp_Count 
      Alignment       =   1  '右揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BorderStyle     =   1  '実線
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10080
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "読込み件数"
      Height          =   255
      Index           =   0
      Left            =   8760
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ファイル名"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Menu SHORI_MENU 
      Caption         =   "処理選択"
      Begin VB.Menu SHORI 
         Caption         =   "読込み"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu SHORI 
         Caption         =   "登録"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu SHORI 
         Caption         =   "終了"
         Index           =   2
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "SEM00501"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim ITEM_CATEGORY   As New XArrayDB

Private Const Min_Row% = 1              '最小行数

Dim Max_Row    As Integer               'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 6              '最大列数



Private Const colNO% = 0                '№
Private Const colJGYOBU% = 1            'BU
Private Const colCATEGORY_CODE% = 2     '品名ｶﾃｺﾞﾘｰｺｰﾄﾞ

Private Const colCATEGORY_NAME% = 3     '品名


Private Const colSEI_LOT% = 4           '入荷予定日
Private Const colKOUSU_LOT% = 5         '入荷予定数
Private Const colKOUSU_QTY% = 6         '仕入／支給先


Private EXCEL_DATA  As Variant

Private Const LAST_UPDATE_DAY$ = "[SEM0050] 2011.12.10 17:00"

Private Sub Command1_Click(Index As Integer)

Dim sWk         As String
Dim i           As Long
Dim j           As Long


    Select Case Index



        Case 0          '読込み


            For i = 0 To UBound(EXCEL_DATA)

                If InStr(Trim(Text1.Text), EXCEL_DATA(i)) <> 0 Then
                    Exit For
                End If

            Next i


            If i > UBound(EXCEL_DATA) Then
                MsgBox "EXCELデータとして認識出来ません。ファイル名を再入力してください。"
                Text1.SetFocus
                Exit Sub
            End If

            '取込みﾃﾞｰﾀ表示
            If List_Disp_Proc() Then
                Unload Me
            End If


            If ITEM_CATEGORY.Count(1) > 0 Then
                Command1(1).Enabled = True
                SHORI(1).Enabled = True
            Else
                Command1(1).Enabled = False
                SHORI(1).Enabled = False
            End If




        Case 1          '登録


            If Update_Proc() Then
                Unload Me
            End If



        Case 2          '終了

            Unload Me
    End Select

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'----------------------------------------------------------------------------
'                   Ｋｅｙ Ｄｏｗｎ 前処理
'----------------------------------------------------------------------------
    
    Select Case KeyCode
        Case vbKeyF12
            Unload Me
    End Select
    
    
    

End Sub

Private Sub Form_Load()


Dim c       As String * 128



    'ステータスウィンドウを作成する
    hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
        "[請求システム]品名カテゴリーマスタ登録", Me.hwnd, 0)
    '最後の要素を-1にすると
    '親ウィンドウの全体の幅の残りの幅を
    '自動的に割り当てる
    Call SendMessageAny(hStatusWnd, SB_SIMPLE, 0, -1)



    Show
                                'ログファイル名取り込み
    If GetIni("FILE", "LOGF", "SYS", c) Then
        Beep
        MsgBox "ログファイル名の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    LOG_F = RTrim(c)

                                
                                '事業部取り込み
    If JGYOB_TB_Set(1) Then
        Beep
        MsgBox "事業部の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
                                'EXCEL拡張子
    If GetIni(App.EXEName, "EXCEL", App.EXEName, c) Then
        Beep
        MsgBox "EXCEL拡張子の獲得に失敗しました。処理を中止して下さい。"
        End
    End If
    EXCEL_DATA = Split(Trim(c), ",", -1)



                                '品名ｶﾃｺﾞﾘｰﾏｽﾀＯＰＥＮ
    If ITEM_CATEGORY_Open(BtOpenNomal) Then
        Unload Me
    End If



    SEM00501.Caption = SEM00501.Caption & " " & LAST_UPDATE_DAY


End Sub


Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Text1.Text = Trim(Data.Files(1))

    Command1(0).Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
Dim sts     As Integer
    
    sts = BTRV(BtOpClose, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
    If sts Then
        If sts <> BtErrNoOpen Then
            Call File_Error(sts, BtOpClose, "品目ｶﾃｺﾞﾘｰﾏｽﾀ")
        End If
    End If
    
    sts = BTRV(BtOpReset, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
    If sts Then
        Call File_Error(sts, BtOpReset, "")
    End If


    Set SEM00501 = Nothing



    End

End Sub

Private Sub SHORI_Click(Index As Integer)

    Select Case Index
    
        Case 0
            Command1(0).Value = True
        Case 1
            Command1(1).Value = True
        Case 2
            Command1(2).Value = True
    End Select



End Sub

Private Sub TDBGrid1_OLEDragDrop(ByVal Data As TrueDBGrid80.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Text1.Text = Trim(Data.Files(0))


    Command1(0).Value = True


End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Text1.Text = Trim(Data.Files(1))
    
    Command1(0).Value = True



End Sub

Private Function Update_Proc() As Integer
'----------------------------------------------------------------------------
'                   「品目カテゴリーマスタ」登録処理
'----------------------------------------------------------------------------

Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer
    
Dim INS_NOW         As String * 14
    

Dim Row             As Long

    
    If ITEM_CATEGORY.Count(1) = 0 Then
        Exit Function
    End If



    Update_Proc = True
    
    Call Input_Lock

hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "品名カテゴリーマスタ登録処理　処理開始！！", Me.hwnd, 0)
                                    
    INS_NOW = Format(Now, "YYYYMMDDHHMMSS")
                                    
                                    'テーブルリセット
    
    For Row = 1 To ITEM_CATEGORY.UpperBound(1)
        
        
        DoEvents
        
        Call UniCode_Conv(K0_ITEM_CATEGORY.JGYOBU, ITEM_CATEGORY(Row, colJGYOBU))
        Call UniCode_Conv(K0_ITEM_CATEGORY.CATEGORY_CODE, ITEM_CATEGORY(Row, colCATEGORY_CODE))
        sts = BTRV(BtOpGetEqual, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
        Select Case sts
            Case BtNoErr
            
                com = BtOpUpdate
            
            
            Case BtErrKeyNotFound
            
                com = BtOpInsert
            
            Case Else
                Call Input_UnLock
                Call File_Error(sts, BtOpGetEqual, "品名ｶﾃｺﾞﾘｰﾏｽﾀ")
                Exit Function
        End Select
        
        If com = BtOpInsert Then
                
            Call UniCode_Conv(ITEM_CATEGORYREC.JGYOBU, ITEM_CATEGORY(Row, colJGYOBU))
            Call UniCode_Conv(ITEM_CATEGORYREC.CATEGORY_CODE, ITEM_CATEGORY(Row, colCATEGORY_CODE))
        
        
            Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_QTY, "0000000000")
            Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_KOURYO, "0000000000.00")
            Call UniCode_Conv(ITEM_CATEGORYREC.TOKU_TANKA_HAKO, "0000000000.00")
        
            Call UniCode_Conv(ITEM_CATEGORYREC.MEMO, "")
            
            Call UniCode_Conv(ITEM_CATEGORYREC.FILLER, "")
        
            Call UniCode_Conv(ITEM_CATEGORYREC.INS_TANTO, App.EXEName)
            Call UniCode_Conv(ITEM_CATEGORYREC.Ins_DateTime, INS_NOW)
                    
            Call UniCode_Conv(ITEM_CATEGORYREC.UPD_TANTO, "")
            Call UniCode_Conv(ITEM_CATEGORYREC.UPD_DATETIME, "")
                    
        
        Else
        
            Call UniCode_Conv(ITEM_CATEGORYREC.UPD_TANTO, App.EXEName)
            Call UniCode_Conv(ITEM_CATEGORYREC.UPD_DATETIME, INS_NOW)
        
        End If
        
        
        Call UniCode_Conv(ITEM_CATEGORYREC.CATEGORY_NAME, Trim(ITEM_CATEGORY(Row, colCATEGORY_NAME)))

        If Len(ITEM_CATEGORY(Row, colSEI_LOT)) < 10 Then
            Call UniCode_Conv(ITEM_CATEGORYREC.SEI_LOT, String(10 - Len(ITEM_CATEGORY(Row, colSEI_LOT)), "0") & ITEM_CATEGORY(Row, colSEI_LOT))
        Else
            Call UniCode_Conv(ITEM_CATEGORYREC.SEI_LOT, ITEM_CATEGORY(Row, colSEI_LOT))
        End If


        If Len(ITEM_CATEGORY(Row, colKOUSU_LOT)) < 10 Then
            Call UniCode_Conv(ITEM_CATEGORYREC.KOUSU_LOT, String(10 - Len(ITEM_CATEGORY(Row, colKOUSU_LOT)), "0") & ITEM_CATEGORY(Row, colKOUSU_LOT))
        Else
            Call UniCode_Conv(ITEM_CATEGORYREC.KOUSU_LOT, ITEM_CATEGORY(Row, colKOUSU_LOT))
        End If

        If Len(ITEM_CATEGORY(Row, colKOUSU_QTY)) < 10 Then
            Call UniCode_Conv(ITEM_CATEGORYREC.KOUSU_QTY, String(10 - Len(ITEM_CATEGORY(Row, colKOUSU_QTY)), "0") & ITEM_CATEGORY(Row, colKOUSU_QTY))
        Else
            Call UniCode_Conv(ITEM_CATEGORYREC.KOUSU_QTY, ITEM_CATEGORY(Row, colKOUSU_QTY))
        End If
                    
        Do
            sts = BTRV(com, ITEM_CATEGORY_POS, ITEM_CATEGORYREC, Len(ITEM_CATEGORYREC), K0_ITEM_CATEGORY, Len(K0_ITEM_CATEGORY), 0)
            Select Case sts
                Case BtNoErr
                    Exit Do
                
                Case BtErrRECORD_INUSE, BtErrFILE_INUSE
                    
                    Beep
                    ans = MsgBox("「品目カテゴリーマスタ」他端末でデータ使用中です。<ITEM_CATEGORY.DAT>", vbRetryCancel + vbQuestion, "確認入力")
                    If ans = vbCancel Then
                        Call Input_UnLock
                        Exit Function
                    End If
                
                Case Else
                    Call Input_UnLock
                    Call File_Error(sts, BtOpInsert, "品目ｶﾃｺﾞﾘﾏｽﾀ")
                    Exit Function
            End Select
        
        Loop
            
    
        Set TDBGrid1.Array = ITEM_CATEGORY
        TDBGrid1.ReBind
        
        TDBGrid1.Update
        TDBGrid1.Bookmark = Row
        

    Next Row


    Set TDBGrid1.Array = ITEM_CATEGORY
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "品名カテゴリーマスタ登録処理　処理終了！！", Me.hwnd, 0)





    Update_Proc = False
    Call Input_UnLock
    Exit Function

Error_Proc:
    
    MsgBox "Err.Number= " & Err.Number & " " & Err.Description
    Call Input_UnLock

End Function

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   「品名カテゴリーマスタ」読込み処理
'----------------------------------------------------------------------------
Dim sts                 As Integer
Dim ans                 As Integer
    

Dim xlApp               As Object
Dim xlBook              As Object
Dim xlSheet             As Object

Dim Row                 As Long

Dim END_GYO             As Integer
Dim SKIP_F              As Boolean


Dim wkJGOYBU            As String * 1
Dim wkCATEGORY_CODE     As String * 8
Dim wkCATEGORY_NAME     As String * 80

Dim wkSEI_LOT           As String * 10
Dim wkKOUSU_LOT         As String * 10
Dim wkKOUSU_QTY         As String * 10




Dim i               As Long
Dim j               As Long
Dim k               As Long

    List_Disp_Proc = True

    Call Input_Lock







    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlApp.Worksheets(1)
    
    
    On Error GoTo Error_Proc

    '2011.12.03
    xlApp.Workbooks.Open (Text1.Text), ReadOnly:=True, UpdateLinks:=0

    On Error GoTo 0



hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "品名カテゴリーＥＸＣＥＬデータ　表示処理開始！！", Me.hwnd, 0)


    Set ITEM_CATEGORY = Nothing
    
    Row = Min_Row - 1
    lblDisp_Count.Caption = ""



    For j = 1 To xlApp.Worksheets.Count
    
        Set xlSheet = xlApp.Worksheets(j)
        xlSheet.Activate


        END_GYO = 0
        For i = 1 To 1048576
            
            SKIP_F = False
            
            
            On Error GoTo Error_Proc        '2011.12.03
    
                
            If Trim(xlSheet.Application.Cells(i, 1)) = "" And _
                Trim(xlSheet.Application.Cells(i, 2)) = "" And _
                Trim(xlSheet.Application.Cells(i, 3)) = "" And _
                Trim(xlSheet.Application.Cells(i, 4)) = "" And _
                Trim(xlSheet.Application.Cells(i, 5)) = "" Then
            
                SKIP_F = True
                END_GYO = END_GYO + 1
                
                If END_GYO > 5 Then
                    Exit For
                End If
            Else
                
                END_GYO = 0
        
                If Trim(xlSheet.Application.Cells(i, 1)) = "" Or _
                    Trim(xlSheet.Application.Cells(i, 2)) = "" Then
            
                    SKIP_F = True
            
                Else
                    
                               
                    For k = 0 To UBound(JGYOBU_T)
                    
                        If Trim(JGYOBU_T(k).CODE) = Trim(xlSheet.Application.Cells(i, 1)) Then
                            Exit For
                        End If
                    
                    Next k
                    
                    If k > UBound(JGYOBU_T) Then
                        SKIP_F = True
                    End If
                    
                    '事業部
                    wkJGOYBU = Trim(xlSheet.Application.Cells(i, 1))
                    '品名ｶﾃｺﾞﾘｰｺｰﾄﾞ
                    wkCATEGORY_CODE = Trim(xlSheet.Application.Cells(i, 2))
                    '品名ｶﾃｺﾞﾘｰ名称
                    wkCATEGORY_NAME = Trim(xlSheet.Application.Cells(i, 3))
                    '見積ﾛｯﾄ数
                    wkSEI_LOT = Trim(xlSheet.Application.Cells(i, 4))
                    If Not IsNumeric(wkSEI_LOT) Then
                        wkSEI_LOT = "0"
                    End If
                    '前後工数(秒/ﾛｯﾄ)
                    wkKOUSU_LOT = Trim(xlSheet.Application.Cells(i, 5))
                    If Not IsNumeric(wkKOUSU_LOT) Then
                        wkKOUSU_LOT = "0"
                    End If
                    '前後工数(秒/個)
                    wkKOUSU_QTY = Trim(xlSheet.Application.Cells(i, 6))
                    If Not IsNumeric(wkKOUSU_QTY) Then
                        wkKOUSU_QTY = "0"
                    End If
                
                    If Not SKIP_F Then
                
                        Row = Row + 1
                        ITEM_CATEGORY.ReDim Min_Row, Row, Min_Col, Max_Col
                    
                        ITEM_CATEGORY(Row, colNO) = Row
    
                        ITEM_CATEGORY(Row, colJGYOBU) = wkJGOYBU
                        ITEM_CATEGORY(Row, colCATEGORY_CODE) = wkCATEGORY_CODE
                        ITEM_CATEGORY(Row, colCATEGORY_NAME) = wkCATEGORY_NAME
                        ITEM_CATEGORY(Row, colSEI_LOT) = Val(wkSEI_LOT)
                        ITEM_CATEGORY(Row, colKOUSU_LOT) = Val(wkKOUSU_LOT)
                        ITEM_CATEGORY(Row, colKOUSU_QTY) = Val(wkKOUSU_QTY)
    
                    End If
                End If
            
            End If
    
    
    
        Next i

    Next j


    On Error GoTo 0                 '2011.12.03


    Set TDBGrid1.Array = ITEM_CATEGORY
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst






    lblDisp_Count.Caption = Format(Row, "#0") & "件"






hStatusWnd = CreateStatusWindow(WS_VISIBLE Or WS_CHILD Or CCS_BOTTOM Or SBARS_SIZEGRIP, _
    "品名カテゴリーＥＸＣＥＬデータ　表示処理終了！！", Me.hwnd, 0)



    Call Input_UnLock

    xlBook.Close False
    xlApp.Quit 'EXCELを閉じる
    Set xlApp = Nothing

    List_Disp_Proc = False
    Exit Function

Error_Proc:
    

    Select Case Err.Number
        
        '52 ファイル名または番号が不正です。
        '53 ファイルが見つかりません。
        '54 ファイル モードが不正です。
        '55 ファイルは既に開かれています。
        '57 デバイス I/O エラーです。
        '59 レコード長が一致しません。
        '61 ディスクの空き容量が不足しています。
        '62 ファイルにこれ以上データがありません。
        '63 レコード番号が不正です。
        '68 デバイスが準備されていません。
        '70 書き込みできません。
        '71 ディスクが準備されていません。
        '75 パス名が無効です。
        '76 パスが見つかりません。
        Case 52, 53, 54, 55, 57, 59, 61, 62, 63, 68, 70, 71, 75, 76
            
            
            MsgBox "指定のファイルが見つかりません。" & Chr(13) & Chr(10) & "正しいファイル名を入力してください。"
            
            
            
            List_Disp_Proc = False      '


    '2011.12.03
        Case 13
        
            MsgBox "読込み対象のEXCELデータに異常が有ります。内容を確認後、再実行してください。"
            
            xlApp.DisplayAlerts = False
            xlBook.Close False
            xlApp.Quit 'EXCELを閉じる
            Set xlApp = Nothing
            
            
            
            
            
            
            List_Disp_Proc = False      '
            


        Case Else
            MsgBox Err.Description
    
    '2011.12.03
    End Select
    
    On Error GoTo 0
    
    
    Call Input_UnLock

End Function






Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    SEM00501.MousePointer = vbHourglass

    Call Ctrl_Lock(SEM00501)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(SEM00501)


    SEM00501.MousePointer = vbDefault

End Sub


