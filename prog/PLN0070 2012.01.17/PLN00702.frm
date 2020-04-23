VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form PLN00702 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "[商品化計画システム]資材所要量使用親品番確認画面"
   ClientHeight    =   5508
   ClientLeft      =   1992
   ClientTop       =   -4800
   ClientWidth     =   4452
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
   MinButton       =   0   'False
   OLEDropMode     =   1  '手動
   ScaleHeight     =   5508
   ScaleWidth      =   4452
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Begin TrueDBGrid80.TDBGrid TDBGrid1 
      Height          =   4212
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3972
      _ExtentX        =   7006
      _ExtentY        =   7430
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "品番"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "商品化予定数"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   720
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2879"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2773"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2498"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2392"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8194"
      Splits(0)._ColumnProps(10)=   "Column(1).FetchStyle=1"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
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
      _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=74,.parent=9,.bgcolor=&HFFFFFF&"
      _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=75,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=102,.parent=67,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=68"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=69"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=71"
      _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=16,.parent=67,.alignment=1,.locked=-1"
      _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=68"
      _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=69"
      _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=71"
      _StyleDefs(46)  =   "Named:id=33:Normal"
      _StyleDefs(47)  =   ":id=33,.parent=0"
      _StyleDefs(48)  =   "Named:id=34:Heading"
      _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   ":id=34,.wraptext=-1"
      _StyleDefs(51)  =   "Named:id=35:Footing"
      _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(53)  =   "Named:id=36:Selected"
      _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=37:Caption"
      _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(57)  =   "Named:id=38:HighlightRow"
      _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=39:EvenRow"
      _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(61)  =   "Named:id=40:OddRow"
      _StyleDefs(62)  =   ":id=40,.parent=33"
      _StyleDefs(63)  =   "Named:id=41:RecordSelector"
      _StyleDefs(64)  =   ":id=41,.parent=34"
      _StyleDefs(65)  =   "Named:id=42:FilterBar"
      _StyleDefs(66)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "閉じる"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.6
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      ToolTipText     =   "商品化構成を保存します"
      Top             =   4920
      Width           =   1380
   End
   Begin VB.Label lblYOTEI_DT 
      Appearance      =   0  'ﾌﾗｯﾄ
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   1800
      TabIndex        =   3
      Top             =   240
      Width           =   1332
   End
   Begin VB.Label Label1 
      Caption         =   "商品化予定日"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1572
   End
End
Attribute VB_Name = "PLN00702"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit






Dim ZAIKO_RIREKI    As New XArrayDB



Private Const Min_Row% = 1              '最小行数
Dim Max_Row    As Long                  'グリッド最大表示件数


Private Const Min_Col% = 0              '最小列数
Private Const Max_Col% = 1              '最大列数

Private Const colHIN_GAI% = 0           '品番
Private Const colYOTEI_QTY% = 1         '予定数









Private Sub Command1_Click(Index As Integer)

    Select Case Index


        Case 0          '読込み
            
            PLN00702.Visible = False

    End Select





End Sub

Private Sub Form_Activate()

    If List_Disp_Proc() Then
        PLN00702.Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    
    Select Case KeyCode
        
        Case vbKeyF12
            PLN00702.Visible = False
    End Select
    
    
    

End Sub

Private Function List_Disp_Proc() As Integer
'----------------------------------------------------------------------------
'                   「資材所要量使用親品番確認画面」表示処理
'----------------------------------------------------------------------------
Dim sts             As Integer
Dim ans             As Integer
Dim com             As Integer


Dim Row             As Long

Dim SKIP_FLG        As Integer


    List_Disp_Proc = True

    Call Input_Lock


    lblYOTEI_DT = DISP_DATE


    Call UniCode_Conv(K1_PLN_tmpP_COMP.YOTEI_DT, Format(DISP_DATE, "YYYYMMDD"))
    
    Call UniCode_Conv(K1_PLN_tmpP_COMP.KO_SYUBETSU, DISP_KO_SYUBETSU_CODE)
    Call UniCode_Conv(K1_PLN_tmpP_COMP.KO_JGYOBU, DISP_KO_JGYOBU)
    Call UniCode_Conv(K1_PLN_tmpP_COMP.KO_NAIGAI, "1")
    Call UniCode_Conv(K1_PLN_tmpP_COMP.KO_HIN_GAI, DISP_KO_HIN_GAI)
    Call UniCode_Conv(K1_PLN_tmpP_COMP.JGYOBU, "")
    Call UniCode_Conv(K1_PLN_tmpP_COMP.NAIGAI, "")
    Call UniCode_Conv(K1_PLN_tmpP_COMP.HIN_GAI, "")
    
    



    Set ZAIKO_RIREKI = Nothing
    Row = Min_Row - 1

    com = BtOpGetGreater
    
    Do
        DoEvents
        sts = BTRV(com, PLN_tmpP_COMP_POS, PLN_tmpP_COMP_REC, Len(PLN_tmpP_COMP_REC), K1_PLN_tmpP_COMP, Len(K1_PLN_tmpP_COMP), 1)
        Select Case sts
            Case BtNoErr
            Case BtErrEOF
                Exit Do
            Case Else
                
                Call File_Error(sts, com, "資材所要量中間ファイル")
                Exit Function
        
        End Select
    
    
        If StrConv(PLN_tmpP_COMP_REC.YOTEI_DT, vbUnicode) <> Format(DISP_DATE, "YYYYMMDD") Or _
            Trim(StrConv(PLN_tmpP_COMP_REC.KO_SYUBETSU, vbUnicode)) <> Trim(DISP_KO_SYUBETSU_CODE) Or _
            StrConv(PLN_tmpP_COMP_REC.KO_JGYOBU, vbUnicode) <> DISP_KO_JGYOBU Or _
            StrConv(PLN_tmpP_COMP_REC.KO_NAIGAI, vbUnicode) <> "1" Or _
            Trim(StrConv(PLN_tmpP_COMP_REC.KO_HIN_GAI, vbUnicode)) <> Trim(DISP_KO_HIN_GAI) Then
            
            Exit Do
        
        End If
    
        Row = Row + 1
        ZAIKO_RIREKI.ReDim Min_Row, Row, Min_Col, Max_Col
    
        ZAIKO_RIREKI(Row, colHIN_GAI) = StrConv(PLN_tmpP_COMP_REC.HIN_GAI, vbUnicode)
        ZAIKO_RIREKI(Row, colYOTEI_QTY) = Format(Val(StrConv(PLN_tmpP_COMP_REC.YOTEI_QTY, vbUnicode)), "#,##0")
    
        com = BtOpGetNext
    
    
    Loop
    
    

    Set TDBGrid1.Array = ZAIKO_RIREKI
    TDBGrid1.ReBind
    
    TDBGrid1.Update
    TDBGrid1.MoveFirst

    Call Input_UnLock


    List_Disp_Proc = False

End Function


Private Sub Input_Lock()
'----------------------------------------------------------------------------
'                   画面項目ロック（イベント取得不可）
'----------------------------------------------------------------------------
Dim i   As Integer


    PLN00702.MousePointer = vbHourglass

    Call Ctrl_Lock(PLN00702)



End Sub

Private Sub Input_UnLock()
'----------------------------------------------------------------------------
'                   画面項目ロック解除（イベント取得可）
'----------------------------------------------------------------------------
Dim i   As Integer

    Call Ctrl_UnLock(PLN00702)


    PLN00702.MousePointer = vbDefault

End Sub




