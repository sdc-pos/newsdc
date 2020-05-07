VERSION 5.00
Begin {2AF752CD-B826-4828-B4C1-13BFC9CC90C1} PR00026F1 
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows の既定値
   _ExtentX        =   26882
   _ExtentY        =   19420
   SectionData     =   "PR00026F1.dsx":0000
End
Attribute VB_Name = "PR00026F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DET_com         As Integer      '明細のBtrieve Operation


'計上年月用添字
Private Const ptxKEIJYO_YM% = 0         '対象年月






Private Sub ActiveReport_DataInitialize()

    Me.Fields.Add "SHIIRE_NAME"         '仕入先名称（ｺｰﾄﾞ含む）
    
    Me.Fields.Add "SHIIRE01"             '資材
    Me.Fields.Add "SHIIRE02"             '工料
    Me.Fields.Add "SHIIRE03"             '家賃
    Me.Fields.Add "SHIIRE04"             'その他
    Me.Fields.Add "SHIIRE05"             '小計
    Me.Fields.Add "SHIIRE06"             '派遣工料
    Me.Fields.Add "SHIIRE07"             '合計
    Me.Fields.Add "SHIIRE08"             '経費仕入
    Me.Fields.Add "SHIIRE09"             '仕入合計
    Me.Fields.Add "SHIIRE10"             '消費税
    Me.Fields.Add "SHIIRE11"             '総合計
    


End Sub

Private Sub ActiveReport_FetchData(eof As Boolean)

Dim sts         As Integer

Dim TOTAL       As Long
Dim ZEI         As Long
 
Dim YMD         As String * 8
    
    
    sts = BTRV(DET_com, P_SHSYU_SUM_POS, P_SHSYU_SUM_REC, Len(P_SHSYU_SUM_REC), K0_P_SHSYU_SUM, Len(K0_P_SHSYU_SUM), 0)
    Select Case sts
        Case BtNoErr
        Case BtErrEOF
            Exit Sub
        Case Else
            Call File_Error(sts, DET_com, "資材仕入集計ﾃﾞｰﾀ")
            Exit Sub
    End Select
    '収支単位
    
    If StrConv(P_SHSYU_SUM_REC.G_SYUSHI, vbUnicode) = "zzz" Then
        Me.Fields("SHIIRE_NAME") = "　　総　合　計　　"
    Else
        Call UniCode_Conv(K0_P_CODE.DATA_KBN, P_KBN03_CD)
        Call UniCode_Conv(K0_P_CODE.C_Code, StrConv(P_SHSYU_SUM_REC.G_SYUSHI, vbUnicode))
        sts = BTRV(BtOpGetEqual, P_CODE_POS, P_CODEREC, Len(P_CODEREC), K0_P_CODE, Len(K0_P_CODE), 0)
        Select Case sts
            Case BtNoErr
                Me.Fields("SHIIRE_NAME") = StrConv(P_SHSYU_SUM_REC.G_SYUSHI, vbUnicode) & " " & StrConv(P_CODEREC.C_RNAME, vbUnicode)
            Case BtErrKeyNotFound
                Me.Fields("SHIIRE_NAME") = StrConv(P_SHSYU_SUM_REC.G_SYUSHI, vbUnicode)
            Case Else
                Call File_Error(sts, BtOpGetEqual, "コードﾏｽﾀ")
                Exit Sub
        End Select
    End If

    Me.Fields("SHIIRE01") = Format(CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)), "#,##0")
    Me.Fields("SHIIRE02") = Format(CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)), "#,##0")
    Me.Fields("SHIIRE03") = Format(CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)), "#,##0")
    Me.Fields("SHIIRE04") = Format(CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode)), "#,##0")
    
    
        
    TOTAL = CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)) + _
            CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)) + _
            CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)) + _
            CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode))

        
    Me.Fields("SHIIRE05") = Format(TOTAL, "#,##0")

    Me.Fields("SHIIRE06") = Format(CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode)), "#,##0")
        
    TOTAL = TOTAL + CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode))

    Me.Fields("SHIIRE07") = Format(TOTAL, "#,##0")
        
    Me.Fields("SHIIRE08") = Format(CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode)), "#,##0")

    TOTAL = TOTAL + CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode))

    Me.Fields("SHIIRE09") = Format(TOTAL, "#,##0")
    


'--------------------------------------------------------------------------
    Me.Fields("SHIIRE10") = ""

    Me.Fields("SHIIRE11") = Format(TOTAL, "#,##0")

    
    
    DET_com = BtOpGetNext
    
            
    eof = False
    
    

End Sub

Private Sub ActiveReport_Initialize()

Dim sts         As Integer
Dim SKIP_Flg    As Boolean
 
Dim TOTAL       As Long
Dim ZEI         As Long
 
Dim YMD         As String * 8
 
    Label1.Caption = PR000261.Text1(ptxKEIJYO_YM).Text                  '計上年月
    Label5.Caption = Trim(StrConv(P_KANRIREC.CENTER_NAME, vbUnicode))   'センター名


'''    '合計ﾚｺｰﾄﾞ（一般仕入分）の読み込み
'''    Call UniCode_Conv(K0_P_SHSYU_SUM.G_SYUSHI, "zzz")
'''
'''    SKIP_Flg = False
'''
'''    sts = BTRV(BtOpGetEqual, P_SHSYU_SUM_POS, P_SHSYU_SUM_REC, Len(P_SHSYU_SUM_REC), K0_P_SHSYU_SUM, Len(K0_P_SHSYU_SUM), 0)
'''    Select Case sts
'''        Case BtNoErr
'''        Case BtErrKeyNotFound
'''
'''            SKIP_Flg = True
'''
'''        Case Else
'''            Call File_Error(sts, BtOpGetEqual, "資材仕入集計ﾃﾞｰﾀ")
'''            Exit Sub
'''    End Select
'''
'''
'''    If SKIP_Flg Then
'''        G_SHIIRE01.Text = "0"
'''        G_SHIIRE02.Text = "0"
'''        G_SHIIRE03.Text = "0"
'''        G_SHIIRE04.Text = "0"
'''        G_SHIIRE05.Text = "0"
'''        G_SHIIRE06.Text = "0"
'''        G_SHIIRE07.Text = "0"
'''        G_SHIIRE08.Text = "0"
'''        G_SHIIRE09.Text = "0"
'''        G_SHIIRE10.Text = "0"
'''        G_SHIIRE11.Text = "0"
'''
'''    Else
'''
'''        G_SHIIRE01.Text = Format(CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)), "#,##0")
'''        G_SHIIRE02.Text = Format(CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)), "#,##0")
'''        G_SHIIRE03.Text = Format(CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)), "#,##0")
'''        G_SHIIRE04.Text = Format(CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode)), "#,##0")
'''
'''        TOTAL = CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(0).SHIIRE, vbUnicode)) + _
'''                CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(1).SHIIRE, vbUnicode)) + _
'''                CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(2).SHIIRE, vbUnicode)) + _
'''                CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(3).SHIIRE, vbUnicode))
'''
'''        G_SHIIRE05.Text = Format(TOTAL, "#,##0")
'''
'''        G_SHIIRE06.Text = Format(CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode)), "#,##0")
'''
'''        TOTAL = TOTAL + CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(4).SHIIRE, vbUnicode))
'''
'''        G_SHIIRE07.Text = Format(TOTAL, "#,##0")
'''
'''
'''        G_SHIIRE08.Text = Format(CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode)), "#,##0")
'''
'''        TOTAL = TOTAL + CLng(StrConv(P_SHSYU_SUM_REC.SHIIRE_TBL(5).SHIIRE, vbUnicode))
'''
'''        G_SHIIRE09.Text = Format(TOTAL, "#,##0")
'''
'''        SHIIRE10.Text = ""
'''
'''
'''        SHIIRE11.Text = Format(TOTAL, "#,##0")
'''
'''        sts = BTRV(BtOpDelete, P_SHSYU_SUM_POS, P_SHSYU_SUM_REC, Len(P_SHSYU_SUM_REC), K0_P_SHSYU_SUM, Len(K0_P_SHSYU_SUM), 0)
'''        Select Case sts
'''            Case BtNoErr
'''
'''            Case Else
'''                Call File_Error(sts, BtOpGetEqual, "資材仕入集計ﾃﾞｰﾀ")
'''                Exit Sub
'''        End Select
'''
'''
'''    End If





    DET_com = BtOpGetFirst


End Sub


Private Sub ActiveReport_ReportStart()
    
    With Me.Printer
        .TrackDefault = False
        .PaperSize = 9
        
        .Orientation = vbPRORLandscape
        .PaperBin = vbPRBNCassette
    End With
    
    
    
    Me.PageBottomMargin = 25
    Me.PageTopMargin = 25
    Me.PageLeftMargin = 0
    Me.PageRightMargin = 0

    Me.documentName = "収支単位別仕入集計表："

    DoEvents

End Sub

