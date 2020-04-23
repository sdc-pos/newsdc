Attribute VB_Name = "MainPI00041"
Option Explicit
Public RYOHEN      As String * 2       '良品返品の要因 2009.07.10

Public GLB_SYUSHI_F     As String


Public Const WEL_MAEGARI_TANA_S_OSAKA$ = "H2"       '「WEL 資材前借入庫」の要因 2016.06.17


Sub Main()
    
    
    
    
    GLB_SYUSHI_F = Trim(Command)


    PI000411.Show
End Sub
