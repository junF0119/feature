Attribute VB_Name = "util00_バー表示"
Option Explicit
' --------------------------------------------------------------------------------
' | @function   EXCELのステータスバーに進行状況を表示
' |-------------------------------------------------------------------------------
' | @moduleName: バー表示
' | @written 2021/12/11
' | @author Jun Fujinawa
' | @license N/A
' | @事前設定
' |
' |
' --------------------------------------------------------------------------------
'                                       +
Public Sub バー表示(ByVal P_statusMsg As String)
'      EXCELのステータスバーに進行状況を表示

        Application.StatusBar = Format(Now(), "m/d hh:mm") & "：" & P_statusMsg & " を処理中です。"
End Sub
