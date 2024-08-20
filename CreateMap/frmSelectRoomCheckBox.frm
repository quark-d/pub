VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectRoomCheckBox 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSelectRoomCheckBox.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSelectRoomCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim selectedItems As New Collection
    Dim ctrl As Control
    Dim ws As Worksheet
    
    ' 選択されたチェックボックスを確認
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" And ctrl.value = True Then
            selectedItems.Add ctrl.Caption
        End If
    Next ctrl
    
    ' チェックされた項目がない場合の処理
    If selectedItems.Count = 0 Then
        MsgBox "項目が選択されていません。", vbExclamation
        Exit Sub
    End If
    
   
    ' Collectionに格納された選択項目の処理
    Call ResetAllSheetTabColors
    
    Dim item As Variant
    For Each item In selectedItems
        ' 該当するシートのタブを SetSheetTabColor で青にする (ColorIndex = 5)
        For Each ws In ThisWorkbook.Sheets
            If ws.Name = item Then
                Call SetSheetTabColor(ws, 5) ' 青色に設定
                Exit For
            End If
        Next ws
    Next item
    
    ' 青色に設定されたシートのみを印刷
    Call ExportSheetsWithColorTabsToPDF(5)
    
    ' フォームを閉じる
    Me.Hide
End Sub



Private Sub UserForm_Initialize()
    Dim idAddrDict As Scripting.Dictionary
    Dim shtNameCollection As Collection
    
    Dim shtName As Variant
    Dim topPosition As Integer
    Dim chkBox As MSForms.CheckBox
    Dim formHeight As Integer
    Dim spacing As Integer
    
    ' チェックボックスの間隔
    spacing = 20
    
    ' ID辞書を取得する
    Set idAddrDict = GetIdAddrDictionary() ' IDを取得する関数を呼び出す
    Set shtNameCollection = GetSheetNamesFromAddrDict(idAddrDict)
    
    ' チェックボックスの配置開始位置
    topPosition = 10
    
    ' シート名ごとにチェックボックスを作成
    For Each shtName In shtNameCollection
        Set chkBox = Me.Controls.Add("Forms.CheckBox.1")
        With chkBox
            .Caption = shtName
            .left = 10
            .top = topPosition
            .width = 150
            topPosition = topPosition + spacing
        End With
    Next shtName
    
    ' フォームの高さを調整
    formHeight = topPosition + 40 ' チェックボックスの最後の位置に余白を追加
    Me.height = formHeight
    
    ' イベント処理を行うために少し待つ
    DoEvents
End Sub


