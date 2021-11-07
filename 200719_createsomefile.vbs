
' Note
' 「Ref1.xlsx」の表から、情報を取り出す
' 「tmp1.xlsx」の形式で特定のセルに行ごとに文字列を別のファイルに書き出す


Sub Main ()
  
  'On Error Resume Next
 
  Dim objXls ' Object定義:エクセル
  Dim objRefBook, objRefSheet, objRefRange 		' Object定義:refference
  Dim objTmpBook, objTmpSheet, objTempRange 	' Object定義:refference
  Dim headstring, bodystring 	' work用文字列
  Dim objWkBk, objWkSt			' work用オブジェクト
  
  Set objXls = CreateObject("Excel.Application")
  'If Not objXls Then Exit Sub
  If objXls Is Nothing Then Exit Sub
  objXls.Visible = True
  objXls.ScreenUpdating = True
  
  ' 参照先のブックとシートをopen
  Set objRefBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\Ref1.xlsx")
  Set objRefSheet = objRefBook.Sheets("Sheet1")
  ' テンプレート先のブックとシートをopen
  Set objTmpBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\tmp1.xlsx")
  Set objTmpSheet = objTmpBook.Sheets("temp1")
 
   ' データ取得範囲を指定
  'Set objRefRange = objSheet.Range("A1:E2") 
   ' データのある範囲を取得
   ' データ範囲は明示指定したほうがよい。。。
  Set objRefRange = objRefSheet.UsedRange
  
    ' 行ごとのデータを取得
  headstring = "output"
  'Msgbox "range:" & CStr(objRefRange(objRefRange.Row + objRefRange.Rows.Count -1 , objRefRange.Column + objRefRange.Columns.Count - 1))
  
  For intR = 2 To objRefRange.Rows.Count
    bodystring = ""
    
    For intC = 1 To objRefRange.Columns.Count
      bodystring = bodystring & CStr(objRefRange(objRefRange.Row + intR -1 , objRefRange.Column + intC - 1))
    Next
    'Msgbox bodystring
    
    ' Workbookを新規作成
    Set objWkBk = objXls.Workbooks.Add()
    'Set objWkSt = objWkBk.Sheets("Sheet1")
    objWkBk.Sheets("Sheet1").Name = ("dummy")
    'Msgbox "cre_dummy"
    
    ' テンプレートシートから生成シートに内容をコピー
    objTmpSheet.Copy(objWkBk.Sheets("dummy"))
    
    objXls.DisplayAlerts = False      ' データがあるとアラート表示されるので一時的に消す。
    objWkBk.Sheets("dummy").Delete
    objXls.DisplayAlerts = True
    'Msgbox "tempcopied"
    ' 文字列書き込み
    objWkBk.Sheets("temp1").Cells(4,3).Value = bodystring
    objWkBk.Sheets("temp1").Cells(4,3).NumberFormatLocal = "@"
    
    ' 同名のファイルが既に存在する場合は上書き
    '一時的にアラートを強制解除
    objXls.DisplayAlerts = False
    ' Workbookを保存
    objWkBk.SaveAs(GetCurrentDirectory() & "\output" & intR - 1 & ".xlsx")
    ' Workbookを閉じる
    objWkBk.Close
    Set objWkBk = Nothing
    'アラートを規定設定に修正
    objXls.DisplayAlerts = True
    
    
  Next
  
    
  
  'objRefBook.Save
 
  objRefBook.Close
  objTmpBook.Close
  objXls.Quit
  Set objRefBook = Nothing
  Set objTmpBook = Nothing
  Set objXls = Nothing
End Sub
 
Function GetCurrentDirectory()
  On Error Resume Next
  Dim objShell : Set objShell = CreateObject("WScript.Shell")
  GetCurrentDirectory = objShell.CurrentDirectory
End Function
 
Main
Msgbox "完了"