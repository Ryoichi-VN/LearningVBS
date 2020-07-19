
' Note
' 「Ref1.xlsx」の表から、情報を取り出し、
' 行ごとに文字列を別のファイルに書き出す

Sub Main ()
  'On Error Resume Next
 
  Dim objXls, objRefBook, objRefSheet, objRefRange, _
  			  headstring, bodystring
  'objTempBook, objTempSheet, objTempRange
  Dim objWkBk, objWkSt
  Set objXls = CreateObject("Excel.Application")
  'If Not objXls Then Exit Sub
  If objXls Is Nothing Then Exit Sub
  objXls.Visible = True
  objXls.ScreenUpdating = True
  
  ' 参照先のブックとシートをopen
  Set objRefBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\Ref1.xlsx")
  Set objRefSheet = objRefBook.Sheets("Sheet1")
  ' テンプレート先のブックとシートをopen
  'Set objTempBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\tmp1.xlsx")
  'Set objTempSheet = objTempBook.Sheets("Sheet1")
 
   ' データ取得範囲を指定
  'Set objRefRange = objSheet.Range("A1:E2") 
   ' データのある範囲を取得
   ' データ範囲は明治指定したほうがよい。。。
  Set objRefRange = objRefSheet.UsedRange
  
    ' 行ごとのデータを取得
  headstring = "output"
  
  For intR = 2 To objRefRange.Rows.Count
    bodystring = ""
    
    For intC = 1 To objRefRange.Columns.Count
      bodystring = bodystring & CStr(objRefRange(objRefRange.Row + intR -1 , objRefRange.Column + intC - 1))
    Next
    
    ' Workbookを新規作成
    Set objWkBk = objXls.Workbooks.Add()
    Set objWkSt = objWkBk.Sheets("Sheet1")
    
     ' 文字列書き込み
    objWkSt.Cells(1,2).Value = bodystring
    objWkSt.Cells(1,2).NumberFormatLocal = "@"
    
    ' 同盟のファイルが既に存在する場合は上書き
    '一時的にアラートを強制解除
    objXls.DisplayAlerts = False
    ' Workbookを保存
    objWkSt.SaveAs(GetCurrentDirectory() & "\output" & intR - 1 & ".xlsx")
    ' Workbookを閉じる
    objWkBk.Close
    'アラートを規定設定に修正
    objXls.DisplayAlerts = True
    
    
  Next
  
    
  
  'objRefBook.Save
 
  objRefBook.Close
  objXls.Quit
  Set objRefBook = Nothing
  Set objRefBook = Nothing
  Set objXls = Nothing
End Sub
 
Function GetCurrentDirectory()
  On Error Resume Next
  Dim objShell : Set objShell = CreateObject("WScript.Shell")
  GetCurrentDirectory = objShell.CurrentDirectory
End Function
 
Main
Msgbox "完了"