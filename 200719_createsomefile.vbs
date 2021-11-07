
' Note & Hist
' v1：「Ref1.xlsx」の表から、情報を取り出す
' v2：「tmp1.xlsx」の形式で特定のセルに行ごとに文字列を別のファイルに書き出す
' v3：文字列生成にはセル内の改行コードを埋め込んで改行表示対応
' v4：ファイルの生成単位をレコードのキー値の条件で判定する(使用ファイル：Ref2,tmp1)
'
' Task(リファクタリング)
' ・セルの値を定数(constant値)で管理
' ・カラム単位で情報取得　使用する列番号を定数で固定


Sub Main ()
  
  'On Error Resume Next
 
  Dim objXls									' Object定義:エクセル
  Dim objRefBook, objRefSheet, objRefRange 		' Object定義:Ref2
  Dim objTmpBook, objTmpSheet, objTempRange 	' Object定義:Ref2
  Dim headstring, bodystring 	' work用文字列
  Dim objWkBk, objWkSt			' work用オブジェクト
  
  Set objXls = CreateObject("Excel.Application")
  
  'If Not objXls Then Exit Sub
  If objXls Is Nothing Then Exit Sub	' エクセル起動失敗時にsubroutinからexit
  objXls.Visible = True					' エクセル画面の表示有無(True/False:表示/非表示)
  objXls.ScreenUpdating = True			' エクセル画面の更新表示(True/False:有効/無効)
  
  ' 参照先のブックとシートをopen
  Set objRefBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\Ref2.xlsx")
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
  eadstring = "output"
  'Msgbox "range:" & CStr(objRefRange(objRefRange.Row + objRefRange.Rows.Count -1 , objRefRange.Column + objRefRange.Columns.Count - 1))
  
  ' 行ごとのデータをテンプレートに書き込み別ファイルで出力
  bodystring = ""
  For intR = 2 To objRefRange.Rows.Count
    
    bodystring = bodystring & CStr(objRefRange(objRefRange.Row + intR -1 , objRefRange.Column + 1))
    
    If objRefRange(objRefRange.Row + intR -1 , objRefRange.Column + 0).text _
    		= objRefRange(objRefRange.Row + intR , objRefRange.Column + 0).text Then 
    	bodystring = bodystring & Chr(10)' ※Chr(10)：セル内改行コード。通常のテキストで表示すると改行されない。
    Else
    	Set objWkBk = objXls.Workbooks.Add()	' 出力用workbookを新規作成
    	'Set objWkSt = objWkBk.Sheets("Sheet1")
    	objWkBk.Sheets("Sheet1").Name = ("dummy")
    	'Msgbox "cre_dummy"
    	
    	' テンプレートシートを生成ブックにコピー
    	objTmpSheet.Copy(objWkBk.Sheets("dummy"))	' シートコピー後、dummyシートの前(左側)に配置
    	Set objWkSt = objWkBk.Sheets("temp1")		' コピー後のシートにオブジェクト名付与
    	
    	' dummyシートを削除
    	' XXX 不要シート削除処理は可能な限り避けたい
    	objXls.DisplayAlerts = False
    	objWkBk.Sheets("dummy").Delete
    	objXls.DisplayAlerts = True
    	'Msgbox "tempcopied"
    	
    	' 文字列書き込み
    	objWkSt.Cells(6,3).Value = bodystring
    	objWkSt.Cells(6,3).NumberFormatLocal = "@"
    	
    	' 同名のファイルが既に存在する場合は上書き
    	objXls.DisplayAlerts = False	'一時的にアラートを強制解除
    	objWkBk.SaveAs(GetCurrentDirectory() & "\output" & intR - 1 & ".xlsx")	' Workbookを保存
    	objWkBk.Close					' Workbookを閉じる
    	Set objWkBk = Nothing			' 出力用オブジェクト開放
    	objXls.DisplayAlerts = True		'アラートを規定設定に修正
    	
    	bodystring = ""
    
    End IF
    
  Next
  ' 全ファイルのクローズ/objectの開放("Quit" / "Nothing")
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