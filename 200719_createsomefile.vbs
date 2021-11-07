
' Note & Hist
' v1：「Ref1.xlsx」の表から、情報を取り出す
' v2：「tmp1.xlsx」の形式で特定のセルに行ごとに文字列を別のファイルに書き出す
' v3：文字列生成にはセル内の改行コードを埋め込んで改行表示対応
' v4：ファイルの生成単位をレコードのキー値の条件で判定する(使用ファイル：MappingTable,tmp1)


' Main関数として主処理を定義。最後続で呼び出して使用。エラーハンドリングを簡易化するため。
Sub Main ()
  
  'On Error Resume Next
  Dim objXls									' Object定義:エクセル
  Dim objRefBook, objRefSheet, objRefRange 		' Object定義:MappingTable
  Dim objTmpBook, objTmpSheet, objTempRange 	' Object定義:MappingTable
  Dim clmHeader, wrtUnit
  Dim strFile, strSheet, clmCounter 	' work用文字列
  Dim objWkBk, objWkSt			' work用オブジェクト
  
  Set objXls = CreateObject("Excel.Application")
  
  'If Not objXls Then Exit Sub
  If objXls Is Nothing Then Exit Sub	' エクセル起動失敗時にsubroutinからexit
  objXls.Visible = True					' エクセル画面の表示有無(True/False:表示/非表示)
  objXls.ScreenUpdating = True			' エクセル画面の更新表示(True/False:有効/無効)
  
  ' 書込情報一覧のブックとシートをopen
  Set objRefBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\MappingTable.xlsx")
  Set objRefSheet = objRefBook.Sheets("MappingTable")
  ' テンプレートファイルのブックとシートをopen
  Set objTmpBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\tmp1.xlsx")
  Set objTmpSheet = objTmpBook.Sheets("temp1")
 
  ' データ取得範囲を指定
  Set objRefRange = objRefSheet.Range("B7:F9") 
  clmHeader = 1		' 列ヘッダー数
  wrtUnit = 4		' 書込情報の単位
      
  'Msgbox "range:" & CStr(objRefRange(objRefRange.Row + objRefRange.Rows.Count -1 , objRefRange.Column + objRefRange.Columns.Count - 1))
  
  ' 行ごとのデータをテンプレートに書き込み 行はファイルごとに準備されている
  clmCounter = 0
  strSheet = "Default"
  
  For intR = 1 To objRefRange.Rows.Count
    
    'clmCounter = clmCounter & CStr(objRefRange(objRefRange.Row + intR -1 , objRefRange.Column + 1))
    
    ' 列ヘッダ情報の取得
    strFile  = objRefRange(intR,1)	'出力ファイル名
    
    Msgbox strFile
    ' 同名のファイルが既に存在する場合は上書き
    Set objWkBk = objXls.Workbooks.Add()	' 出力用workbookを新規作成
    objWkBk.Sheets("Sheet1").Name = ("dummy")
    'Msgbox "cre_dummy"
    ' テンプレートシートを生成ブックにコピー
    objTmpBook.Sheets.Copy(objWkBk.Sheets("dummy"))	' コピー先ファイルに
    
    
    ' 文字列書き込み
    For intC = 1 To objRefRange.Columns.Count
    	clmCounter = intC-clmHeader Mod wrtUnit
    	Msgbox clmCounter
    	IF clmCounter = 1 Then 
    	 strSheet = objRefRange(intR,intC)	'出力シート名
    	 Msgbox "(" & intR & "," & intC & "), " & strSheet
    	 
    	 ' Sheetの存在有無を考慮したエラーハンドリング
    	 On Error Resume Next
    	 Set objWkSt = objWkBk.Sheets(strSheet)
    	 On Error GoTo 0
    	 
    	 If Not objWkSt Is Nothing Then '該当するシートが存在するとき
    	   objWkSt.Range("C6").Value = "エクセル"
    	   objWkSt.Range("C6").NumberFormatLocal = "@"
    	   
    	 End If
    	 Set objWkSt = Nothing
    	 
    	End If
    Next
    
     ' dummyシートを削除
    objXls.DisplayAlerts = False
    objWkBk.Sheets("dummy").Delete
    objXls.DisplayAlerts = True
    'Msgbox "tempcopied"
    objXls.DisplayAlerts = False	'一時的にアラートを強制解除
    objWkBk.SaveAs(GetCurrentDirectory() & "\output" & intR - 1 & ".xlsx")	' Workbookを保存
    objWkBk.Close					' Workbookを閉じる
    Set objWkBk = Nothing			' 出力用オブジェクト開放
    objXls.DisplayAlerts = True		'アラートを規定設定に修正
    clmCounter = 0

    
    
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