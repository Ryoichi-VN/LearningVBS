
' Note & Hist
' v1：「Ref1.xlsx」の表から、情報を取り出す
' v2：「tmp1.xlsx」の形式で特定のセルに行ごとに文字列を別のファイルに書き出す
' v3：文字列生成にはセル内の改行コードを埋め込んで改行表示対応
' v4：ファイルの生成単位をレコードのキー値の条件で判定する(使用ファイル：MappingTable,tmp1)
' v5：v4の稼働版
' v6：v5にコメント付与。MappingTableの列読込中断条件を付与。

' Main関数として主処理を定義。最後続で呼び出して使用。エラーハンドリングを簡易化するため。
Sub Main ()
  
  'On Error Resume Next
  Dim objXls									' Object定義:エクセル
  Dim objRefBook, objRefSheet, objRefRange 		' Object定義:MappingTable
  Dim objTmpBook, objTempRange 					' Object定義:MappingTable
  Dim clmHeader, wrtUnit, clmCounter			' MappingTableのデータ構造情報
  Dim strDataArea								' MappingTableのデータ範囲
  Dim strTmpName								' テンプレートファイル名文字列
  Dim strFile, strShtName, strOutCel, strtOutString, strOutFmt	' 出力用文字列群
  Dim objWkBk, objWkSt			' work用オブジェクト
  
  Set objXls = CreateObject("Excel.Application")
  
  'If Not objXls Then Exit Sub
  If objXls Is Nothing Then Exit Sub	' エクセル起動失敗時にsubroutinからexit
  objXls.Visible = True					' エクセル画面の表示有無(True/False:表示/非表示)
  objXls.ScreenUpdating = True			' エクセル画面の更新表示(True/False:有効/無効)
  
  ' 書込情報一覧のブックとシートをopen
  Set objRefBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\MappingTable.xlsx")
  Set objRefSheet = objRefBook.Sheets("MappingTable")
  ' テンプレートファイル名の読込
  strTmpName = objRefSheet.Range("C2").Text
  strDataArea = objRefSheet.Range("C3").Text
  MsgBox "テンプレート＞" & strTmpName & ", 範囲＞" & strDataArea
  
  ' データ取得範囲を指定
  Set objRefRange = objRefSheet.Range(strDataArea) 
  clmHeader = 1		' 列ヘッダー数
  wrtUnit = 4		' 書込文字列の情報の単位
  

  ' テンプレートファイルのブックとシートをopen
  Set objTmpBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\" & strTmpName & ".xlsx")
  
  ' 行ごとのデータをテンプレートに書き込み 行はファイル単位で準備されている
  For intR = 1 To objRefRange.Rows.Count 'RowsRoutine-------------------------------------------------
    clmCounter = 0
    
    strFile  = objRefRange(intR,1)	'出力ファイル名
    
    Msgbox strFile
    
    ' 出力用workbookを新規作成
    Set objWkBk = objXls.Workbooks.Add()			' 同名のファイルが既に存在する場合は上書き
    objWkBk.Sheets("Sheet1").Name = ("xxxDUMMYxxx")	' デフォルトのシート名をダミー名に変換
    
    ' テンプレートファイルの全シートを新規生成ブックにコピー
    objTmpBook.Sheets.Copy(objWkBk.Sheets("xxxDUMMYxxx"))
     ' xxxDUMMYxxxシートを削除
    objXls.DisplayAlerts = False			' 一時的にアラートを強制解除
    objWkBk.Sheets("xxxDUMMYxxx").Delete	' ダミーシート削：テンプレートファイルと同一シート構成に
    objXls.DisplayAlerts = True				' アラートを規定設定に修正
    ' office2010ではデフォルトのシート数が3となるので、ダミー数注意
    
    ' 文字列書き込み
    For intC = 1 To objRefRange.Columns.Count 'ColsRoutine---------------------------------------------
    	clmCounter = (intC-clmHeader) Mod wrtUnit		' オフセットを除いたデータ列を構成単位で除算
    	
    	IF clmCounter = 1 and objRefRange(intR,intC)<>"***" Then 			' データ単位で書込み処理を実施
    	 strShtName = objRefRange(intR,intC)		' 出力シート名
    	 strOutCel = objRefRange(intR,intC+1)		' 出力セル位置
    	 strtOutString = objRefRange(intR,intC+2)	' 出力文字列
    	 strOutFmt = objRefRange(intR,intC+3)		' 出力フォーマット
    	 ' Msgbox "(" & intR & "," & intC & "), SheetName:" & strShtName
    	 
    	 ' Sheetの存在有無を考慮したエラーハンドリング
    	 On Error Resume Next
    	 Set objWkSt = objWkBk.Sheets(strShtName)
    	 On Error GoTo 0
    	 
    	 If Not objWkSt Is Nothing Then '該当するシートが存在するとき
    	   objWkSt.Range(strOutCel).Value = strtOutString
    	   objWkSt.Range(strOutCel).NumberFormatLocal = strOutFmt
    	   
    	 End If
    	 Set objWkSt = Nothing
    	 
    	End If
    Next 'ColsRoutine-------------------------------------------------------------------------------------------
    
    objXls.DisplayAlerts = False	'一時的にアラートを強制解除
    objWkBk.SaveAs(GetCurrentDirectory() & "\" & strFile & ".xlsx")	' Workbookを保存
    objWkBk.Close					' Workbookを閉じる
    Set objWkBk = Nothing			' 出力用オブジェクト開放
    objXls.DisplayAlerts = True		'アラートを規定設定に修正
    
  Next 'RowsRoutine----------------------------------------------------------------------------------------------
  
  
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