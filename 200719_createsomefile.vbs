
' Note
' �uRef1.xlsx�v�̕\����A�������o���A
' �s���Ƃɕ������ʂ̃t�@�C���ɏ����o��

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
  
  ' �Q�Ɛ�̃u�b�N�ƃV�[�g��open
  Set objRefBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\Ref1.xlsx")
  Set objRefSheet = objRefBook.Sheets("Sheet1")
  ' �e���v���[�g��̃u�b�N�ƃV�[�g��open
  'Set objTempBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\tmp1.xlsx")
  'Set objTempSheet = objTempBook.Sheets("Sheet1")
 
   ' �f�[�^�擾�͈͂��w��
  'Set objRefRange = objSheet.Range("A1:E2") 
   ' �f�[�^�̂���͈͂��擾
   ' �f�[�^�͈͖͂����w�肵���ق����悢�B�B�B
  Set objRefRange = objRefSheet.UsedRange
  
    ' �s���Ƃ̃f�[�^���擾
  headstring = "output"
  
  For intR = 2 To objRefRange.Rows.Count
    bodystring = ""
    
    For intC = 1 To objRefRange.Columns.Count
      bodystring = bodystring & CStr(objRefRange(objRefRange.Row + intR -1 , objRefRange.Column + intC - 1))
    Next
    
    ' Workbook��V�K�쐬
    Set objWkBk = objXls.Workbooks.Add()
    Set objWkSt = objWkBk.Sheets("Sheet1")
    
     ' �����񏑂�����
    objWkSt.Cells(1,2).Value = bodystring
    objWkSt.Cells(1,2).NumberFormatLocal = "@"
    
    ' �����̃t�@�C�������ɑ��݂���ꍇ�͏㏑��
    '�ꎞ�I�ɃA���[�g����������
    objXls.DisplayAlerts = False
    ' Workbook��ۑ�
    objWkSt.SaveAs(GetCurrentDirectory() & "\output" & intR - 1 & ".xlsx")
    ' Workbook�����
    objWkBk.Close
    '�A���[�g���K��ݒ�ɏC��
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
Msgbox "����"