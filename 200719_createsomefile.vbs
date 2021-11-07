
' Note
' �uRef1.xlsx�v�̕\����A�������o��
' �utmp1.xlsx�v�̌`���œ���̃Z���ɍs���Ƃɕ������ʂ̃t�@�C���ɏ����o��


Sub Main ()
  
  'On Error Resume Next
 
  Dim objXls ' Object��`:�G�N�Z��
  Dim objRefBook, objRefSheet, objRefRange 		' Object��`:refference
  Dim objTmpBook, objTmpSheet, objTempRange 	' Object��`:refference
  Dim headstring, bodystring 	' work�p������
  Dim objWkBk, objWkSt			' work�p�I�u�W�F�N�g
  
  Set objXls = CreateObject("Excel.Application")
  'If Not objXls Then Exit Sub
  If objXls Is Nothing Then Exit Sub
  objXls.Visible = True
  objXls.ScreenUpdating = True
  
  ' �Q�Ɛ�̃u�b�N�ƃV�[�g��open
  Set objRefBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\Ref1.xlsx")
  Set objRefSheet = objRefBook.Sheets("Sheet1")
  ' �e���v���[�g��̃u�b�N�ƃV�[�g��open
  Set objTmpBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\tmp1.xlsx")
  Set objTmpSheet = objTmpBook.Sheets("temp1")
 
   ' �f�[�^�擾�͈͂��w��
  'Set objRefRange = objSheet.Range("A1:E2") 
   ' �f�[�^�̂���͈͂��擾
   ' �f�[�^�͈͖͂����w�肵���ق����悢�B�B�B
  Set objRefRange = objRefSheet.UsedRange
  
    ' �s���Ƃ̃f�[�^���擾
  headstring = "output"
  'Msgbox "range:" & CStr(objRefRange(objRefRange.Row + objRefRange.Rows.Count -1 , objRefRange.Column + objRefRange.Columns.Count - 1))
  
  For intR = 2 To objRefRange.Rows.Count
    bodystring = ""
    
    For intC = 1 To objRefRange.Columns.Count
      bodystring = bodystring & CStr(objRefRange(objRefRange.Row + intR -1 , objRefRange.Column + intC - 1))
    Next
    'Msgbox bodystring
    
    ' Workbook��V�K�쐬
    Set objWkBk = objXls.Workbooks.Add()
    'Set objWkSt = objWkBk.Sheets("Sheet1")
    objWkBk.Sheets("Sheet1").Name = ("dummy")
    'Msgbox "cre_dummy"
    
    ' �e���v���[�g�V�[�g���琶���V�[�g�ɓ��e���R�s�[
    objTmpSheet.Copy(objWkBk.Sheets("dummy"))
    
    objXls.DisplayAlerts = False      ' �f�[�^������ƃA���[�g�\�������̂ňꎞ�I�ɏ����B
    objWkBk.Sheets("dummy").Delete
    objXls.DisplayAlerts = True
    'Msgbox "tempcopied"
    ' �����񏑂�����
    objWkBk.Sheets("temp1").Cells(4,3).Value = bodystring
    objWkBk.Sheets("temp1").Cells(4,3).NumberFormatLocal = "@"
    
    ' �����̃t�@�C�������ɑ��݂���ꍇ�͏㏑��
    '�ꎞ�I�ɃA���[�g����������
    objXls.DisplayAlerts = False
    ' Workbook��ۑ�
    objWkBk.SaveAs(GetCurrentDirectory() & "\output" & intR - 1 & ".xlsx")
    ' Workbook�����
    objWkBk.Close
    Set objWkBk = Nothing
    '�A���[�g���K��ݒ�ɏC��
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
Msgbox "����"