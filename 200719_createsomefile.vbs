
' Note & Hist
' v1�F�uRef1.xlsx�v�̕\����A�������o��
' v2�F�utmp1.xlsx�v�̌`���œ���̃Z���ɍs���Ƃɕ������ʂ̃t�@�C���ɏ����o��
' v3�F�����񐶐��ɂ̓Z�����̉��s�R�[�h�𖄂ߍ���ŉ��s�\���Ή�
' v4�F�t�@�C���̐����P�ʂ����R�[�h�̃L�[�l�̏����Ŕ��肷��(�g�p�t�@�C���FRef2,tmp1)
'
' Task(���t�@�N�^�����O)
' �E�Z���̒l��萔(constant�l)�ŊǗ�
' �E�J�����P�ʂŏ��擾�@�g�p�����ԍ���萔�ŌŒ�


Sub Main ()
  
  'On Error Resume Next
 
  Dim objXls									' Object��`:�G�N�Z��
  Dim objRefBook, objRefSheet, objRefRange 		' Object��`:Ref2
  Dim objTmpBook, objTmpSheet, objTempRange 	' Object��`:Ref2
  Dim headstring, bodystring 	' work�p������
  Dim objWkBk, objWkSt			' work�p�I�u�W�F�N�g
  
  Set objXls = CreateObject("Excel.Application")
  
  'If Not objXls Then Exit Sub
  If objXls Is Nothing Then Exit Sub	' �G�N�Z���N�����s����subroutin����exit
  objXls.Visible = True					' �G�N�Z����ʂ̕\���L��(True/False:�\��/��\��)
  objXls.ScreenUpdating = True			' �G�N�Z����ʂ̍X�V�\��(True/False:�L��/����)
  
  ' �Q�Ɛ�̃u�b�N�ƃV�[�g��open
  Set objRefBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\Ref2.xlsx")
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
  eadstring = "output"
  'Msgbox "range:" & CStr(objRefRange(objRefRange.Row + objRefRange.Rows.Count -1 , objRefRange.Column + objRefRange.Columns.Count - 1))
  
  ' �s���Ƃ̃f�[�^���e���v���[�g�ɏ������ݕʃt�@�C���ŏo��
  bodystring = ""
  For intR = 2 To objRefRange.Rows.Count
    
    bodystring = bodystring & CStr(objRefRange(objRefRange.Row + intR -1 , objRefRange.Column + 1))
    
    If objRefRange(objRefRange.Row + intR -1 , objRefRange.Column + 0).text _
    		= objRefRange(objRefRange.Row + intR , objRefRange.Column + 0).text Then 
    	bodystring = bodystring & Chr(10)' ��Chr(10)�F�Z�������s�R�[�h�B�ʏ�̃e�L�X�g�ŕ\������Ɖ��s����Ȃ��B
    Else
    	Set objWkBk = objXls.Workbooks.Add()	' �o�͗pworkbook��V�K�쐬
    	'Set objWkSt = objWkBk.Sheets("Sheet1")
    	objWkBk.Sheets("Sheet1").Name = ("dummy")
    	'Msgbox "cre_dummy"
    	
    	' �e���v���[�g�V�[�g�𐶐��u�b�N�ɃR�s�[
    	objTmpSheet.Copy(objWkBk.Sheets("dummy"))	' �V�[�g�R�s�[��Adummy�V�[�g�̑O(����)�ɔz�u
    	Set objWkSt = objWkBk.Sheets("temp1")		' �R�s�[��̃V�[�g�ɃI�u�W�F�N�g���t�^
    	
    	' dummy�V�[�g���폜
    	' XXX �s�v�V�[�g�폜�����͉\�Ȍ����������
    	objXls.DisplayAlerts = False
    	objWkBk.Sheets("dummy").Delete
    	objXls.DisplayAlerts = True
    	'Msgbox "tempcopied"
    	
    	' �����񏑂�����
    	objWkSt.Cells(6,3).Value = bodystring
    	objWkSt.Cells(6,3).NumberFormatLocal = "@"
    	
    	' �����̃t�@�C�������ɑ��݂���ꍇ�͏㏑��
    	objXls.DisplayAlerts = False	'�ꎞ�I�ɃA���[�g����������
    	objWkBk.SaveAs(GetCurrentDirectory() & "\output" & intR - 1 & ".xlsx")	' Workbook��ۑ�
    	objWkBk.Close					' Workbook�����
    	Set objWkBk = Nothing			' �o�͗p�I�u�W�F�N�g�J��
    	objXls.DisplayAlerts = True		'�A���[�g���K��ݒ�ɏC��
    	
    	bodystring = ""
    
    End IF
    
  Next
  ' �S�t�@�C���̃N���[�Y/object�̊J��("Quit" / "Nothing")
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