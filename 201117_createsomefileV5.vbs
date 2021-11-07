
' Note & Hist
' v1�F�uRef1.xlsx�v�̕\����A�������o��
' v2�F�utmp1.xlsx�v�̌`���œ���̃Z���ɍs���Ƃɕ������ʂ̃t�@�C���ɏ����o��
' v3�F�����񐶐��ɂ̓Z�����̉��s�R�[�h�𖄂ߍ���ŉ��s�\���Ή�
' v4�F�t�@�C���̐����P�ʂ����R�[�h�̃L�[�l�̏����Ŕ��肷��(�g�p�t�@�C���FMappingTable,tmp1)


' Main�֐��Ƃ��Ď又�����`�B�Ō㑱�ŌĂяo���Ďg�p�B�G���[�n���h�����O���ȈՉ����邽�߁B
Sub Main ()
  
  'On Error Resume Next
  Dim objXls									' Object��`:�G�N�Z��
  Dim objRefBook, objRefSheet, objRefRange 		' Object��`:MappingTable
  Dim objTmpBook, objTmpSheet, objTempRange 	' Object��`:MappingTable
  Dim clmHeader, wrtUnit
  Dim strFile, strSheet, clmCounter 	' work�p������
  Dim objWkBk, objWkSt			' work�p�I�u�W�F�N�g
  
  Set objXls = CreateObject("Excel.Application")
  
  'If Not objXls Then Exit Sub
  If objXls Is Nothing Then Exit Sub	' �G�N�Z���N�����s����subroutin����exit
  objXls.Visible = True					' �G�N�Z����ʂ̕\���L��(True/False:�\��/��\��)
  objXls.ScreenUpdating = True			' �G�N�Z����ʂ̍X�V�\��(True/False:�L��/����)
  
  ' �������ꗗ�̃u�b�N�ƃV�[�g��open
  Set objRefBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\MappingTable.xlsx")
  Set objRefSheet = objRefBook.Sheets("MappingTable")
  ' �e���v���[�g�t�@�C���̃u�b�N�ƃV�[�g��open
  Set objTmpBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\tmp1.xlsx")
  Set objTmpSheet = objTmpBook.Sheets("temp1")
 
  ' �f�[�^�擾�͈͂��w��
  Set objRefRange = objRefSheet.Range("B7:F9") 
  clmHeader = 1		' ��w�b�_�[��
  wrtUnit = 4		' �������̒P��
      
  'Msgbox "range:" & CStr(objRefRange(objRefRange.Row + objRefRange.Rows.Count -1 , objRefRange.Column + objRefRange.Columns.Count - 1))
  
  ' �s���Ƃ̃f�[�^���e���v���[�g�ɏ������� �s�̓t�@�C�����Ƃɏ�������Ă���
  clmCounter = 0
  strSheet = "Default"
  
  For intR = 1 To objRefRange.Rows.Count
    
    'clmCounter = clmCounter & CStr(objRefRange(objRefRange.Row + intR -1 , objRefRange.Column + 1))
    
    ' ��w�b�_���̎擾
    strFile  = objRefRange(intR,1)	'�o�̓t�@�C����
    
    Msgbox strFile
    ' �����̃t�@�C�������ɑ��݂���ꍇ�͏㏑��
    Set objWkBk = objXls.Workbooks.Add()	' �o�͗pworkbook��V�K�쐬
    objWkBk.Sheets("Sheet1").Name = ("dummy")
    'Msgbox "cre_dummy"
    ' �e���v���[�g�V�[�g�𐶐��u�b�N�ɃR�s�[
    objTmpBook.Sheets.Copy(objWkBk.Sheets("dummy"))	' �R�s�[��t�@�C����
    
    
    ' �����񏑂�����
    For intC = 1 To objRefRange.Columns.Count
    	clmCounter = intC-clmHeader Mod wrtUnit
    	Msgbox clmCounter
    	IF clmCounter = 1 Then 
    	 strSheet = objRefRange(intR,intC)	'�o�̓V�[�g��
    	 Msgbox "(" & intR & "," & intC & "), " & strSheet
    	 
    	 ' Sheet�̑��ݗL�����l�������G���[�n���h�����O
    	 On Error Resume Next
    	 Set objWkSt = objWkBk.Sheets(strSheet)
    	 On Error GoTo 0
    	 
    	 If Not objWkSt Is Nothing Then '�Y������V�[�g�����݂���Ƃ�
    	   objWkSt.Range("C6").Value = "�G�N�Z��"
    	   objWkSt.Range("C6").NumberFormatLocal = "@"
    	   
    	 End If
    	 Set objWkSt = Nothing
    	 
    	End If
    Next
    
     ' dummy�V�[�g���폜
    objXls.DisplayAlerts = False
    objWkBk.Sheets("dummy").Delete
    objXls.DisplayAlerts = True
    'Msgbox "tempcopied"
    objXls.DisplayAlerts = False	'�ꎞ�I�ɃA���[�g����������
    objWkBk.SaveAs(GetCurrentDirectory() & "\output" & intR - 1 & ".xlsx")	' Workbook��ۑ�
    objWkBk.Close					' Workbook�����
    Set objWkBk = Nothing			' �o�͗p�I�u�W�F�N�g�J��
    objXls.DisplayAlerts = True		'�A���[�g���K��ݒ�ɏC��
    clmCounter = 0

    
    
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