
' Note & Hist
' v1�F�uRef1.xlsx�v�̕\����A�������o��
' v2�F�utmp1.xlsx�v�̌`���œ���̃Z���ɍs���Ƃɕ������ʂ̃t�@�C���ɏ����o��
' v3�F�����񐶐��ɂ̓Z�����̉��s�R�[�h�𖄂ߍ���ŉ��s�\���Ή�
' v4�F�t�@�C���̐����P�ʂ����R�[�h�̃L�[�l�̏����Ŕ��肷��(�g�p�t�@�C���FMappingTable,tmp1)
' v5�Fv4�̉ғ���
' v6�Fv5�ɃR�����g�t�^�BMappingTable�̗�Ǎ����f������t�^�B

' Main�֐��Ƃ��Ď又�����`�B�Ō㑱�ŌĂяo���Ďg�p�B�G���[�n���h�����O���ȈՉ����邽�߁B
Sub Main ()
  
  'On Error Resume Next
  Dim objXls									' Object��`:�G�N�Z��
  Dim objRefBook, objRefSheet, objRefRange 		' Object��`:MappingTable
  Dim objTmpBook, objTempRange 					' Object��`:MappingTable
  Dim clmHeader, wrtUnit, clmCounter			' MappingTable�̃f�[�^�\�����
  Dim strDataArea								' MappingTable�̃f�[�^�͈�
  Dim strTmpName								' �e���v���[�g�t�@�C����������
  Dim strFile, strShtName, strOutCel, strtOutString, strOutFmt	' �o�͗p������Q
  Dim objWkBk, objWkSt			' work�p�I�u�W�F�N�g
  
  Set objXls = CreateObject("Excel.Application")
  
  'If Not objXls Then Exit Sub
  If objXls Is Nothing Then Exit Sub	' �G�N�Z���N�����s����subroutin����exit
  objXls.Visible = True					' �G�N�Z����ʂ̕\���L��(True/False:�\��/��\��)
  objXls.ScreenUpdating = True			' �G�N�Z����ʂ̍X�V�\��(True/False:�L��/����)
  
  ' �������ꗗ�̃u�b�N�ƃV�[�g��open
  Set objRefBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\MappingTable.xlsx")
  Set objRefSheet = objRefBook.Sheets("MappingTable")
  ' �e���v���[�g�t�@�C�����̓Ǎ�
  strTmpName = objRefSheet.Range("C2").Text
  strDataArea = objRefSheet.Range("C3").Text
  MsgBox "�e���v���[�g��" & strTmpName & ", �͈́�" & strDataArea
  
  ' �f�[�^�擾�͈͂��w��
  Set objRefRange = objRefSheet.Range(strDataArea) 
  clmHeader = 1		' ��w�b�_�[��
  wrtUnit = 4		' ����������̏��̒P��
  

  ' �e���v���[�g�t�@�C���̃u�b�N�ƃV�[�g��open
  Set objTmpBook = objXls.Workbooks.Open(GetCurrentDirectory() & "\" & strTmpName & ".xlsx")
  
  ' �s���Ƃ̃f�[�^���e���v���[�g�ɏ������� �s�̓t�@�C���P�ʂŏ�������Ă���
  For intR = 1 To objRefRange.Rows.Count 'RowsRoutine-------------------------------------------------
    clmCounter = 0
    
    strFile  = objRefRange(intR,1)	'�o�̓t�@�C����
    
    Msgbox strFile
    
    ' �o�͗pworkbook��V�K�쐬
    Set objWkBk = objXls.Workbooks.Add()			' �����̃t�@�C�������ɑ��݂���ꍇ�͏㏑��
    objWkBk.Sheets("Sheet1").Name = ("xxxDUMMYxxx")	' �f�t�H���g�̃V�[�g�����_�~�[���ɕϊ�
    
    ' �e���v���[�g�t�@�C���̑S�V�[�g��V�K�����u�b�N�ɃR�s�[
    objTmpBook.Sheets.Copy(objWkBk.Sheets("xxxDUMMYxxx"))
     ' xxxDUMMYxxx�V�[�g���폜
    objXls.DisplayAlerts = False			' �ꎞ�I�ɃA���[�g����������
    objWkBk.Sheets("xxxDUMMYxxx").Delete	' �_�~�[�V�[�g��F�e���v���[�g�t�@�C���Ɠ���V�[�g�\����
    objXls.DisplayAlerts = True				' �A���[�g���K��ݒ�ɏC��
    ' office2010�ł̓f�t�H���g�̃V�[�g����3�ƂȂ�̂ŁA�_�~�[������
    
    ' �����񏑂�����
    For intC = 1 To objRefRange.Columns.Count 'ColsRoutine---------------------------------------------
    	clmCounter = (intC-clmHeader) Mod wrtUnit		' �I�t�Z�b�g���������f�[�^����\���P�ʂŏ��Z
    	
    	IF clmCounter = 1 and objRefRange(intR,intC)<>"***" Then 			' �f�[�^�P�ʂŏ����ݏ��������{
    	 strShtName = objRefRange(intR,intC)		' �o�̓V�[�g��
    	 strOutCel = objRefRange(intR,intC+1)		' �o�̓Z���ʒu
    	 strtOutString = objRefRange(intR,intC+2)	' �o�͕�����
    	 strOutFmt = objRefRange(intR,intC+3)		' �o�̓t�H�[�}�b�g
    	 ' Msgbox "(" & intR & "," & intC & "), SheetName:" & strShtName
    	 
    	 ' Sheet�̑��ݗL�����l�������G���[�n���h�����O
    	 On Error Resume Next
    	 Set objWkSt = objWkBk.Sheets(strShtName)
    	 On Error GoTo 0
    	 
    	 If Not objWkSt Is Nothing Then '�Y������V�[�g�����݂���Ƃ�
    	   objWkSt.Range(strOutCel).Value = strtOutString
    	   objWkSt.Range(strOutCel).NumberFormatLocal = strOutFmt
    	   
    	 End If
    	 Set objWkSt = Nothing
    	 
    	End If
    Next 'ColsRoutine-------------------------------------------------------------------------------------------
    
    objXls.DisplayAlerts = False	'�ꎞ�I�ɃA���[�g����������
    objWkBk.SaveAs(GetCurrentDirectory() & "\" & strFile & ".xlsx")	' Workbook��ۑ�
    objWkBk.Close					' Workbook�����
    Set objWkBk = Nothing			' �o�͗p�I�u�W�F�N�g�J��
    objXls.DisplayAlerts = True		'�A���[�g���K��ݒ�ɏC��
    
  Next 'RowsRoutine----------------------------------------------------------------------------------------------
  
  
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