call catch()
If Err.Number <> 0 Then
	WScript.Echo Now() & ":" & Err.Number & ":" & Err.Description
	call sendMsg("cpi-it-infra@mail.canon","","�y�G���[�z�C���^�[�l�b�g���������e","���s�t�H���_�� ie_mainte.log ���m�F���Ă�������")
	WScript.Quit(Err.Number)
Else
	WScript.Quit(0)
End If

Sub catch()
	On Error Resume Next
	Call ie_mainte()
End Sub

sub ie_mainte()
	WScript.Echo Now() & ":�����J�n" 

	'�V�F�����N������
	Dim wsh
	Set wsh = WScript.CreateObject("WScript.Shell")

	'On Error Resume Next
	Dim rs
	Dim rs2
	Dim rs3
	Dim rs4
	dim sql
	dim msg

	Dim oParam 
	Set oParam = WScript.Arguments

	' IE�N��
	Set objIE = CreateObject("InternetExplorer.Application")
	objIE.Visible = True

	objIE.Navigate "http://cipapp.cgn.canon.co.jp/cip/app/WXAH/WXAH_05/WXAHA300.asp"
	waitIE(objIE)

	'���O�C��
	objIE.document.all.inuserid.Value = oParam(0)
	objIE.document.all.inpasswd.Value = oParam(1)
	objIE.document.all.incfcompany.Value = "ACH ,�L���m���v���V�W����"
	objIE.document.all.login.Click
	waitIE(objIE)

	'FlowLites�ڑ�
	Dim connect
	Set connect = CreateObject("ADODB.Connection")
	connect.Open "Provider=OraOLEDB.Oracle;" & _
		     "Data Source=CPIORCL;User ID=FL_MNT_USER;Password=FL_MNT_PASS"

	'�؂�ւ��ς݂Ŋ�]������15���ȓ��y�ѐ؂�ւ��҂��Ŋ�]�������ς݂̃f�[�^���擾
        sql= _
         "SELECT DISTINCT                                                                                                                                                                           "&_
         "    A.SINSEI_CODE,                                                                                                                                                                        "&_
         "    A.HAKKOU_DATE,                                                                                                                                                                        "&_
         "    A.HAKKOU_SYAIN_SIMEI,                                                                                                                                                                 "&_
         "    A.AUTO_NO,                                                                                                                                                                            "&_           
         "    A.SYORUI_CODE                                                                                                                                                                         "&_
         "FROM                                                                                                                                                                                      "&_
         "    (SELECT SINSEI_CODE, HAKKOU_DATE, HAKKOU_SYAIN_SIMEI,AUTO_NO,SYORUI_CODE FROM TR_SINSEI_DATA_HEADER@FLOWDB  WHERE SYORUI_CODE ='IT_18' AND JOUTAI_KBN IN ('0','1'))A,                 "&_
         "    (SELECT SINSEI_CODE, KOUMOKU_VALUE                           FROM TR_SINSEI_DATA_DAT@FLOWDB     WHERE KOUMOKU_KEY = 'DAT1')B,                                                         "&_
         "     TR_SINSEI_SYONIN@FLOWDB C                                                                                                                                                            "&_
         "WHERE                                                                                                                                                                                     "&_
         "    A.SINSEI_CODE = B.SINSEI_CODE(+) AND                                                                                                                                                  "&_
         "    A.SINSEI_CODE = C.SINSEI_CODE(+) AND                                                                                                                                                  "&_
         "     NOT EXISTS (SELECT 1 FROM FL_MNT_USER.FL_AUTO_NO_RIREKI WHERE AUTO_NO = A.AUTO_NO) AND                                                                                               "&_
         "    ((SUBSTR(B.KOUMOKU_VALUE,1,8) BETWEEN TO_CHAR(SYSDATE-15,'YYYYMMDD') AND TO_CHAR(SYSDATE,'YYYYMMDD') AND C.SYONIN_JOUTAI_KBN = 1 AND C.SYONIN_NAME='�{�Ԑؑ֎��{' )  OR               "&_
         "     (SUBSTR(B.KOUMOKU_VALUE,1,8) <= TO_CHAR(SYSDATE,'YYYYMMDD') AND C.SYONIN_JOUTAI_KBN = 9 AND C.SYONIN_NAME='�{�Ԑؑ֎��{'))                                                           "


	Set rs = connect.Execute(sql)
	dim chkFlg
	dim nonFlg
	dim retn
	dim exstFlg
	Do Until rs.Eof = True
		'�܂��͑��݃`�F�b�N
		sql= _ 
                    "SELECT                                                                                                                                    "&_
                    " LPAD(B.KOUMOKU_VALUE,6,'0'),                                                                                                             "&_
                    " A.KOUMOKU_VALUE                                                                                                                          "&_
                    "FROM                                                                                                                                      "&_
                    " (SELECT * FROM FLOW_ACH.TR_SINSEI_HYO_SEL@FLOWDB WHERE SINSEI_CODE='" & rs(0) & "' )A,                                                   "&_
                    " (SELECT * FROM FLOW_ACH.TR_SINSEI_HYO_CHR@FLOWDB WHERE SINSEI_CODE='" & rs(0) & "'  AND KOUMOKU_KEY = 'CHR1_T1')B                        "&_
                    "WHERE                                                                                                                                     "&_
                    " A.SINSEI_CODE = B.SINSEI_CODE AND                                                                                                        "&_
                    " A.HYO_GYONO = B.HYO_GYONO     AND                                                                                                        "&_
                    " KOUMOKU_NAME IS NOT NULL                                                                                                                 "

		Set rs2 = connect.Execute(sql)
		exstFlg=1
		Do Until rs2.Eof = True
			if InternetExistChk(objIE, rs2(0))=0 then
				exstFlg=0
			end if
			rs2.MoveNext
		Loop
		'�R�����g���݃`�F�b�N
		sql="SELECT COUNT(*) FROM TR_SINSEI_SYONIN WHERE SYONIN_COMMENT IS NOT NULL AND SYONIN_NAME='�{�Ԑؑ֎��{' AND SINSEI_CODE='" & rs(0) & "'"
                     
		Set rs4 = connect.Execute(sql)
		WScript.Echo(rs(2) & ":" & exstFlg & ":" & CInt(rs4(0)))

		objIE.Navigate "http://cipapp.cgn.canon.co.jp/cip/app/WXAH/WXAH_05/WXAHA300.asp"
		waitIE(objIE)
		if exstFlg=0 and CInt(rs4(0))=0 then
		    '�R�����g������Ζ���Ƃő��݂��Ȃ��Ј�������Ƃ������ƂȂ̂ŋp���ʒm���[�����M�B�����ɂȂ��Đl���}�X�^�Ɍ����Ј�������̂Ŏ��O�ɋp���͖����B
			'�C���t�����[�����O���X�g�Ƀ��[�����M

				msg=               "��������������������������������������Ꮘ���˗���⁡����������������������������������"
				msg=msg & vbcrlf & "���݂��Ȃ��Ј�ID���܂܂�Ă��܂�"
				msg=msg & vbcrlf & "���[�N�t���[�V�X�e���Ŋm�F�Ə������肢���܂��B"
				msg=msg & vbcrlf & "----------------------------------------------------------------------------------------"
				msg=msg & vbcrlf & "�y ���@�� �z�@�C���^�[�l�b�g���p�\����"     
				msg=msg & vbcrlf & "�y�\���ԍ��z�@" & rs(3)
				msg=msg & vbcrlf & "�y ���s�� �z�@" & rs(2)
				msg=msg & vbcrlf & "�y ���s�� �z�@" & mid(rs(1),1,4) & "/" & mid(rs(1),5,2) & "/" & mid(rs(1),7,2)
				msg=msg & vbcrlf & "----------------------------------------------------------------------------------------"
				msg=msg & vbcrlf & "����������������������������������������������������������������������������������������"
				call sendMsg("cpi-it-infra@mail.canon","","�����˗�(���߂�){���ށF�C���^�[�l�b�g���p�\����}",msg)
			
			
		else
		'�����e���{
 	   	        sql= _ 
                            "SELECT                                                                                                                                    "&_
                            " LPAD(B.KOUMOKU_VALUE,6,'0'),                                                                                                             "&_
                            " A.KOUMOKU_VALUE                                                                                                                          "&_
                            "FROM                                                                                                                                      "&_
                            " (SELECT * FROM FLOW_ACH.TR_SINSEI_HYO_SEL@FLOWDB WHERE SINSEI_CODE='" & rs(0) & "' )A,                            "&_
                            " (SELECT * FROM FLOW_ACH.TR_SINSEI_HYO_CHR@FLOWDB WHERE SINSEI_CODE='" & rs(0) & "'  AND KOUMOKU_KEY = 'CHR1_T1')B "&_
                            "WHERE                                                                                                                                     "&_
                            " A.SINSEI_CODE = B.SINSEI_CODE AND                                                                                                        "&_
                            " A.HYO_GYONO = B.HYO_GYONO     AND                                                                                                        "&_
                            " KOUMOKU_NAME IS NOT NULL                                                                                                                 "

			Set rs2 = connect.Execute(sql)
			chkFlg=0
			nonFlg=0
			Do Until rs2.Eof = True
				chkFlg=1
				if rs2(1)="01" then
					'�J�n�i�����e���͍̂�ƍρE���F�ς݂Ɋւ�炸�������Ɏ��{�j
					retn=InternetUpd(objIE, rs2(0),"1")
					WScript.Echo(rs2(0) & ":�J�n:" & retn)
					if retn<>0 then
					    nonFlg=1
					end if
				else
					'��~�i�����e���͍̂�ƍρE���F�ς݂Ɋւ�炸�������Ɏ��{�j
					retn=InternetUpd(objIE, rs2(0),"0")
					WScript.Echo(rs2(0) & ":�I��:" & retn)
					if retn<>0 then
					    nonFlg=1
					end if
				end if
				rs2.MoveNext
			Loop
			if chkFlg=0 then
				call sendMsg("cpi-it-infra@mail.canon","","�y�x���z�C���^�[�l�b�gID",rs(2) & "�̈˗����Â����[�łȂ����m�F���Ă��������B")
			end if


				'�C���t�����[�����O���X�g�ɏ��F�ʒm���[�����M

					msg=               "��������������������������������������Ꮘ���˗���⁡����������������������������������"
					msg=msg & vbcrlf & "���ނ̊m�F�˗������Ă��܂��B"
					msg=msg & vbcrlf & "���[�N�t���[�V�X�e���Ŋm�F�Ə������肢���܂��B"
					msg=msg & vbcrlf & "----------------------------------------------------------------------------------------"
					msg=msg & vbcrlf & "�y ���@�� �z�@�C���^�[�l�b�g���p�\����" 
					msg=msg & vbcrlf & "�y�\���ԍ��z�@" & rs(3) 
					msg=msg & vbcrlf & "�y ���s�� �z�@" & rs(2)
					msg=msg & vbcrlf & "�y ���s�� �z�@" & mid(rs(1),1,4) & "/" & mid(rs(1),5,2) & "/" & mid(rs(1),7,2)
					msg=msg & vbcrlf & "----------------------------------------------------------------------------------------"
					msg=msg & vbcrlf & "����������������������������������������������������������������������������������������"
					call sendMsg("cpi-it-infra@mail.canon","","�����˗�{���ށF�C���^�[�l�b�g���p�\����}",msg)
				
				
			
		end if
		    '�����e�[�u���ɏ���o�^'
			    sql="INSERT INTO FL_MNT_USER.FL_AUTO_NO_RIREKI VALUES ('" & rs(4) & "','" & rs(3) & "',SYSDATE)"
			    WScript.Echo(sql)
			    connect.Execute(sql)
		rs.MoveNext

		rs2.Close
		Set rs2 = Nothing
	Loop
	rs.Close
	Set rs = Nothing

	'�X�V

	'���O�A�E�g
	objIE.Navigate "http://cipcom.cgn.canon.co.jp/cip/zz/asp/WXZZA304.asp"
	waitIE(objIE)

	' �����j��
	objIE.Quit
	Set objIE = Nothing
	WScript.Echo Now() & ":�����I��" 
end sub



Function InternetExistChk(objIE, uid) 
'���^�[���@����:1�A���[�U�[����:0

'�ΏێҌ���
	objIE.document.all.CIP_CD_PERSON.Value = uid
	objIE.document.all.submit1.Click
	waitIE(objIE)

	flg=0
	For Each objITEM In objIE.Document.getElementsByTagName("INPUT")
		if trim(objITEM.Name)="PROXY_ROUTE_0" then
			flg=1
		End If
	Next
	objIE.document.all.btnReturn.Click
	waitIE(objIE)

	'���^�[��
	InternetExistChk=flg

end Function

Function InternetUpd(objIE, uid,kengen) 
'���^�[���@����:0�A���[�U�[����:1
'�����@kengen�c1:�A0:�s��

'�ΏێҌ���
	objIE.document.all.CIP_CD_PERSON.Value = uid
	objIE.document.all.submit1.Click
	waitIE(objIE)

	'���W�I�{�b�N�X�Q��
'	For Each objITEM In objIE.document.all.PROXY_ROUTE_0
'		Wscript.Echo(Now() & ":" & objITEM.Value & ":" & objITEM.Checked)
'	Next
	
	flg=0
	For Each objITEM In objIE.Document.getElementsByTagName("INPUT")
		if trim(objITEM.Name)="PROXY_ROUTE_0" then
			flg=1
		End If
	Next

	if flg=1 then
		'���W�I�{�b�N�X�X�V
		For Each objITEM In objIE.document.all.PROXY_ROUTE_0  
			flg=0
			if objITEM.Value=kengen then
				objITEM.Checked=true
			end if
		Next
		if flg=0 then
			'�T�u�~�b�g
			objIE.document.forms(0).Submit
			waitIE(objIE)
		end if
		'�߂�
		objIE.document.all.SubBtn.Click
		waitIE(objIE)
	else
		flg=1
		objIE.document.all.btnReturn.Click
		waitIE(objIE)
	end if

	objIE.Navigate "http://cipapp.cgn.canon.co.jp/cip/app/WXAH/WXAH_05/WXAHA300.asp"
	waitIE(objIE)

	'���^�[��
	InternetUpd=flg

end Function

' IE���r�W�[��Ԃ̊ԑ҂��܂�
Sub waitIE(objIE)
    
    Do While objIE.Busy = True Or objIE.readystate <> 4
        WScript.Sleep 100
    Loop
    WScript.Sleep 500
    
End Sub

sub sendMsg(mailTo,mailCc,mailTitle,mailText)
	Set oMsg = CreateObject("CDO.Message")
	oMsg.From = "ACH-CSYDB@prec.canon.co.jp"
	oMsg.To = mailTo
	oMsg.Cc = mailCc
	oMsg.Subject = mailTitle
	oMsg.TextBody = mailText
	oMsg.Configuration.Fields.Item _
	  ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	oMsg.Configuration.Fields.Item _
	  ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
	    "nonauth-smtp.global.canon.co.jp"
	oMsg.Configuration.Fields.Item _
	  ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	oMsg.Configuration.Fields.Update
	oMsg.Send
end sub

