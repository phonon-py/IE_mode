call catch()
If Err.Number <> 0 Then
	WScript.Echo Now() & ":" & Err.Number & ":" & Err.Description
	call sendMsg("cpi-it-infra@mail.canon","","【エラー】インターネット自動メンテ","実行フォルダの ie_mainte.log を確認してください")
	WScript.Quit(Err.Number)
Else
	WScript.Quit(0)
End If

Sub catch()
	On Error Resume Next
	Call ie_mainte()
End Sub

sub ie_mainte()
	WScript.Echo Now() & ":処理開始" 

	'シェルを起動する
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

	' IE起動
	Set objIE = CreateObject("InternetExplorer.Application")
	objIE.Visible = True

	objIE.Navigate "http://cipapp.cgn.canon.co.jp/cip/app/WXAH/WXAH_05/WXAHA300.asp"
	waitIE(objIE)

	'ログイン
	objIE.document.all.inuserid.Value = oParam(0)
	objIE.document.all.inpasswd.Value = oParam(1)
	objIE.document.all.incfcompany.Value = "ACH ,キヤノンプレシジョン"
	objIE.document.all.login.Click
	waitIE(objIE)

	'FlowLites接続
	Dim connect
	Set connect = CreateObject("ADODB.Connection")
	connect.Open "Provider=OraOLEDB.Oracle;" & _
		     "Data Source=CPIORCL;User ID=FL_MNT_USER;Password=FL_MNT_PASS"

	'切り替え済みで希望日から15日以内及び切り替え待ちで希望日到来済みのデータを取得
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
         "    ((SUBSTR(B.KOUMOKU_VALUE,1,8) BETWEEN TO_CHAR(SYSDATE-15,'YYYYMMDD') AND TO_CHAR(SYSDATE,'YYYYMMDD') AND C.SYONIN_JOUTAI_KBN = 1 AND C.SYONIN_NAME='本番切替実施' )  OR               "&_
         "     (SUBSTR(B.KOUMOKU_VALUE,1,8) <= TO_CHAR(SYSDATE,'YYYYMMDD') AND C.SYONIN_JOUTAI_KBN = 9 AND C.SYONIN_NAME='本番切替実施'))                                                           "


	Set rs = connect.Execute(sql)
	dim chkFlg
	dim nonFlg
	dim retn
	dim exstFlg
	Do Until rs.Eof = True
		'まずは存在チェック
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
		'コメント存在チェック
		sql="SELECT COUNT(*) FROM TR_SINSEI_SYONIN WHERE SYONIN_COMMENT IS NOT NULL AND SYONIN_NAME='本番切替実施' AND SINSEI_CODE='" & rs(0) & "'"
                     
		Set rs4 = connect.Execute(sql)
		WScript.Echo(rs(2) & ":" & exstFlg & ":" & CInt(rs4(0)))

		objIE.Navigate "http://cipapp.cgn.canon.co.jp/cip/app/WXAH/WXAH_05/WXAHA300.asp"
		waitIE(objIE)
		if exstFlg=0 and CInt(rs4(0))=0 then
		    'コメント無ければ未作業で存在しない社員がいるということなので却下通知メール送信。当日になって人事マスタに現れる社員もいるので事前に却下は無理。
			'インフラメーリングリストにメール送信

				msg=               "■■■■■■■■■■■■■■■■■■≪≪処理依頼≫≫■■■■■■■■■■■■■■■■■■"
				msg=msg & vbcrlf & "存在しない社員IDが含まれています"
				msg=msg & vbcrlf & "ワークフローシステムで確認と処理お願いします。"
				msg=msg & vbcrlf & "----------------------------------------------------------------------------------------"
				msg=msg & vbcrlf & "【 書　類 】　インターネット利用申請書"     
				msg=msg & vbcrlf & "【申請番号】　" & rs(3)
				msg=msg & vbcrlf & "【 発行者 】　" & rs(2)
				msg=msg & vbcrlf & "【 発行日 】　" & mid(rs(1),1,4) & "/" & mid(rs(1),5,2) & "/" & mid(rs(1),7,2)
				msg=msg & vbcrlf & "----------------------------------------------------------------------------------------"
				msg=msg & vbcrlf & "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■"
				call sendMsg("cpi-it-infra@mail.canon","","処理依頼(差戻し){書類：インターネット利用申請書}",msg)
			
			
		else
		'メンテ実施
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
					'開始（メンテ自体は作業済・承認済みに関わらず無条件に実施）
					retn=InternetUpd(objIE, rs2(0),"1")
					WScript.Echo(rs2(0) & ":開始:" & retn)
					if retn<>0 then
					    nonFlg=1
					end if
				else
					'停止（メンテ自体は作業済・承認済みに関わらず無条件に実施）
					retn=InternetUpd(objIE, rs2(0),"0")
					WScript.Echo(rs2(0) & ":終了:" & retn)
					if retn<>0 then
					    nonFlg=1
					end if
				end if
				rs2.MoveNext
			Loop
			if chkFlg=0 then
				call sendMsg("cpi-it-infra@mail.canon","","【警告】インターネットID",rs(2) & "の依頼が古い帳票でないか確認してください。")
			end if


				'インフラメーリングリストに承認通知メール送信

					msg=               "■■■■■■■■■■■■■■■■■■≪≪処理依頼≫≫■■■■■■■■■■■■■■■■■■"
					msg=msg & vbcrlf & "書類の確認依頼が来ています。"
					msg=msg & vbcrlf & "ワークフローシステムで確認と処理お願いします。"
					msg=msg & vbcrlf & "----------------------------------------------------------------------------------------"
					msg=msg & vbcrlf & "【 書　類 】　インターネット利用申請書" 
					msg=msg & vbcrlf & "【申請番号】　" & rs(3) 
					msg=msg & vbcrlf & "【 発行者 】　" & rs(2)
					msg=msg & vbcrlf & "【 発行日 】　" & mid(rs(1),1,4) & "/" & mid(rs(1),5,2) & "/" & mid(rs(1),7,2)
					msg=msg & vbcrlf & "----------------------------------------------------------------------------------------"
					msg=msg & vbcrlf & "■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■"
					call sendMsg("cpi-it-infra@mail.canon","","処理依頼{書類：インターネット利用申請書}",msg)
				
				
			
		end if
		    '履歴テーブルに情報を登録'
			    sql="INSERT INTO FL_MNT_USER.FL_AUTO_NO_RIREKI VALUES ('" & rs(4) & "','" & rs(3) & "',SYSDATE)"
			    WScript.Echo(sql)
			    connect.Execute(sql)
		rs.MoveNext

		rs2.Close
		Set rs2 = Nothing
	Loop
	rs.Close
	Set rs = Nothing

	'更新

	'ログアウト
	objIE.Navigate "http://cipcom.cgn.canon.co.jp/cip/zz/asp/WXZZA304.asp"
	waitIE(objIE)

	' 制御を破棄
	objIE.Quit
	Set objIE = Nothing
	WScript.Echo Now() & ":処理終了" 
end sub



Function InternetExistChk(objIE, uid) 
'リターン　存在:1、ユーザー無し:0

'対象者検索
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

	'リターン
	InternetExistChk=flg

end Function

Function InternetUpd(objIE, uid,kengen) 
'リターン　正常:0、ユーザー無し:1
'引数　kengen…1:可、0:不可

'対象者検索
	objIE.document.all.CIP_CD_PERSON.Value = uid
	objIE.document.all.submit1.Click
	waitIE(objIE)

	'ラジオボックス参照
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
		'ラジオボックス更新
		For Each objITEM In objIE.document.all.PROXY_ROUTE_0  
			flg=0
			if objITEM.Value=kengen then
				objITEM.Checked=true
			end if
		Next
		if flg=0 then
			'サブミット
			objIE.document.forms(0).Submit
			waitIE(objIE)
		end if
		'戻る
		objIE.document.all.SubBtn.Click
		waitIE(objIE)
	else
		flg=1
		objIE.document.all.btnReturn.Click
		waitIE(objIE)
	end if

	objIE.Navigate "http://cipapp.cgn.canon.co.jp/cip/app/WXAH/WXAH_05/WXAHA300.asp"
	waitIE(objIE)

	'リターン
	InternetUpd=flg

end Function

' IEがビジー状態の間待ちます
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

