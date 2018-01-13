<%
'===================================
'类　　名 : InputCheck.asp
'功　　能 : 输入信息校验和过滤类
'作　　者 : Mr.Lion
'程序版本 : Version 1.0
'完成时间 : 2011/05/10
'说明：此类必须与DBCtrl类和md5函数共同使用
'增加功能 ：
'===================================

Class InputCheck


	
	Private Sub Class_Initialize()
		TheUserLoginStatus = "unlogin"
		TheLoginRemainTime = 1000
	End Sub
	
	Private Sub Class_Terminate()
		
	End Sub
	
	
	Public Function CodeCheck(Code)
	
		dim IsCheckOK:IsCheckOK=true
		
		if Code = "" or Code <> Session("GetCode") 	then 
			IsCheckOK=false
		end if
		
		CodeCheck = IsCheckOK
		
	End Function

end Class
%>