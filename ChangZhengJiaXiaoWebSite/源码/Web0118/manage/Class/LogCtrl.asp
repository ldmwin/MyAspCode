<%
'===================================
'类　　名 : LogCtrl.asp
'功　　能 : 日志操作类
'作　　者 : Mr.Lion
'程序版本 : Version 1.0
'完成时间 : 2011/05/12
'说明：
'增加功能 ：
'===================================

Class LogCtrl

	Private Sub Class_Initialize()

	End Sub
	
	Private Sub Class_Terminate()
		
	End Sub
	
	
	Public Function LogAdd(ByVal LogType,ByVal Operator,ByVal Desc)
		
		dim EventDesc 
		EventDesc ="Insert into Sys_Log(Operate_Type,Operate_Desc,Operate_Time,Operate_IP,Operator_ID) values(" & LogType & ",'" & Desc &"',#" & now() & "#,'" & GetIP() & "'," & Operator & ")"
			
		db.DoExecute(EventDesc)
		
		'LogAdd = true
		
	End Function
	
	Private Function GetIP()  
    
          Dim   strIPAddr   
          If   Request.ServerVariables("HTTP_X_FORWARDED_FOR")   =   ""   OR   InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),   "unknown")   >   0   Then   
                  strIPAddr   =   Request.ServerVariables("REMOTE_ADDR")   
          ElseIf   InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),   ",")   >   0   Then   
                  strIPAddr   =   Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),   1,   InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),   ",")-1)   
          ElseIf   InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),   ";")   >   0   Then   
                  strIPAddr   =   Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),   1,   InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),   ";")-1)   
          Else   
                  strIPAddr   =   Request.ServerVariables("HTTP_X_FORWARDED_FOR")   
          End   If   
          getIP   =   Trim(Mid(strIPAddr,   1,   30))   
	End   Function

end Class
%>