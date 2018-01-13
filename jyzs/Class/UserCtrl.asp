<%
'===================================
'类　　名 : User.asp
'功　　能 : 用户管理类（验证，分权，基础信息查询）
'作　　者 : Mr.Lion
'程序版本 : Version 1.1
'完成时间 : 2011/05/07
'修改时间 : 2011/09/20
'说明：此类必须与DBCtrl类和md5函数共同使用
'增加功能 ：
'v1.1 加入两个只读类属性StatusShow,UserStatus，修改了函数UserInit()及其存储过程使之能够载入新增属性，加入通过id和密码校验用户的函数CheckUserByID()，加入密码修改函数PasswordChange()，加入单项用户基本信息查询函数UserBaseProfile()
'===================================

Class UserCtrl

	Private TheUserID
	Private TheUserName
	Private TheRealName
	Private TheLoginStatus
	Private TheUseerr
	Private TheLoginRemainTime
	Private TheStatusShow
	Private TheStatus

'属性，只读，用户ID	
	Public Property Get UserID()
		UserID = TheUserID
	End Property
	
'属性，只读，用户名	
	Public Property Get UserName()
		UserName = TheUserName
	End Property
	
'属性，只读，真实姓名	
	Public Property Get RealName()
		RealName = TheRealName
	End Property
	
'属性，只读，用户登录状态，unlogin,login	
	Public Property Get LoginStatus()
		LoginStatus = TheLoginStatus
	End Property
	
'属性，只读，错误信息	
	Public Property Get UserErr()
		UserErr = TheUseerr
	End Property
	
'属性，只读，用户状态（字典查询后）	
	Public Property Get StatusShow()
		StatusShow = TheStatusShow
	End Property
	
'属性，只读，用户状态（字典查询后）	
	Public Property Get UserStatus()
		UserStatus = TheStatus
	End Property

'属性，可写，登录状态保持时间	
    Public Property Let LoginRemainTime(TimeSet)
        TheLoginRemainTime = TimeSet
    End Property
    Public Property Get LoginRemainTime()
        TheLoginRemainTime = LoginRemainTime
    End Property
	
	Private Sub Class_Initialize()
		TheLoginStatus = "unlogin"
		TheLoginRemainTime = 1000
		TheUserID = 0
	End Sub
	
	Private Sub Class_Terminate()
		
	End Sub
	
	
	Public Function CheckUser(UserName,Password)
		
		Dim IsCheckOK : IsCheckOK = false
		Dim DBUserName
		Dim DBUserPW
		Dim DBUserStatus
		'Dim DBUserID
		
		Set rs_user = db.getRecordBySQL("select * from Sys_User where UserName='" & UserName & "'")
		if rs_user.eof or rs_user.bof then
			TheUseerr = "用户名不存在"
		else
			DBUserName = rs_user("username")
			DBUserPW = rs_user("password")
			DBUserStatus = rs_user("Status")
			'DBUserID = rs_user("id")
			if DBUserStatus = "1" then
				
				if DBUserPW = md5(Password) then
				
					'IsCheckOK = DBUserID
					IsCheckOK = true
				
				else
					TheUseerr = "密码错误"
				end if
			else
				TheUseerr = "用户已被锁定"
			end if
		end if
		db.C(rs_user)
		
		CheckUser = IsCheckOK
		
	End Function
	
'通过用户ID和密码查询用户状态是否正常	
	Public Function CheckUserByID(UserID,Password)
		
		Dim IsCheckOK : IsCheckOK = false
		'Dim DBUserName
		Dim DBUserPW
		Dim DBUserStatus
		Dim DBUserID
		
		Set rs_user = db.getRecordBySQL("select * from Sys_User where ID=" & UserID)
		if rs_user.eof or rs_user.bof then
			TheUseerr = "用户名ID不存在"
		else
			'DBUserName = rs_user("username")
			DBUserPW = rs_user("password")
			
			'response.Write(DBUserPW)
			'response.Write("&&" &Password&"&&")
			
		'	response.End()
			DBUserStatus = rs_user("Status")
			DBUserID = rs_user("id")
			if DBUserStatus = "1" then
				
				if DBUserPW = md5(Password) then
				
					'IsCheckOK = DBUserID
					IsCheckOK = true
				
				else
					TheUseerr = "密码错误"
				end if
			else
				TheUseerr = "用户已被锁定"
			end if
		end if
		db.C(rs_user)
		
		CheckUserByID = IsCheckOK
		
	End Function

'通过ID查询用户单一基础资料，以后补写一个数组子类，创建子类后将信息一次性加载在子类中，减少数据库查询操作次数。	
	Public Function UserBaseProfile(UserID,Profile)
		
		Dim QueryResult,SqlStr
		
		select case Profile
			case "Org"
				'SqlStr = "select * from Sys_User where ID='" & UserID & "'"
			case "CreateTime"
				SqlStr = "select Add_Time from Sys_User where ID=" & UserID
			case "LoginNum"
				SqlStr = "select Login_Num from Sys_User where ID=" & UserID
			case "LastLoginTime"
				SqlStr = "select Last_Login_Time from Sys_User where ID=" & UserID
			case "LastLoginIP"
				SqlStr = "select Last_Login_IP from Sys_User where ID=" & UserID
			case else			
			
		end select
		
		if SqlStr<>"" then
		
			Set rs_user = db.getRecordBySQL(SqlStr)
			if rs_user.eof or rs_user.bof then
				QueryResult = "数据查询失败"
			else
				QueryResult = rs_user(0)
			end if
			db.C(rs_user)
		
		end if
		
		UserBaseProfile = QueryResult
		
	End Function
	
	Public Function PassWordChange(UserID,Password)
		
		Dim IsWorkOK : IsCheckOK = false
		Dim Sqlstr
		
		Sqlstr = "update Sys_User set [PassWord] ='" & md5(Password) & "',Last_PWChange_Time=#" & now() & "#,Last_PWChange_IP='" & GetIP() & "' where status=1 and id=" & UserID
		
		db.DoExecute(Sqlstr)
					
		IsWorkOK = true
		
		PassWordChange = IsWorkOK
		
	End Function
	
	Public Function UserLogin(UserName)
		
		Dim IsLoginOK : IsLoginOK = false
		Dim DBUserName
		Dim DBUserPW
		Dim DBUserStatus
		Dim DBUserID
		Dim DBRealName
		
		Set rs_user = db.getRecordBySQL("select * from Sys_User where UserName='" & UserName & "'")
		if rs_user.eof or rs_user.bof then
			TheUseerr = "用户名不存在"
		else
			DBUserName = rs_user("username")
			DBUserPW = rs_user("password")
			DBUserStatus = rs_user("Status")
			DBUserID = rs_user("id")
			DBRealName  = rs_user("RealName")
			if DBUserStatus = "1" then
				
					'IsCheckOK = DBUserID
					
					Session.timeout = TheLoginRemainTime
					Session("UserName") = DBUserName
					TheUserName = DBUserName
					Session("UserID") = DBUserID
					TheUserID = DBUserID
					TheRealName = DBRealName
					'Session("LoginStatus") = "login"		
					TheLoginStatus = "login"
					
					db.DoExecute("update Sys_User set Login_Num=Login_Num +1,Last_Login_Time=#" & now() & "#,Last_Login_IP='" & GetIP() & "' where id=" & DBUserID)	
					'response.Write("update Sys_User set Login_Num=Login_Num +1,Last_Login_Time=getdate(),Last_Login_IP='" & GetIP() & "' where id=" & DBUserID)		
					
					IsLoginOK = true
				
			else
				TheUseerr = "用户已被锁定"
			end if
		end if
		db.C(rs_user)
		
		
		UserLogin = IsLoginOK
	End Function
	
'页面用户登录状态检查，检查session状态确定用户登录	
	
	Public Function LoginCheck()
		
		Dim IsLoginOK : IsLoginOK = false
		
		if Session("UserID") = null or Session("UserID") = "" or Session("UserID") = 0 then
			
			TheUseerr = "用户登录状态失效或未登录"
			
		else
			
			TheUserID = Session("UserID")
			TheUserName = Session("UserName")
			TheLoginStatus = "login"
			
			IsLoginOK = true
			
		end if
		
		LoginCheck = IsLoginOK
		
	End Function

'用户状态校验及数据（包括角色，权限等）装载
	
	Public Function UserInit()
	
		Dim InitUserID : InitUserID = TheUserID
		Dim IsInitSucess : IsInitSucess = false
		Dim QueryStatus

		'调用存储过程，装载用户数据
		
		Set rs_user = db.getRecordBySQL("select RealName,status from Sys_User where ID =" & TheUserID)
		if rs_user.eof or rs_user.bof then
			TheUseerr = "用户数据装载失败，用户信息不存在"
		else
			TheRealName = rs_user("realname")				
			TheStatusShow = StatusResult(6,rs_user("Status"))
			TheStatus = rs_user("Status")
	
			if TheStatus = 1 then
				TheLoginStatus = "inited"			
				IsInitSucess = true
			else
				TheUseerr = "用户数据装载失败，用户状态不正常1"
			end if
		end if
		db.C(rs_user)		
		
		UserInit = 	IsInitSucess
		
	End Function
	
	'自带ip获取函数，以减少该类对其他文件的依赖
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
	
	Public Function StatusResult(StatusType,StatusValue)
	
		Dim StrStatusShow : StrStatusShow = "无状态"
		
		select case StatusType
			case 6
				select case StatusValue
					case 0 
						StrStatusShow = "初始"
					case 1
						StrStatusShow = "正常"
					case 2
						StrStatusShow = "锁定"
					case 3
						StrStatusShow = "测试"
					case 4
						StrStatusShow = "注销"
					case else
						StrStatusShow = "无状态"
				end select
		end select 
	
		
		StatusResult = StrStatusShow
		
	End Function
	

end Class
%>