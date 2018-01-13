<%
'===================================
'类　　名 : Config.asp
'功　　能 : 常用信息和配置参数配置类
'作　　者 : Mr.Lion
'程序版本 : Version 1.0
'完成时间 : 2011/07/02
'说明：
'增加功能 ：
'===================================

Class ClsConfig

	Private TheImgUrl
	Private TheVideoUrl
	Private TheSysVersion
	Private TheSysAuthor
	Private TheSysCopyRight
	Private TheSysName
	Private TheSysCNName
	
	Private TheSqlUsername           'SQL数据库用户名
	Private TheSqlPassword          'SQL数据库用户密码
	Private TheSqlDatabaseName 
	Private TheSqlHostIP                'SQL主机IP地址（本地可用“127.0.0.1”或“(local)”，非本机请用真实IP

'属性，只读，系统上传图片访问根路径	
	Public Property Get ImgUrl()
		ImgUrl = TheImgUrl
	End Property

'属性，只读，系统视频访问根路径	
	Public Property Get VideoUrl()
		VideoUrl = TheVideoUrl
	End Property
	
'属性，只读，系统版本号	
	Public Property Get SysVersion()
		SysVersion = TheSysVersion
	End Property
	
'属性，只读，系统作者	
	Public Property Get SysAuthor()
		SysAuthor = TheSysAuthor
	End Property
	
'属性，只读，系统版权	
	Public Property Get SysCopyRight()
		SysCopyRight = TheSysCopyRight
	End Property
	
'属性，只读，系统名称	
	Public Property Get SysName()
		SysName = TheSysName
	End Property

'属性，只读，系统中文名称	
	Public Property Get SysCNName()
		SysCNName = TheSysCNName
	End Property
	
	Private Sub Class_Initialize()
		TheImgUrl = "../../Pictures"
		TheVideoUrl = ""
		TheSqlUsername = "JLYGroupSiteAdmin"
		TheSqlPassword = "manage@jlygroupwmsdb#03"
		TheSqlDatabaseName = "../sitedata/lioncms.mdb"
		TheSqlHostIP = ""  
		Call SysInfoInit()
	End Sub
	
	Private Sub Class_Terminate()
		
	End Sub
	
	Private Sub SysInfoInit()
		TheSysVersion = "1.0(access)"
		TheSysAuthor = "Mr.Lion"
		TheSysCopyRight = "Mr.Lion"
		TheSysName = "Lion CMS"
		TheSysCNName = "网站内容管理系统"
		
	End Sub
	
	'组装数据库连接字段
	Private Function CreatConnStr(ByVal dbType, ByVal strDB, ByVal strServer, ByVal strUid, ByVal strPwd)
		Dim TempStr
		Select Case dbType
			Case "0","MSSQL"
				'TempStr = "driver={sql server};server="&strServer&";uid="&strUid&";pwd="&strPwd&";database="&strDB
				TempStr = "Provider = Sqloledb; User ID = " & strUid & "; Password = " & strPwd & "; Initial Catalog = " & strDB & "; Data Source = " & strServer & ";"
				'response.Write(tempstr)
				'response.End()
			Case "1","ACCESS"
				Dim tDb : If Instr(strDB,":")>0 Then : tDb = strDB : Else : tDb = Server.MapPath(strDB) : End If
				TempStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&tDb&";Jet OLEDB:Database Password="&strPwd&";"
			Case "3","MYSQL"
				TempStr = "Driver={mySQL};Server="&strServer&";Port=3306;Option=131072;Stmt=; Database="&strDB&";Uid="&strUid&";Pwd="&strPwd&";"
			Case "4","ORACLE"
				TempStr = "Driver={Microsoft ODBC for Oracle};Server="&strServer&";Uid="&strUid&";Pwd="&strPwd&";"
		End Select
		CreatConnStr = TempStr
	End Function
	
	'生成并返回数据库连接字串
	Public Function ConnStr(ByVal DBType,ByVal DBPath)
		ConnStr = CreatConnStr(DBType,DBPath & TheSqlDatabaseName,TheSqlHostIP,TheSqlUsername,TheSqlPassword)
	End Function
	
end Class	
%>