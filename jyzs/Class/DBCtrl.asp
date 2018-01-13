<%
'===================================
'类　　名 : DbCtrl.asp
'功　　能 : 数据库操作类
'作　　者 : Dingjun
'程序版本 : Version 1.0
'完成时间 : 2008.04.15
'增加功能 ：2009.12.09 分页
'===================================

Class DbCtrl

	Private db_connstr
	Private db_sql
	
	Private db_debug
	Private db_conn
	Private db_err
	
	Public pd_rscount		'每页记录数
	Public pd_count			'每次显示页数
	Public pd_url			'URL路径
	Public pd_id
	Public pd_class			'div类型
'	Public rs
	Private curpage
	
	Private db_cmd			'调用存储过程的cmd对象,by mr.lion 2011.5.12
	Private db_rs			'调用存储过程的rs对象,by mr.lion 2011.5.12
'	Private TheDBerr		'数据库出错信息,by mr.lion 2011.5.12
	
	
	Private Sub Class_Initialize()
		db_sql = ""
		db_debug = true					'调试模式是否开启
		db_err = "出现错误："
	End Sub
	
	Private Sub Class_Terminate()
		Set db_conn = Nothing
		If db_debug And db_err <> "出现错误：" Then Response.Write(db_err)
	End Sub

''属性，只读，错误信息	by mr.lion 2011.5.12
'	Public Property Get DBErr()
'		DBErr = TheDBerr
'	End Property


'dbConnStr
'属性，要操作的数据库连接

    Public Property Let dbConnStr(ByVal ConnStr)
        db_connstr = ConnStr
    End Property
    Public Property Get dbConnStr()
        dbConnStr = db_connstr
    End Property

'dbConn
'属性，要操作的数据库连接
	
	Public Property Let dbConn(ConnObj)
		If IsObject(ConnObj) Then
			Set db_conn = ConnObj
		End If
	End Property
	
'dbSql
'属性，要执行的sql语句
	
    Public Property Let dbSql(ByVal SqlStr)
        db_sql = SqlStr
    End Property
    Public Property Get dbSql()
        dbSql = db_sql
    End Property

'dbErr
'属性，只读，错误信息
	
	Public Property Get dbErr()
		dbErr = db_err
	End Property

'dbVersion
'属性，只读，版本信息
	
	Public Property Get dbVersion
		dbVersion = "ASP Database Control Class V1.0 By Dingjun"
	End Property


'db.OpenConn()
'功  能 : 建立数据库连接对象

	Function OpenConn()
		On Error Resume Next
		Dim objConn
		Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open db_connstr
		If Err.number <> 0 Then
			Response.Write("<script>alert('数据库连接出错！出错原因:（" & err.description & "）');window.location.href='../';</Script>")
			objConn.Close
			Set objConn = Nothing
			Response.End
		End If
		Set db_conn = objConn
		'response.Write("aaaaaaaaaaa")
	End Function

'db.CloseConn()
'功  能 : 释放数据库连接对象

	Function CloseConn()
		db_conn.close
		Set db_conn = Nothing
	End Function

'db.AddRecordByRS()
'功  能 : recordset 方式记录添加 by mr.lion 2011.11.29

	Function AddRecordByRS(Action,Table)
		'On Error Resume Next
		Dim RSID : RSID = 0
		
		if Action = "addnew" then
			set rs=server.createObject("Adodb.Recordset")
			sql="select top 1 * from " & Table
			rs.open sql,db_conn,1,3
			rs.addnew
	
			Set db_rs = rs
		elseif Action = "update" then
			db_rs.update
			
			RSID = db_rs(0)
			
			db_rs.close
			set db_rs = nothing	
			
		end if
		
		AddRecordByRS = RSID
		
	End Function
	
'db.RSCmdAddPra()
'功  能 : 添加rs数值 by mr.lion 2011.11.29

	Function RSCmdAddPra(TableItem,ItemValue)
		On Error Resume Next
		db_rs(TableItem) = ItemValue
		
		
	End Function	
	
'db.CreateProCmd()
'功  能 : 建立存储过程访问用cmd by mr.lion 2011.5.12

	Function CreateProCmd(Proname)
		'On Error Resume Next
		Set objCmd = Server.CreateObject("ADODB.Command")
		
		objCmd.ActiveConnection = db_conn
		objCmd.CommandTimeout= 60
		'使用存储过程
		objCmd.CommandType = 4
		objCmd.CommandText = Proname

		Set db_cmd = objCmd
		'response.Write("aaaaaaaaaaa")
	End Function
	
'db.DestroyProCmd()
'功  能 : 建立存储过程访问用cmd by mr.lion 2011.5.12

	Function DestroyProCmd()

		Set db_cmd = nothing
		'response.Write("aaaaaaaaaaa")
	End Function
	
	
'db.ProCmdAddPra()
'功  能 : 建立存储过程cmd参数 by mr.lion 2011.5.12

	Function ProCmdAddPra(ByVal PrmName,ByVal PrmType,ByVal PrmDirection,ByVal PrmSize,ByVal PrmValue)
		On Error Resume Next
		db_cmd.parameters.append db_cmd.createparameter(PrmName,PrmType,PrmDirection,PrmSize,PrmValue)

	End Function
	
	
	
'db.ProCmdExcute()
'功  能 : 存储过程cmd执行查询 by mr.lion 2011.5.12

	Function ProCmdExcute()
		
		Dim IsExeSuccess : IsExeSuccess = false
		
		On Error Resume Next

		
		db_cmd.Prepared = true 
	
		db_cmd.Execute()
		
		If Err.number <> 0 Then

				db_err = "存储过程查询失败，出错原因（" & err.description  &"）" 
				set db_cmd= nothing

		else
			IsExeSuccess = true
		
		End If

		ProCmdExcute = IsExeSuccess		
		
	End Function	

	
	
	
'db.ProCmdGetOutPra()
'功  能 : 存储过程开放，输出参数 by mr.lion 2011.5.12

	Function ProCmdGetOutPra(ByVal PrmName)
		On Error Resume Next		
	
		ProCmdGetOutPra = db_cmd(PrmName)			
		
	End Function

'db.ProCmdGetOutRS()
'功  能 : 存储过程开，输出参数 by mr.lion 2011.5.12

	Function ProCmdGetOutRS()
		On Error Resume Next

		Set rs = Server.CreateObject("ADODB.RecordSet")
		
		rs.cursorlocation = 3
 
		rs.open db_cmd,,3

		set ProCmdGetOutRS = rs		
		
	End Function
	
'db.AddRecordBySql()
'功  能 : 插入记录，并返回记录ID by mr.lion 2011.5.12

	Function AddRecordBySql(ByVal SqlStr)
		On Error Resume Next
		
		Dim RSID : RSID = 0
		
		set rs=db_conn.Execute(SqlStr)     '为ms sql数据准备的插入方式
	
		If Err.number <> 0 Then
			db_err = db_err & "写入数据库出错！<br />"
				If db_debug Then db_err = db_err & "错误信息："& Err.Description
				'DoExecute "ROLLBACK TRAN Tran_Insert"	'如果存在添加事务（事务滚回）
				Exit Function
		else
			RSID = rs(0)
		end if
		
		rs.close
		set rs=nothing	
		
		AddRecordBySql = RSID		
		
	End Function

'db.getRecord(TableName,FieldsList,Condition,OrderField,ShowN)
'功  能	: 取得符合条件的纪录集
'返回值	: Object 纪录集对象
'参  数 :
'TableName	: String  表名称
'FieldsList	: String  字段名称，用逗号隔开，留空则为全部字段
'Condition	: String  查询条件
'OrderField	: String  排序方式
'ShowN		: Integer 获取纪录的数量，相当于sql中的 Select Top N

	Public Function getRecord(ByVal TableName,ByVal FieldsList,ByVal Condition,ByVal OrderField,ByVal ShowN)
		On Error Resume Next
		Dim rs
		db_sql = ""
		db_sql="select "
		If ShowN > 0 Then
			db_sql = db_sql & " top " & ShowN & " "
		End If
		If FieldsList<>"" Then
			db_sql = db_sql & FieldsList
		Else
			db_sql = sdb_sql & " * "
		End If
		db_sql = db_sql & " from [" & TableName & "]"
		If Condition <> "" Then
			db_sql = db_sql & " where " & Condition
		End If
		If OrderField <> "" Then
			db_sql = db_sql & " order by " & OrderField
		End If
		Set rs=Server.CreateObject("adodb.recordset")
			With rs
			.ActiveConnection = db_conn
			.CursorType = 3
			.LockType = 3
			.Source = db_sql
			.Open 
			If Err.number <> 0 Then
				db_err = db_err & "无效的查询条件！<br />"
				If db_debug Then db_err = db_err & "错误信息："& Err.Description
				.Close
				Set rs = Nothing
				Response.End()
				Exit Function
			End If	
		End With
		Set getRecord = rs
	End Function

'db.getRecordBySQL(SqlStr)
'功  能 : 根据sql语句取得纪录集
'返回值 : Object 纪录集对象
'参  数 :
'SqlStr	: String  用于生成记录集的SQL语句

	Public Function getRecordBySQL(ByVal SqlStr)
		'On Error Resume Next
		Dim rs
		db_sql = ""
		db_sql = SqlStr
		Set rs=Server.CreateObject("adodb.recordset")
			With rs
			.ActiveConnection =db_conn
			.CursorType = 3
			.LockType = 3
			.Source = db_sql
			.Open 
			If Err.number <> 0 Then
				db_err = db_err & "无效的查询条件！<br />"
				If db_debug Then db_err = db_err & "错误信息："& Err.Description
				.Close
				Set rs = Nothing
				Response.End()
				Exit Function
			End If	
		End With
		'response.Write(rs.recordcount&"+++++++++++++++++")
		Set getRecordBySQL = rs
	End Function

'分页用
	Public Function getRecordBySQL_PD(ByVal SqlStr)
		On Error Resume Next
		db_sql = ""
		db_sql = SqlStr
		Set rs=Server.CreateObject("adodb.recordset")
			With rs
			.ActiveConnection =db_conn
			.CursorType = 3
			.LockType = 3
			.Source = db_sql
			.Open 
			If Err.number <> 0 Then
				db_err = db_err & "无效的查询条件！<br />"
				If db_debug Then db_err = db_err & "错误信息："& Err.Description
				.Close
				Set rs = Nothing
				Response.End()
				Exit Function
			End If	
		End With

		if pd_rscount = 0 then
			rs.pagesize = rs.recordcount
		else
			rs.pagesize = pd_rscount				'每页记录条数
		end if
		curpage = request.querystring(pd_id) 	'将URL参数传给curpage变量
		if curpage = "" then curpage = 1
		rs.absolutepage = curpage
		
		Set getRecordBySQL_PD = rs
	End Function

'db.AddRecord(TableName, ValueList)
'功  能 : 添加一个新的纪录
'参  数 :
'TableName : String  表名称
'ValueList : Array   插入的字段及值，只能是数组且应遵循前面的参数约定

	Public Function AddRecord(ByVal TableName, ByVal ValueList)
		On Error Resume Next
		Dim tmp_filed, tmp_value
		db_sql = ""
		tmp_filed = ValueToSql(TableName,ValueList,2)
		tmp_value = ValueToSql(TableName,ValueList,3)
		db_sql = "Insert Into [" & TableName & "] (" & tmp_filed & ") Values (" & tmp_value & ")"
		'response.write db_sql & "<hr />"
		DoExecute(db_sql)
		If Err.number <> 0 Then
			db_err = db_err & "写入数据库出错！<br />"
			If db_debug Then db_err = db_err & "错误信息："& Err.Description
			'DoExecute "ROLLBACK TRAN Tran_Insert"	'如果存在添加事务（事务滚回）
			Exit Function
		End If
	End Function
	
'db.UpdateRecord(TableName, Condition, ValueList)
'功  能 : 根据指定条件更新纪录
'参  数 :
'TableName : String  表名称
'Condition : String  更新条件，如果是数组应遵循前面的参数约定
'ValueList : Array   更新的字段及值，只能是数组且应遵循前面的参数约定

	Public Function UpdateRecord(ByVal TableName,ByVal Condition,ByVal ValueList)
		On Error Resume Next
		Dim tmp_str
		db_sql = ""
		tmp_str = ValueToSql(TableName,ValueList,1)
		db_sql = "Update [" & TableName & "] Set " & tmp_str & " Where " & Condition
		DoExecute(db_sql)
		If Err.number <> 0 Then
			db_err = db_err & "更新数据库出错！<br />"
			If db_debug Then db_err = db_err & "错误信息："& Err.Description
			'DoExecute "ROLLBACK TRAN Tran_Update"	'如果存在添加事务（事务滚回）			
			Exit Function
		End If
	End Function
	
'db.DeleteRecord(TableName,IDFieldName,IDValues)
'功  能 : 删除符合条件的纪录
'参  数 :
'TableName	 : String  表名称
'IDFieldName : String  表的Id字段的名称
'IDValues	 : String  删除条件，可以是由逗号隔开的多个Id号

	Public Function DeleteRecord(ByVal TableName,ByVal IDFieldName,ByVal IDValues)
		On Error Resume Next
		db_sql = ""
		db_sql = "Delete From [" & TableName & "] Where [" & IDFieldName & "] In (" &IDValues & ")"
		DoExecute(db_sql)
		If Err.number <> 0 Then
			db_err = db_err & "删除数据出错！<br />"
			If db_debug Then db_err = db_err & "错误信息："& Err.Description
			'DoExecute "ROLLBACK TRAN Tran_Delete"	'如果存在添加事务（事务滚回）
			Exit Function
		End If
	End Function
	
'db.C(ObjRs)
'功  能 : 关闭纪录集对象
'参  数 :
'ObjRs	: Object  纪录集对象

	Public Function C(ByVal ObjRs)
		ObjRs.close()
		Set ObjRs = Nothing
	End Function

'db.ValueToSql(TableName, ValueList, sType)
'功  能 : 格式化sql语句
'参  数 :
'TableName : String  表名称
'ValueList : Array   字段及值
'sType	   : Integer 类型

	Private Function ValueToSql(ByVal TableName, ByVal ValueList, ByVal sType)
		Dim tmp_str
		tmp_str = ValueList
		If IsArray(ValueList) Then
			tmp_str = ""
			Dim rs, tmp_field, tmp_value, i
			Set rs = Server.CreateObject("adodb.recordset")
			With rs
				.ActiveConnection = db_conn
				.CursorType = 3
				.LockType = 3
				.Source ="select * from [" & TableName & "] where 1 = -1"
				.Open
				For i = 0 to Ubound(ValueList)
					tmp_field = Left(ValueList(i),Instr(ValueList(i),":")-1)
					tmp_value = Mid(ValueList(i),Instr(ValueList(i),":")+1)
					If i <> 0 Then
						tmp_str = tmp_str & ","
					End If
					Select Case .Fields(tmp_field).Type
						Case 7,130,133,134,135,8,129,200,201,202,203
							tmp_value = "'"&tmp_value&"'"
						Case 11
							If UCase(cstr(Trim(tmp_value)))="TRUE" Then
								tmp_value = "1"
							Else 
								tmp_value = "0"
							End If
						Case Else
							tmp_value = tmp_value
					End Select
					Select Case sType
						Case 1
							tmp_str = tmp_str & "[" & tmp_field & "] = " & tmp_value
						Case 2
							tmp_str = tmp_str & "[" & tmp_field & "]"
						Case 3
							tmp_str = tmp_str & tmp_value
						Case Else
							tmp_str = tmp_str
					End Select
				Next
			End With
			If Err.number <> 0 Then
				db_err = db_err & "生成SQL语句出错！<br />"
				If db_debug Then db_err = db_err & "错误信息："& Err.Description
				rs.close()
				Set rs = Nothing
				Exit Function
			End If
			rs.Close()
			Set rs = Nothing
		End If
		ValueToSql = tmp_str
	End Function

'db.DoExecute(SqlStr)
'功  能 : 执行指定的sql语句
'参  数 :
'SqlStr	: String  需要执行的sql语句

	Public Function DoExecute(ByVal SqlStr)
		Dim ExecuteCmd
		Set ExecuteCmd = Server.CreateObject("ADODB.Command")
		With ExecuteCmd
			.ActiveConnection = db_conn
			.CommandText = SqlStr
			.Execute
		End With
		Set ExecuteCmd = Nothing
	End Function

'pd.ShowPages()
'功  能	: 显示分页链接
	
	Public Function GetPages(rs)
	
		Dim path : path = pd_url & pd_id & "=" 
		Dim pagestr :  pagestr = "<div id='" & pd_class & "'>"

		if rs.pagecount <> 0 then
		
			if curpage=1 then
				pagestr = pagestr & " &lt;&lt; "
			else
				pagestr = pagestr & "<a href='" & path & 1 & "'> &lt;&lt; </a>"
			end if

			if curpage=1 then
				pagestr = pagestr & " &lt; "
			else
				pagestr = pagestr & "<a href='" & path & curpage-1 & "'> &lt; </a>"
			end if

			Dim plink : plink = int((curpage-1)/pd_count)*pd_count
			for i = plink+1 to plink+pd_count
				if i <= rs.pagecount then
					if i = Cint(curpage) then
						pagestr = pagestr & "<span>" & i & " </span>"	
					else
						pagestr = pagestr & "<a href='" & path & i & "'> " & i & " </a>"
					end if
				end if
			next

			if rs.pagecount<curpage+1 then
				pagestr = pagestr & " &gt; "
			else
				pagestr = pagestr & "<a href='" & path & curpage+1 & "'> &gt; </a>"
			end if

			if rs.pagecount<curpage+1 then
				pagestr = pagestr & " &gt;&gt; "
			else
				pagestr = pagestr & "<a href='" & path & rs.pagecount & "'> &gt;&gt; </a>"
			end if

			pagestr = pagestr & "<select name=" & pd_id & " id=" & pd_id & " onchange=""window.location.href='"&path&"'+this.options[this.options.selectedIndex].text;"">"
			for i = 1 to rs.pagecount
				if i = Cint(curpage) then
					pagestr = pagestr & "<option value=" & i & " selected='selected'>" & i & "</option>"
				else
					pagestr = pagestr & "<option value=" & i & ">" & i & "</option>"
				end if
			next
			pagestr = pagestr & "</select>"
		end if
		
		pagestr = pagestr & "共" & rs.pagecount & "页 " & rs.recordcount & "条记录"
		pagestr = pagestr & "</div>"

		GetPages = pagestr

	End Function
	
End Class
%>