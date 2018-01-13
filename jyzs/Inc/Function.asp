<%
'===================================
'文 件 名 : /inc/Function.asp
'功　　能 : 通用函数集
'作　　者 : Mr.Lion
'建立时间 : 2011/05/12
'===================================


'Fun No.01 : 获取ip 

Function   getIP()  
    
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


'选择框比较选定函数
function object_selected(current_value,value)
                      if current_value=value then
                          sel_str="selected"
                      else
                          sel_str=""
                      end if
                      object_selected=sel_str
end function

function IsSelected(current_value,value,obj_type)
		dim selstr
		if obj_type="radio" then
			selstr = "checked"
		else
			selstr = "selected"
		end if
		
		if current_value=value then
		  sel_str= selstr
		else
		  sel_str=""
		end if
		
		IsSelected=sel_str
end function

'显示分页导航条
Function ShowPage(total_rs,total_page,current_page) 
	pagebar = "共 " & total_rs & " 条记录 共 " & total_page & " 页 当前第 " & current_page & " 页"
  if total_page>1 then   
      if current_Page = 1 then 
          pagebar = pagebar & "  |  首页"
          pagebar = pagebar & "  |  上页"
          pagebar = pagebar & "  |  <a href='javascript:viewPage(" & cstr(current_Page+1) &")' language='javascript'>下页</a>"
          pagebar = pagebar & "  |  <a href='javascript:viewPage(" & cstr(total_page) & ")' language='javascript'>尾页</a>"
      elseif current_Page = total_page then                 
          pagebar = pagebar & "  |  <a href='javascript:viewPage(" & cstr(1) & ")' language='javascript'>首页</a>" 
          pagebar = pagebar & "  |  <a href='javascript:viewPage(" & cstr(current_Page-1) & ")' language ='javascript'>上页</a>"
          pagebar = pagebar & "  |  下页"
          pagebar = pagebar & "  |  尾页"
      else 
          pagebar = pagebar & "  |  <a href='javascript:viewPage(" & cstr(1) & ")' language ='javascript'>首页</a>"
          pagebar = pagebar & "  |  <a href='javascript:viewPage(" & cstr(current_Page-1) & ")' language ='javascript'>上页</a>"
          pagebar = pagebar & "  |  <a href='javascript:viewPage(" & cstr(current_Page+1) &")' language ='javascript'>下页</a>"
          pagebar = pagebar & "  |  <a href='javascript:viewPage(" & cstr(total_page) & ")' language ='javascript'>尾页</a>"
      end if 

  else            
      pagebar = pagebar & "  |  首页  |  上页  |  下页  |  尾页"
  end if  

		pagebar = pagebar & "转到第<input type=""text"" name=""goto_page"" value=""" & current_page &""" size=3 maxlength=3>页 <input type=""button"" value=""转到"" name=""cmd_goto"" onClick=""javascript:viewPage(document.all.goto_page.value);"">" 
		
  ShowPage = pagebar
  
end Function

'用户登录后校验及信息装载，装载成功返回true
Function IsUserInit()
	Dim IsInited:IsInited = false
	
	if not (typeName(User)="Empty" or typeName(User)="Nothing") then
		if User.LoginCheck() then	
			if User.UserInit() then
				IsInited = true
			end if	
		end if
	end if  
	
	IsUserInit = IsInited 
End Function

'组装数据库连接字段
'Function CreatConnStr(ByVal dbType, ByVal strDB, ByVal strServer, ByVal strUid, ByVal strPwd)
'	Dim TempStr
'	Select Case dbType
'		Case "0","MSSQL"
'			'TempStr = "driver={sql server};server="&strServer&";uid="&strUid&";pwd="&strPwd&";database="&strDB
'			TempStr = "Provider = Sqloledb; User ID = " & strUid & "; Password = " & strPwd & "; Initial Catalog = " & strDB & "; Data Source = " & strServer & ";"
'			'response.Write(tempstr)
'			'response.End()
'		Case "1","ACCESS"
'			Dim tDb : If Instr(strDB,":")>0 Then : tDb = strDB : Else : tDb = Server.MapPath(strDB) : End If
'			TempStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&tDb&";Jet OLEDB:Database Password="&strPwd&";"
'		Case "3","MYSQL"
'			TempStr = "Driver={mySQL};Server="&strServer&";Port=3306;Option=131072;Stmt=; Database="&strDB&";Uid="&strUid&";Pwd="&strPwd&";"
'		Case "4","ORACLE"
'			TempStr = "Driver={Microsoft ODBC for Oracle};Server="&strServer&";Uid="&strUid&";Pwd="&strPwd&";"
'	End Select
'	CreatConn = TempStr
'End Function

Function ParaEncode(reString) 

	Dim Str:Str=reString
	
	If Not IsNull(Trim(Str)) Then
	
	Str = Replace(Str, "&", "&amp;")	
	Str = Replace(Str, ">", "&gt;")	
	Str = Replace(Str, "<", "&lt;")	
	Str = Replace(Str, CHR(34),"&quot;")	
	Str = Replace(Str, CHR(39),"&#39;")	
	Str = Replace(Str, CHR(13), "")	
	'Str = Replace(Str, " ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)	
	'Str = Replace(Str, " ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)	
	'Str = Replace(Str, " ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)	
	'Str = Replace(Str, " ", "&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)	
	'Str = Replace(Str, " ", "&nbsp;&nbsp;&nbsp;", 1, -1, 1)	
	'Str = Replace(Str, " ", "&nbsp;&nbsp;", 1, -1, 1)	
	Str = Replace(Str, CHR(10), "<br>")	
	Str = Replace(Str, "sel&#101;ct", "select")	
	Str = Replace(Str, "jo&#105;n", "join")	
	Str = Replace(Str, "un&#105;on", "union")	
	Str = Replace(Str, "wh&#101;re", "where")	
	Str = Replace(Str, "ins&#101;rt", "insert")	
	Str = Replace(Str, "del&#101;te", "delete")	
	Str = Replace(Str, "up&#100;ate", "update")	
	Str = Replace(Str, "lik&#101;", "like")	
	Str = Replace(Str, "dro&#112;", "drop")	
	Str = Replace(Str, "cr&#101;ate", "create")	
	Str = Replace(Str, "mod&#105;fy", "modify")	
	Str = Replace(Str, "ren&#097;me", "rename")	
	Str = Replace(Str, "alt&#101;r", "alter")	
	Str = Replace(Str, "ca&#115;t", "cast")
	
	ParaEncode=Str
	
	end if 

End Function

'正则表达验证函数
Function InfoRegularCheck(infostr,regular)'infostr为被验证字符串，regular为配套正则表达式
	Dim CheckResult:CheckResult = false 
	Dim regEx,Match 
	Set regEx = New RegExp 
	regEx.Pattern = regular 
	regEx.IgnoreCase = True
	 
	Set Match = regEx.Execute(infostr) 
	
	if match.count then 
		CheckResult = true
	end if 
	
	InfoRegularCheck = CheckResult
End Function

Function StatusResult(StatusType,StatusValue)
	
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
			case 3
				select case StatusValue
					case 0 
						StrStatusShow = "下线"
					case 1
						StrStatusShow = "上线"
					case 4
						StrStatusShow = "作废"
					case else
						StrStatusShow = "无状态"
				end select
			case 1
				select case StatusValue
					case 0 
						StrStatusShow = "初始"
					case 1
						StrStatusShow = "上线"
					case 2
						StrStatusShow = "下线"
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

Function TopFirstNav(site_id,pid,menustr,path)
	Set rs_menutree = db.getRecordBySQL("select m.*,(select count(*) from Site_Columns s where s.Parent_id=m.id and s.status<>4) as childnodenum,(select Target from Sys_Target where m.Target = Sys_Target.id) as Link_Target,(select top 1 id from Site_Columns where parent_id = m.id and status=1 order by show_order desc,id desc) as childid from Site_Columns m where m.status<>4 and m.parent_id=0 and m.site_id=" & site_id & " order by m.show_order desc,m.id desc")
	
	'response.Write(rs_menutree.recordcount)
	
	i  = 0
	do while not rs_menutree.eof
	
		dim target : target = rs_menutree("Link_Target") 
		
		'response.Write(target)
						
		'if i <> 0 then
			'menustr = menustr & " | "					
		'end if
		
		'response.Write(str)
'					
		if rs_menutree("column_type") = 1 then
			menustr = menustr & "<div class=""header_mainnav_text""><a href=""" & rs_menutree("Link") & """ target=""" & target & """>"						
									
		elseif rs_menutree("column_type") = 2 then					
			
			Set rs_link = db.getRecordBySQL("select top 1 * from Site_Column_Mould where status=1 and id=" & rs_menutree("mould_page") & " and (site_id=0 or site_id=" & site_id & ")order by show_order desc,id desc") 
			
			'response.Write("select top 1 * from Site_Column_Mould where status=1 and id=" & rs_menutree("mould_page") & " and (site_id=0 or site_id=" & site_id & ")order by show_order desc,id desc")
			
			if not(rs_link.eof and rs_link.bof) then
				
				if rs_menutree("UseChildContent") = 0 then 
					menustr = menustr & "<div class=""header_mainnav_text""><a href=""" & path & rs_link("Mould_Path") & rs_link("Mould_File") & "?id=" & rs_menutree("id") &  """ target=""" & target & """>"
				else
					menustr = menustr & "<div class=""header_mainnav_text""><a href=""" & path & rs_link("Mould_Path") & rs_link("Mould_File") & "?id=" & rs_menutree("childid") &  """ target=""" & target & """>"
				
				end if
			else
				menustr = menustr & "<div class=""header_mainnav_text""><a href=""#"" target=""" & target & """>"
			end if
			
			db.C(rs_link)
		else
			menustr = menustr & "<div class=""header_mainnav_text""><a href=""#"" target=""" & target & """>"
		end if 
		 
		menustr = menustr & rs_menutree("column_name") & "</a></div>"

		'Call TopNav(site_id,rs_menutree("ID"),str) '递归调用
							
		rs_menutree.movenext()
		i = i + 1
		
	Loop
	
	db.C(rs_menutree)
	
	TopFirstNav = menustr
			
	'response.Write("yes")
		
End Function

	Function LeftSecNav(site_id,Parent_ID,menustr,path)
		  Set rs_menutree = db.getRecordBySQL("select m.*,(select count(*) from Site_Columns s where s.Parent_id=m.id and s.status<>4) as childnodenum,(select Target from Sys_Target where m.Target = Sys_Target.id) as Link_Target,(select top 1 id from Site_Columns where parent_id = m.id and status=1 order by show_order desc,id desc) as childid from Site_Columns m where m.status<>4 and m.parent_id=" & Parent_ID & " and m.site_id=" & site_id & " order by m.show_order desc,m.id desc")

		  do while not rs_menutree.eof
		  
			  dim target : target = rs_menutree("Link_Target") 

			  menustr = menustr & "<div class=""main_left_text"">"					

			  
			  'response.Write(str)
	  '					
			  if rs_menutree("column_type") = 1 then
				  menustr = menustr & "<a href=""" & path & rs_menutree("Link") & """ target=""" & target & """>"						
										  
			  elseif rs_menutree("column_type") = 2 then					
				  
				  Set rs_link = db.getRecordBySQL("select top 1 * from Site_Column_Mould where status=1 and id=" & rs_menutree("mould_page") & " and (site_id=0 or site_id=" & site_id & ")order by show_order desc,id desc") 
				  
				  if not(rs_link.eof and rs_link.bof) then
					  
					  if rs_menutree("UseChildContent") = 0 then 
						  menustr = menustr & "<a href=""" & path & rs_link("Mould_Path") & rs_link("Mould_File") & "?id=" & rs_menutree("id") &  """ target=""" & target & """>"
					  else
						  menustr = menustr & "<a href=""" & path & rs_link("Mould_Path") & rs_link("Mould_File") & "?id=" & rs_menutree("childid") &  """ target=""" & target & """>"
					  
					  end if
				  else
					  menustr = menustr & "<a href=""#"" target=""" & target & """>"
				  end if
				  
				  db.C(rs_link)
			  else
				  menustr = menustr & "<a href=""#"" target=""" & target & """>"
			  end if 
			   
			  menustr = menustr & rs_menutree("column_name") & "</a></div>"
								  
			  rs_menutree.movenext()
			  
		  Loop
		  
		  db.C(rs_menutree)
		  
		  LeftSecNav = menustr
	End Function
%>