<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>
<!--#include file="../Class/DBCtrl.asp" -->
<!--#include file="../Inc/Config.asp" -->
<!--#include file="../Inc/Function.asp" -->
<%
Dim Config: Set Config = New ClsConfig

Dim db : Set db = New DbCtrl
db.dbConnStr = Config.ConnStr(1,"../")
db.OpenConn()

Dim path : path ="../"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta name="author" content="#" /> 
	<meta name="robots" content="all" />
	<meta name="keywords" content="" />
	<meta name="description" content="" />
	<title>久友木门</title>
	<link rel="shortcut icon" href="../images/favicon.ico" type="image/x-icon" />
	<link href="../style/basic.css" rel="stylesheet" type="text/css" />
	<link href="../style/main.css" rel="stylesheet" type="text/css" />
	<link href="../style/product.css" rel="stylesheet" type="text/css" />
	<script language=javascript>
    function secBoard(n)
    {
    for(i=0;i<menu.childNodes.length;i++)
    menu.childNodes[i].className="sec1";
    menu.childNodes[n].className="sec2";
    for(i=0;i<bottom_nav_text.childNodes.length;i++)
    bottom_nav_text.childNodes[i].style.display="none";
    bottom_nav_text.childNodes[n].style.display="block";
    }
    </script>
</head>

<body>

<!--顶部开始-->
	<!--#include file="../inc/header.asp"-->
<!--顶部结束-->

<!--容器开始-->
	<div id="container">
<!--主体开始-->
		<div id="main">
		
<!--主体上部flash开始-->
			<!--#include file="../inc/main_boxtop.html"-->
<!--主体上部flash结束-->

<!--主体左侧开始-->
			<!--#include file="../inc/product_left_nav.html"-->
<!--主体左侧结束-->

<!--主体右侧开始-->
			<div id="main_right">
				<div id="contact">
					<div id="contact_left">
						<div id="contact_left_title">产品展示 &nbsp; Product</div>
						<div id="contact_left_box">
														
						<%Sub ProductsList(InfoType,PageNum)%>
						<%
						sqlstr="select * from [Products] where status=1 and Product_Type = " & InfoType & " order by show_order desc,id desc"
											
						Dim rs_Products : Set rs_Products = db.getRecordBySQL(sqlstr)
						
						rs_Products.PageSize=PageNum
					
						pre = true
						last = true
						page = trim(Request.QueryString("page"))
						  
						if len(page) = 0 then
							intpage = 1
							pre = false
						else
							if cint(page) =< 1 then
								intpage = 1
								pre = false
							else
								if cint(page) >= rs_Products.PageCount then
									intpage = rs_Products.PageCount
									last = false
								else
									intpage = cint(page)
								end if
							end if
						end if
						
						if not rs_Products.eof then
							rs_Products.AbsolutePage = intpage
						end if
					
						if not (rs_Products.eof or rs_Products.bof) then
							do until rs_Products.eof
						%>
						
						<div class="honor_pic">
                        <a href="productdetail.asp?id=<%=rs_Products("ID")%>">
                        <img src="<%=Config.ImgUrl()%>Pictures/<%=rs_Products("Pic_View")%>" width="150" height="150" /></a>
                        <span><a href="productdetail.asp?id=<%=rs_Products("ID")%>"><%=rs_Products("Product_Name")%></a></span>
                        </div>
						<%
							rs_Products.movenext
							loop
						else
							
						%>
							没有数据
						<%
						end if											
						%>
                        
						<div style="clear:both;"></div>
						
                        <div id="page">
						<%if rs_Products.pagecount > 0 then%>
							共<%=rs_Products.recordcount%>款产品&nbsp;当前页<%=intpage%>/<%=rs_Products.PageCount%>
						<%else%>
							当前页0/0
						<%end if%>
							<a href="index.asp?page=1&type=<%=InfoType%>">首页</a>| 
						<%if pre then%>
							<a href="index.asp?page=<%=intpage -1%>&type=<%=InfoType%>">上页</a>| 
						<%end if%>
						<%if last then%>
							<a href="index.asp?page=<%=intpage +1%>&type=<%=InfoType%>">下页</a>| 
						<%end if%>
							<a href="index.asp?page=<%=rs_Products.PageCount%>&type=<%=InfoType%>">尾页</a>|转到第 
							<select name="sel_page" onchange="javascript:location=this.options[this.selectedIndex].value;">
						<%
						for i = 1 to rs_Products.PageCount
						if i = intpage then
						%>
							<option value="index.asp?page=<%=i%>&type=<%=InfoType%>" selected><%=i%></option>
						<%else%>
							<option value="index.asp?page=<%=i%>&type=<%=InfoType%>"><%=i%></option>
						<%
						end if
						next
						%>
						</select>
						页
						</div>
						
						<%
						db.C(rs_Products)
						%>
						<%end sub%>
                        
                        <%
						Function InfoTypeReceive(ParentType)
						  InfoType=request.QueryString("type")
		  
						  if InfoType = "" or InfoType = 0  then
						  		
							  sqlstr="select ID from [Product_Type] where status=1 and parent_id = " & ParentType & " order by show_order desc,id desc"
											  
							  Dim rs_InfoType : Set rs_InfoType = db.getRecordBySQL(sqlstr)
							  
							  if not (rs_InfoType.eof or rs_InfoType.bof) then
								  InfoType = rs_InfoType("id")
							  else
								  InfoType = 0
							  end if	
							  
							  db.C(rs_InfoType)
							  
						  end if
						  
						  InfoType=cint(InfoType)
						  
						  InfoTypeReceive = InfoType
						
						End Function
						
						%>
                        <%
						ProductsType = InfoTypeReceive(1)
						'response.Write(ProductsType)
						%>
                        <%Call ProductsList(ProductsType,12)%>
                            
                            <div style="clear:both;"></div>
						</div>
					</div>
					<div id="contact_right">
						<div class="contact_right_pic"><img src="../images/05.jpg" width="125" height="75" /></div>
						<div class="contact_right_pic"><img src="../images/06.jpg" width="125" height="75" /></div>
						<div class="contact_right_pic"><img src="../images/15.jpg" width="125" height="75" /></div>
					</div>
					<div style="clear:both;"></div>
				</div>
			</div>	
<!--主体右侧结束-->

<!--主体底部开始-->

			<!--#include file="../inc/main_bottom.asp"-->

<!--主体底部结束-->
			
		</div>	
<!--主体结束-->

	</div>
<!--容器结束-->			

<!--尾部容器开始-->
	<!--#include file="../inc/footer.html"-->
<!--尾部容器结束-->
</body>
</html>
<%
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>