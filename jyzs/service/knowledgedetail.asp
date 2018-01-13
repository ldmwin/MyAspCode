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
	<link href="../style/news.css" rel="stylesheet" type="text/css" />
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
			<!--#include file="../inc/service_left_nav.html"-->
<!--主体左侧结束-->

<!--主体右侧开始-->
			<div id="main_right">
				<div id="contact">
					<div id="contact_left">
					  <div id="contact_left_box">
                       	  <%Call InfoDetail()%>
				
						<%
                            Sub InfoDetail()
                            
                            Info_id=request.QueryString("id")
                    
                            if Info_id="" or Info_id=0 then
                                response.write "<script>alert('参数出错');history.go(-1);</Script>"
                                response.end()
                            end if
                            
                            Info_id=cint(Info_id)
                            
                            
                            sqlstr = "select * from Informations where (status=1 or status=5) and id = " & Info_id
                            
                            Dim rs_Infos : Set rs_Infos = db.getRecordBySQL(sqlstr)
                            
                            if rs_Infos.eof or rs_Infos.bof then
                                response.write "<script>alert('数据查询失败');history.go(-1);</Script>"
                                response.End()
                            end if	
                            %>
                            
 
                                    <div class="newsdetail_title"><%=rs_Infos("Title")%></div>
                                
                                    <div class="newsdetail_time"><%response.Write(formatdatetime(rs_Infos("Add_Time"),2))%></div>
                                
                                    <div class="newsdetail_text">
                                    
                                      <%=rs_Infos("Content")%>
            
                                    </div>
                                                                    
                            
                        <%
                            db.C(rs_Infos)
                            
                            End Sub
                            %>
							
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

			<!--#include file="../inc/main_bottom.html"-->

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