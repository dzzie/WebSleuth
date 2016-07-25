

<h1>Query String Args</h1><br><hr>
<%

for i=1 to Request.QueryString.Count 

Response.write Request.QueryString.Key(i) & "=" & unescape(Request.QueryString.item(i)) & "<BR>"
next

%>
<h1>Form Args</h1><br><hr>
<%
for i=1 to Request.Form.Count 
Response.write Request.Form.key(i) & "=" &  unescape(Request.Form.Item(i)) & "<BR>"
next
%>
<br><hr><br><br>

<form name=fred action="test.asp" method=get>
<input type=text name=one>
<input type=text name=two>
<input type=text name=three>
</form>

<a href="http://localhost/test.asp?one=a&two=b&three=c">cgi href to test</a>
