<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Index</title>
</head>
<body>
    <%
        Response.Cookies("c_name").expires="2005-01-01"
        if (Request.Cookies("c_name")<>"") then
        Response.write("<a href='home.asp?jr=1'>进入主页</a>")
        else
        Response.redirect("login.asp")
        end if
    %>
</body>
</html>