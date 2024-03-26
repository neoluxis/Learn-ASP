<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Home</title>
</head>
<body>
    <%
        dim lb, username
        username = Request.Cookies("c_name")
        lb = request.QueryString("jr")
        If (lb = "0") Then
            Response.Write("Welcome " & username & "! <br/>")
            Response.Write("This is your first visit to our site.")
        Else
            Response.Write("Welcome back " & username & "!<br/>")
            Response.Write("You have visited our site " & lb & " times.")
        End If
    %>
</body>
</html>