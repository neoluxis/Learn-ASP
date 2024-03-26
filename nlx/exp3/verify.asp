<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
</head>
<body>
    <%
        Dim errmsg
        errmsg = ""
        If (Trim(Request("username"))="") Then
            errmsg = "Please enter your username"
        End If
        if (Trim(Request("password"))="" or len(trim(request("password"))) < 6) Then
            errmsg = "Please enter your password and it must be at least 6 characters"
        End If
        if (trim(request("password")) <> trim(request("confirm_password"))) Then
            errmsg = "Passwords do not match"
        End If
        if (instr(Request("email"), "@")=0) then
        errmsg = "Please enter a valid email address"
        End If
        if (request("phone")<>"" and not isnumeric(request("phone"))) then
        errmsg = "Please enter a valid phone number"
        End If

        if (errmsg<>"") then
            response.write(errmsg & "<p>Please go <a href='login.asp'>back</a> and check your info</p>")
        else
            Response.Cookies("c_name") = Trim(Request("username"))
            Response.Cookies("c_name").expires = date() + 3*365
            Response.Write("You passed the varification! ")
        end if
    %>
</body>
</html>