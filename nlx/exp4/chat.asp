<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Chat</title>
</head>
<body>
    <%
    Session("user_name") = Request("user_name")
    
    Dim saystr
    saystr = "Visit from " & Request.ServerVariables("REMOTE_ADDR")
    saystr = saystr & <b> Session("user_name") & </b>
    saystr = saystr & " at " & Now()

    Application.Lock
    Application("show") = saystr & "<br>" & Application("show")
    application("user_online") = application("user_online") + 1
    Application.UnLock
    %>
    <frameset rows="80%,20%">
        <frame src="chat.asp" name="chat">
        <frame src="input.asp" name="input">
        <noframes>
            <body>
                <p>This page uses frames, but your browser doesn't support them.</p>
            </body>
        </noframes>
    </frameset>
</body>
</html>