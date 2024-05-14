<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Chattig Room</title>
    <Script Language="'JavaScript">
        // <!--
        function check_null() {
            if (document.form1.user_name.value == "") {
                alert("Please enter your name!");
                document.form1.user_name.focus();
                return false;
            }
            return true;
        }    
        // -->
    </Script>
</head>
<body>
    <h2>Kleine Chatting Room</h2>
    <center>
        Now there are <%=Application("user_online")%> people online.
        <form method="post" action="chat.asp" name="form1" onsubmit="return check_null()">
            User Name: <input type="text" name="user_name" size="20">
            <input type="submit" value="Enter">
        </form>
    </center>
</body>
</html>