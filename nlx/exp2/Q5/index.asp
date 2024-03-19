<html>
<body>
	<%
	Dim strA
	strA = Split("VBSc*ript*is*go*od!", "*") ' count number of stars
    Response.Write "The number of stars in the string is: " & UBound(strA) + 1
	%>
</body> 
</html> 
