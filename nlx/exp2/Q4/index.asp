<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Document</title>
  </head>
  <body>
    <!-- Use VBscript to show calender of this month -->
    <table border="1">
      <tr>
        <th>Mon</th>
        <th>Tue</th>
        <th>Wed</th>
        <th>Thu</th>
        <th>Fri</th>
        <th>Sat</th>
        <th>Sun</th>
      </tr>
      <%
        Dim d, m, y, fd, ld, da
        d = Date()
        m = Month(d)
        y = Year(d)
        fd = Weekday(DateSerial(y, m, 1))
        ld = Day(DateSerial(y, m + 1, 0))
        da = 1
        
        For i = 0 To 5 Step 1
            Response.Write("<tr>")
            For j = 0 To 6 Step 1
                If i = 0 And j < fd-2 Then
                Response.Write("<td></td>")
                ElseIf da > ld Then
                Exit For
                Else
                Response.Write("<td>" & da & "</td>")
                da = da + 1
                End If
            Next
            Response.Write("</tr>")
        Next
      %>
    </table>
  </body>
</html>
