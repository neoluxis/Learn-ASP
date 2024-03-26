<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Register</title>
  </head>
  <body>
    <h2>Register</h2>
    <form method="POST" action="verify.asp" name="myform">
      <table border="0" align="center">
        <tr>
          <td>Username:</td>
          <td><input type="text" name="username" size="20" /></td>
        </tr>
        <tr>
          <td>Password:</td>
          <td><input type="password" name="password" size="20" /></td>
        </tr>
        <tr>
          <td>Confirm Password:</td>
          <td><input type="password" name="confirm_password" size="20" /></td>
        </tr>
        <tr>
          <td>Email:</td>
          <td><input type="text" name="email" size="20" /></td>
        </tr>
        <tr>
          <td>Phone:</td>
          <td><input type="text" name="phone" size="20" /></td>
        </tr>
        <tr>
          <td colspan="2">
            <p align="center">
              <input type="submit" value="Submit" name="B1" />
              &nbsp;&nbsp;
              <input type="reset" value="Reset" name="B2" />
            </p>
          </td>
        </tr>
      </table>
    </form>
    <% %>
  </body>
</html>
