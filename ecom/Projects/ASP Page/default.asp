
<HTML>
<H1>Anoop's Socket Object Test </H1>
<form method="POST" action="../asp%20page/default.asp">
  <table border="0" width="100%">
    <tr>
      <td width="24%">Data To Send</td>
      <td width="76%"><input type="text" name="data" size="20" value="Your data"></td>
    </tr>
    <tr>
      <td width="24%">Remote IP</td>
      <td width="76%"><input type="text" name="ip" size="20" value="localhost"></td>
    </tr>
    <tr>
      <td width="24%">Remote Port</td>
      <td width="76%"><input type="text" name="port" size="20" value="4000"></td>
    </tr>
  </table>
  <p><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></p>
</form>

<%

'Get the form contents
ip=request.form("ip")
port=request.form("port")
data=request.form("data")

if ip<>"" and port<>"" and data<>"" then

'Simply declaring a variable.
Dim mySock 

'Create an instance
set mySock=server.createobject("sockobject.socket")

'Now attempt to connect. Connect method
'Param 1: Host Name
'Param 2: Port

Result=mySock.connect(ip,port)


'If result is successfull, send that data

'Param 1: Data to send
'Param 2: Timeout in Seconds

'Timeout is the timedelay our COM component may wait, for getting the result back.

Mydata="Your Data" 'Construct the data string
if Result="NOERROR" then
	Ret=MySock.SendData(Data,10)
	response.write "Result="  & ret
end if



'The socket will be closed automatically, terminating the class
set mySock=nothing

end if

%>

</HTML>
