<script runat="server">
    If (Not IsObject(Session("playerName")) Then
    Session("playerName") = Request.Form("playerName")
End If

If Not IsObject(Session("randomNumber")) Then
    Randomize
    Session("randomNumber") = Int((100 * Rnd) + 1) ' Random number between 1 and 100
End If

Dim playerName, guess, message
playerName = Session("playerName")
guess = Request.Form("guess")
message = "" ' Initialize message
</script>
 

<% @Language="VBScript" 

If IsEmpty(Session("playerName")) And Request.Form("playerName") <> "" Then 
    Session("playerName") = Request.Form("playerName") 
End If 

  

If IsEmpty(Session("randomNumber")) Then 
    Randomize 
    Session("randomNumber") = Int((100 * Rnd) + 1) ' Random number between 1 and 100 
End If
Dim playerName, guess, message 
playerName = Session("playerName") 
guess = Request.Form("guess") 
message = "Initialize message" 
%>