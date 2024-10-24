
<%
Language="VBScript" 

If IsEmpty(Session("playerName")) And Request.Form("playerName") <> "" Then 
    Session("playerName") = Request.Form("playerName") 
End If 

If IsEmpty(Session("randomNumber")) Then 

    Randomize  
    Session("randomNumber") = Int((100 * Rnd) + 1)

End If 

Dim playerName, guess, message 
playerName = Session("playerName") 
guess = Request.Form("guess") 
message = "" ' Initialize message 

%> 
