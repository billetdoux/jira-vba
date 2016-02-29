Option Explicit

Private cookie As String
Private url As String
Private tracing As Boolean
Private xmlHttp As New MSXML2.XMLHTTP60
Private jsonParser As New jsonlib

Public Sub Init(jiraUrl As String, tracingOn As Boolean)
    url = jiraUrl
    tracing = tracingOn
End Sub

Public Function LogIn() As Boolean
    If cookie = "" Then
        traceDebug "logging in to " & url
        With xmlHttp
            .Open "POST", url & "/jira/rest/auth/1/session", False
            .setRequestHeader "Content-Type", "application/json"
            .setRequestHeader "Accept", "application/json"
            .setRequestHeader "Authorization", "Basic"
            .send ""
            
            traceDebug "http status=" & CStr(.Status)
            If .Status <> 401 Then
                traceDebug "logged in"
                LogIn = True
                cookie = "set"
            Else
                traceDebug "failed to log in"
                LogIn = False
            End If
        End With
    Else
        traceDebug "already logged in"
        LogIn = True
    End If
End Function

Public Sub LogOut()
    cookie = ""
End Sub

' returns Dictionary, if jira is found
' returns Nothing, if jira is not found
Public Function GetJira(issue As String) As Object
    With xmlHttp
        .Open "GET", url & "/rest/api/latest/issue/" & issue, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Accept", "application/json"
        .send
        
        traceDebug "http status=" & CStr(.Status)
        If .Status = 200 Then
            traceDebug .responseText
            Set GetJira = jsonParser.parse(.responseText)
        Else
            Set GetJira = Nothing
        End If
    End With
End Function

Public Function CreateStory(ByVal projKey, ByVal summary As String, ByVal epic As String) As Object
    ' TODO understand how to add epic
    Dim jsonTxt As String
    jsonTxt = "{""fields"":{""project"":{""key"":""" & projKey & """},""issuetype"":{""name"":""Story""},""summary"":""" & summary & """ ,""description"" :"""" , ""priority"":{""name"":""Lowest""},""components"":[]}}"
    Set CreateStory = PostIssue(jsonTxt)
End Function

Public Function CreateEpic(ByVal projKey, ByVal summary As String) As Object
    Dim jsonTxt As String
    jsonTxt = "{""fields"":{""project"":{""key"":""" & projKey & """},""issuetype"":{""name"":""Epic""},""summary"":""" & summary & """ ,""customfield_10005"":""" & summary & """ , ""description"" :"""" , ""priority"":{""name"":""Lowest""},""components"":[]}}"
    Set CreateEpic = PostIssue(jsonTxt)
End Function

Public Function PostIssue(ByVal jSon As String) As Object
    traceDebug jSon
    With xmlHttp
       .Open "POST", url & "/rest/api/2/issue/", False
       .setRequestHeader "Content-Type", "application/json"
       .setRequestHeader "Accept", "application/json"
       .setRequestHeader "Authorization", "Basic"
       .setRequestHeader "X-Atlassian-Token", "nocheck"
       .send jSon
       
       If .Status <> 401 Then
            traceDebug .responseText
            Set PostIssue = jsonParser.parse(.responseText)
       Else
            Set PostIssue = Nothing
       End If
    End With
End Function

Private Sub traceDebug(msg As String)
    trace "Debug", msg
End Sub

Private Sub traceWarn(msg As String)
    trace "Warn", msg
End Sub

Private Sub traceErr(msg As String)
    trace "Error", msg
End Sub

End Sub

Private Sub trace(topic As String, msg As String)
    If tracing Then
        Debug.Print topic & ": " & msg
    End If
End Sub
