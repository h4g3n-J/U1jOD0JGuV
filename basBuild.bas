Attribute VB_Name = "basBuild"
Option Compare Database
Option Explicit

' builds the application form scratch
' work in progress
Public Function BuildApplication()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basBuild.BuildAppliation ausfuehren"
    End If
    
    ' build querys
    basBuild.BuildQryAngebotAuswahl
    basBuild.BuildQryAngebot
    
    ' build forms
    basBuild.BuildForms
    basAngebotSuchen.BuildAngebotSuchen
End Function

' build qryAngebotAuswahl
Public Sub BuildQryAngebotAuswahl(Optional varSearchTerm As Variant = "*")
    
    ' NULL handler
    If IsNull(varSearchTerm) Then
        varSearchTerm = "*"
    End If
        
    ' transform to string
    Dim strSearchTerm As String
    strSearchTerm = CStr(varSearchTerm)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basBuild.BuildQryAngebotAuswahl ausfuehren"
    End If
    
    ' define query name
    Dim strQueryName As String
    strQueryName = "qryAngebotAuswahl"
    
    ' set current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
            
    ' delete existing query of the same name
    basBuild.DeleteQuery strQueryName
    
    ' set query
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
    
    With qdfQuery
        ' set query Name
        .Name = strQueryName
        ' set query SQL
        .SQL = SqlQryAngebotAuswahl(strSearchTerm)
    End With
    
    ' save query
    With dbsCurrentDB.QueryDefs
        .Append qdfQuery
        .Refresh
    End With

ExitProc:
    qdfQuery.Close
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
    Set qdfQuery = Nothing
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basBuild.BuildQryAngebotAuswahl ausgeführt"
    End If
    
End Sub

' build qryAngebot
Public Sub BuildQryAngebot()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basBuild.BuildQryAngebot ausfuehren"
    End If
    
    ' define query name
    Dim strQueryName As String
    strQueryName = "qryAngebot"
    
    ' set current database
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb
            
    ' delete existing query of the same name
    basBuild.DeleteQuery strQueryName
    
    ' set query
    Dim qdfQuery As DAO.QueryDef
    Set qdfQuery = dbsCurrentDB.CreateQueryDef
    
    ' set query Name
    qdfQuery.Name = strQueryName
    
    ' set query SQL
    qdfQuery.SQL = SqlQryAngebot
    
    ' save query
    With dbsCurrentDB.QueryDefs
        .Append qdfQuery
        .Refresh
    End With

ExitProc:
    qdfQuery.Close
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
    Set qdfQuery = Nothing
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basBuild.BuildQryAngebotAuswahl ausgeführt"
    End If
    
End Sub

' delete query
' 1. check if query exists
' 2. close if query is loaded
' 3. delete query
Private Sub DeleteQuery(strQueryName As String)
    
    ' command message
    If gconVerbatim Then
        Debug.Print "basSupport.DeleteQuery ausfuehren"
    End If
    
    ' set dummy object
    Dim objDummy As Object
    ' search object list >>AllQueries<< for strQueryName
    For Each objDummy In Application.CurrentData.AllQueries
        If objDummy.Name = strQueryName Then
            
            ' check if query isloaded
            If objDummy.IsLoaded Then
                ' close query
                DoCmd.Close acQuery, strQueryName, acSaveYes
                ' verbatim message
                If gconVerbatim Then
                    Debug.Print "basSupport.DeleteQuery: " & strQueryName & " ist geoeffnet, Abfrage geschlossen"
                End If
            End If
    
            ' delete query
            DoCmd.DeleteObject acQuery, strQueryName
            
            ' event message
            If gconVerbatim = True Then
                Debug.Print "basSupport.DeleteQuery: " & strQueryName & " existierte bereits, Abfrage geloescht"
            End If
            
            ' exit loop
            Exit For
        End If
    Next
    
    If gconVerbatim Then
        Debug.Print "basBuild.DeleteQuery ausgeführt"
    End If
    
End Sub
    
Private Function SqlQryAngebot() As String
    
    If gconVerbatim Then
        Debug.Print "basBuild.SqlQryAngebot ausfuehren"
    End If
    
    SqlQryAngebot = " SELECT tblAngebot.*" & _
            " FROM tblAngebot" & _
            " ;"
End Function

Private Function SqlQryAngebotAuswahl(strSearchTerm As String)
    
    If gconVerbatim Then
        Debug.Print "basBuild.SqlQryAngebotAuswahl ausfuehren"
    End If
    
    SqlQryAngebotAuswahl = " SELECT qryAngebot.*" & _
            " FROM qryAngebot" & _
            " WHERE qryAngebot.BWIKey LIKE '*" & strSearchTerm & "*'" & _
            " ;"
End Function

Private Sub BuildForms()
    
    If gconVerbatim Then
        Debug.Print "basBuild.Forms ausfuehren"
    End If
    
    ' build Hauptmenue
    ' basHauptmenue.BuildFormHauptmenue
    
    ' build subformular AngebotSuchenSub
    basAngebotSuchenSub.BuildFormAngebotSuchenSub
    
    ' build AngebotSuchen
    basAngebotSuchen.BuildAngebotSuchen
    
    ' open frmHauptmenue
    DoCmd.OpenForm "frmHauptmenue", acNormal
End Sub
