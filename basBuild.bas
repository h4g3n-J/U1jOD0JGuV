Attribute VB_Name = "basBuild"
Option Compare Database
Option Explicit

' builds the application form scratch
' work in progress
Public Function BuildApplication()
    
    ' command message
    If gconVerbatim Then
        Debug.Print "execute basBuild.BuildAppliation"
    End If
    
    ' build querys
    basBuild.BuildQryAngebotAuswahl
    basBuild.BuildQryAngebot
    
    ' build forms
    basAngebotSuchenSub.BuildAngebotSuchenSub
    basAngebotSuchen.BuildAngebotSuchen
    
    basAuftragSuchenSub.BuildAuftragSuchenSub
    basAuftragSuchen.BuildAuftragSuchen
    
    basRechnungSuchenSub.BuildRechnungSuchenSub
    basRechnungSuchen.BuildRechnungSuchen
    
    ' open frmHauptmenue
    DoCmd.OpenForm "frmHauptmenue", acNormal
    
    ' event message
    If gconVerbatim Then
        Debug.Print "basBuild.BuildAppliation executed"
    End If
    
End Function

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

' build qryAngebot
Private Sub BuildQryAngebot()
    
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
