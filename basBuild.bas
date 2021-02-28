Attribute VB_Name = "basBuild"
Option Compare Database
Option Explicit

' builds the application form scratch
' work in progress
Public Sub BuildApplication()
    ' build querys
    qryAngebot
    
    ' build frmHauptmenue
    basHauptmenue.CreateFormHautpmenue
End Sub

Private Sub buildQuerys()
    
    Dim aQuerySet(1, 1) As Variant
    ' 0 = query name
    ' 1 = SQL-source
    aQuerySet(0, 0) = "qryAngebot"
    aQuerySet(1, 0) = basBuild.SqlQryAngebot
    aQuerySet(0, 1) = "qryAngebotAuswahl"
    aQuerySet(1, 1) = basBuild.SqlQryAngebotAuswahl
       
    Dim dbsCurrentDB As DAO.Database
    Set dbsCurrentDB = CurrentDb

    Dim qdfQuery As DAO.QueryDef
    
    Dim objQuery As Object
    Dim inti As Integer
        
    For inti = LBound(aQuerySet, 2) To UBound(aQuerySet, 2)
        
        ' delete existing query of the same name
        For Each objQuery In dbsCurrentDB.QueryDefs
            If objQuery.Name = aQuerySet(0, inti) Then
                DoCmd.DeleteObject acQuery, aQuerySet(0, inti)
                ' verbatim message
                Debug.Print "basBuild.buildQuerys: " & aQuerySet(0, inti) _
                    ; " existiert bereits, Objekt geloescht"
            End If
        Next objQuery
        
        ' write SQL to query
        Set qdfQuery = dbsCurrentDB.CreateQueryDef
        With qdfQuery
            .SQL = aQuerySet(1, inti)
            .Name = aQuerySet(0, inti)
        End With
    
        ' save query
        With dbsCurrentDB.QueryDefs
            .Append qdfQuery
            .Refresh
        End With
    
        ' verbatim message
        If gconVerbatim Then
            Debug.Print "basBuild.buildQuerys ausgefuehrt, " & aQuerySet(0, inti) & " erstellt"
        End If
    
        qdfQuery.Close
    Next
    
ExitProc:
    dbsCurrentDB.Close
    Set dbsCurrentDB = Nothing
    Set qdfQuery = Nothing
End Sub
    
Private Function SqlQryAngebot() As String
    SqlQryAngebot = " SELECT tblAngebot.*" & _
            " FROM tblAngebot" & _
            " ;"
End Function

Private Function SqlQryAngebotAuswahl()
    SqlQryAngebotAuswahl = " SELECT qryAngebot.*" & _
            " FROM qryAngebot" & _
            " ;"
End Function
