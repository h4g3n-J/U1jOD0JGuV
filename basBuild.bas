Attribute VB_Name = "basBuild"
Option Compare Database
Option Explicit

' builds the application form scratch
' work in progress
Public Function BuildApplication()
    
    If gconVerbatim Then
        Debug.Print "basBuild.BuildAppliation ausfuehren"
    End If
    
    ' build querys
    basBuild.BuildQuerys
    
    ' build forms
    basBuild.BuildForms
    basAngebotSuchen.BuildAngebotSuchen
End Function

Private Sub BuildQuerys()
    
    If gconVerbatim Then
        Debug.Print "basBuild.BuildQuerys ausfuehren"
    End If
    
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
        basSupport.ClearQuery aQuerySet(0, inti)
        
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
    
    If gconVerbatim Then
        Debug.Print "basBuild.SqlQryAngebot ausfuehren"
    End If
    
    SqlQryAngebot = " SELECT tblAngebot.*" & _
            " FROM tblAngebot" & _
            " ;"
End Function

Private Function SqlQryAngebotAuswahl()
    
    If gconVerbatim Then
        Debug.Print "basBuild.SqlQryAngebotAuswahl ausfuehren"
    End If
    
    SqlQryAngebotAuswahl = " SELECT qryAngebot.*" & _
            " FROM qryAngebot" & _
            " ;"
End Function

Private Sub BuildForms()
    
    If gconVerbatim Then
        Debug.Print "basBuild.Forms ausfuehren"
    End If
    
    ' build Hauptmenue
    ' basHauptmenue.BuildFormHauptmenue
    
    ' build AngebotSuchenSub
    basAngebotSuchenSub.BuildFormAngebotSuchenSub
    
    ' build AngebotSuchen
    basAngebotSuchen.BuildAngebotSuchen
    
    ' open frmHauptmenue
    DoCmd.OpenForm "frmHauptmenue", acNormal
End Sub
