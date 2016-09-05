Option Explicit

Function save_settings(lo As ListObject) As Collection

    Set save_settings = New Collection
    With save_settings
        .Add Key:="ShowHeaders", Item:=lo.ShowHeaders
        .Add Key:="ShowTotals", Item:=lo.ShowTotals
        .Add Key:="ShowAutoFilter", Item:=lo.ShowAutoFilter
    End With

End Function

Sub load_settings(lo As ListObject, saved_settings As Collection)

    lo.ShowHeaders = saved_settings("ShowHeaders")
    lo.ShowTotals = saved_settings("ShowTotals")
    lo.ShowAutoFilter = saved_settings("ShowAutoFilter")

End Sub

Sub resize_databodyrange(lo As ListObject, row_size As Long)

    Dim saved_lo_settings As Collection
    Set saved_lo_settings = save_settings(lo)

    With lo

        If Not .DataBodyRange Is Nothing Then
            ' Need headers on - otherwise may delete table when DataBodyRange is deleted
            .ShowHeaders = True
            .DataBodyRange.Delete
        End If
        
        If row_size > 0 Then
            .ListRows.Add
            .ShowHeaders = False
            .ShowTotals = False
            .Resize .Range.Resize(RowSize:=row_size)
        End If

    End With
    
    Call load_settings(lo, saved_settings:=saved_lo_settings)

End Sub
