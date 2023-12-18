Sub CountComments()
    Dim UnresolvedCommentsCount As Integer
    Dim ResolvedCommentsCount As Integer
    UnresolvedCommentsCount = 0
    ResolvedCommentsCount = 0

    Dim aComment As Comment
    For Each aComment In ActiveDocument.Comments
        If aComment.Done Then
            ResolvedCommentsCount = ResolvedCommentsCount + 1
        Else
            UnresolvedCommentsCount = UnresolvedCommentsCount + 1
        End If
    Next aComment

    MsgBox "There are " & UnresolvedCommentsCount & " unresolved comments and " & ResolvedCommentsCount & " resolved comments in this document."
End Sub
