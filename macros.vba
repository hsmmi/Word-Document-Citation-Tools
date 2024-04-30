Sub LinkCitationsToReferences()
    Dim refItem As Paragraph
    Dim doc As Document
    Dim refNum As String
    Dim bmName As String
    Dim citation As Range
    Dim refPattern As String
    Dim startPos As Integer
    Dim endPos As Integer

    Set doc = ActiveDocument

    ' Remove existing citation hyperlinks and bookmarks
    RemoveCitationHyperlinksAndBookmarks

    ' Create bookmarks for each reference in the bibliography section
    For Each refItem In doc.Paragraphs
        If Left(refItem.Range.Text, 1) = "[" And InStr(refItem.Range.Text, "]") > 0 Then
            refNum = Mid(refItem.Range.Text, 2, InStr(refItem.Range.Text, "]") - 2)
            bmName = "Ref_" & refNum
            doc.Bookmarks.Add Name:=bmName, Range:=refItem.Range
        End If
    Next refItem

    ' Link any matching number in the text to the corresponding bookmark
    For Each citation In doc.StoryRanges
        With citation.Find
            .ClearFormatting
            .Text = "\[[0-9]{1,}\]" ' Matches numbers within brackets
            .MatchWildcards = True
            While .Execute
                refNum = Mid(citation.Text, 2, Len(citation.Text) - 2)
                bmName = "Ref_" & refNum
                If doc.Bookmarks.Exists(bmName) Then
                    startPos = citation.Start + 1
                    endPos = citation.End - 1
                    doc.Hyperlinks.Add Anchor:=doc.Range(startPos, endPos), Address:="", SubAddress:=bmName
                End If
                citation.Collapse wdCollapseEnd
            Wend
        End With
    Next citation
End Sub



Sub RemoveCitationHyperlinksAndBookmarks()
    Dim doc As Document
    Dim i As Integer
    Dim bmName As String

    Set doc = ActiveDocument

    ' Remove hyperlinks that point to citation bookmarks
    For i = doc.Hyperlinks.Count To 1 Step -1
        bmName = doc.Hyperlinks(i).SubAddress
        If bmName Like "Ref_*" Then ' Check if the hyperlink points to a bookmark starting with "Ref_"
            doc.Hyperlinks(i).Delete
        End If
    Next i

    ' Remove bookmarks related to citations
    For i = doc.Bookmarks.Count To 1 Step -1
        If doc.Bookmarks(i).Name Like "Ref_*" Then ' Check if the bookmark name starts with "Ref_"
            doc.Bookmarks(i).Delete
        End If
    Next i
End Sub

