Attribute VB_Name = "NewMacros"
Sub totalCount()

    ' dims
    Dim secCount As Integer
    Dim section As Integer
    Dim summary As String
    Dim wordCount As Integer
    Dim preChap As Integer
    Dim titleCount As Integer
    
    ' unique variables
    preChap = 2
    titleCount = 2
    
    ' responsive variables
    secCount = ActiveDocument.Sections.Count
    summary = "Word Count" & vbCrLf & "________________" & vbCrLf & vbCrLf

    ' loops
    For section = 1 To secCount
        If (section - preChap) > 0 Then
            wordCount = ActiveDocument.Sections(section).Range.ComputeStatistics(wdStatisticWords)
            
            If (wordCount - titleCount) > 0 Then
                summary = summary & "Chapter " & (section - preChap) & ": " _
                  & (wordCount - titleCount) _
                  & vbCrLf & vbCrLf
            End If
            
        End If
    Next

    ' output
    MsgBox prompt:=summary, Title:="Word Counter"

End Sub

Sub chapCount()

    ' dims
    Dim inputValue As Integer
    Dim secCount As Integer
    Dim section As Integer
    Dim summary As String
    Dim wordCount As Integer
    Dim preChap As Integer
    Dim titleCount As Integer
    Dim Msg, Style, Title, Response As String
    Dim Msg2, Style2, Title2, Response2 As String
    
    ' unique variables
    preChap = 2
    titleCount = 2
    
    ' msgbox variables
    Style = vbExclamation
    Style2 = vbInformation
    Title = "Error"
    Title2 = "Word Count"
    Msg = "Nonexistent chapter"
    Msg2 = "No word count"
    
    'responsive variables
    secCount = ActiveDocument.Sections.Count
    
    'InputBox
    inputValue = inputBox("Enter chapter number")
    
    
    ' Chapter Doesn't Exist
    If (inputValue) > (secCount - preChap) Then
        Response = MsgBox(Msg, Style, Title)
    End If
    
    ' Loop data
    ' Chapter Does Exist
    For section = 1 To secCount
        If (section - preChap) = inputValue Then
        
            wordCount = ActiveDocument.Sections(section).Range.ComputeStatistics(wdStatisticWords)
            summary = "Chapter " & (section - preChap) & ": " _
              & (wordCount - titleCount)
            
            ' Section Exists but no Title or Body
            If (wordCount - titleCount) < 0 Then
                Response = MsgBox(Msg, Style, Title)
                
            ' Section Exists with Title but no Body
            ElseIf (wordCount - titleCount) = 0 Then
                Response2 = MsgBox(Msg2, Style2, Title2)
                
            ' Section Exists with Title and Body
            ElseIf (wordCount - titleCount) > 0 Then
                MsgBox prompt:=summary, Title:="Word Counter"
            End If
            
        End If
    Next

End Sub


Sub pageBreaker()

    ' variables
    Dim section As Integer
    Dim preChap As Integer
    Dim modifier As Integer
    Dim correctedPreChap As Integer
    Dim chapNumber As Integer
    
    ' unique variables
    preChap = 2
    
    ' standard variables
    modifier = 1
    correctedPreChap = preChap - modifier
    section = ActiveDocument.Sections.Count
    
    ' end of section
    Selection.EndKey Unit:=wdStory
    
    ' insert page break
    Selection.InsertBreak Type:=wdSectionBreakNextPage
    
    ' end of document
    Selection.EndKey Unit:=wdStory
    
    ' insert text
    chapNumber = section - correctedPreChap
    
    ActiveDocument.Content.InsertAfter Text:="Chapter " & chapNumber
    Selection.Style = "Heading 1"
    
    ' end of document
    Selection.EndKey Unit:=wdStory
    
End Sub

Sub expander()

    ActiveDocument.ActiveWindow.View.ExpandAllHeadings

End Sub

Sub collapser()

    ActiveDocument.ActiveWindow.View.CollapseAllHeadings

End Sub
