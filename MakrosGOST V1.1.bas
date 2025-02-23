Attribute VB_Name = "MakrosGOST"
Sub MacrosGost()
    
    ReplaceLinesCarrets1
    ReplaceSpaces2
    
    ReplacePages3
    DelExcessCarrets4
    ReplaceHeaders5
    ReplaceBulletsWithDash6
    CheckListMarkers7
    
    SetAllIndentsToZero8
    Numbering9
    Fields10
    ChangeFont11
    SetLineSpacingToOnePointFive12
    ParagraphIndent13
    ResizeTablesToWindowWidth14
    Reddots15
    
    ReplaceWord16

    
    AlignJustify17
    CenterAlignIfImageOrTable18
    CenterAllImages19
    FormatTables20
    CheckHeaders21
    PageHeaders22
    ReplaceBulletsNumbers23
    RedName24
    Formula25
    BlackLiterature26
End Sub

Sub ReplaceLinesCarrets1()
    selection.Find.ClearFormatting
    selection.Find.Replacement.ClearFormatting
    With selection.Find
        .Text = Chr(11)
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    selection.Find.Execute Replace:=wdReplaceAll
End Sub

Sub ReplaceSpaces2()

    Dim doc As Document
    Dim rng As Range
    Dim firstPageEnd As Long
    Dim firstPage As Range
   
    Set doc = ActiveDocument
    
    Set firstPage = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=2)
    
    firstPageEnd = firstPage.Start
    
    Set rng = doc.Range(Start:=firstPageEnd, End:=doc.Content.End)
    rng.Select
    rng.Find.ClearFormatting
    rng.Find.Replacement.ClearFormatting

MsgBox "�������� ��������." & vbCrLf & "�� - �������������� ������ ��������; " & vbCrLf & "��� - ������������� ������ ��������"
With rng.Find
    .ClearFormatting
    .Text = "       "  ' �������� ������ �������� ������
    .Replacement.ClearFormatting
    .Replacement.Text = " "  ' ����� ������ �������� ������
    .Wrap = wdFindAsk
    .Forward = True

    rng.Find.Execute Replace:=wdReplaceAll
End With
With rng.Find
    .ClearFormatting
    .Text = "      "  ' �������� ������ �������� ������
    .Replacement.ClearFormatting
    .Replacement.Text = " "  ' ����� ������ �������� ������
    .Wrap = wdFindAsk
    .Forward = True

    rng.Find.Execute Replace:=wdReplaceAll
End With
    With rng.Find
        .ClearFormatting
        .Text = "    "  ' �������� ������ �������� ������
        .Replacement.ClearFormatting
        .Replacement.Text = " "  ' ����� ������ �������� ������
        .Wrap = wdFindAsk
        .Forward = True

        ' ��������� ������
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    With rng.Find
        .ClearFormatting
        .Text = "   "  ' �������� ������ �������� ������
        .Replacement.ClearFormatting
        .Replacement.Text = " "  ' ����� ������ �������� ������
        .Wrap = wdFindAsk
        .Forward = True

        ' ��������� ������
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    With rng.Find
        .ClearFormatting
        .Text = "  "  ' �������� ������ �������� ������
        .Replacement.ClearFormatting
        .Replacement.Text = " "  ' ����� ������ �������� ������
        .Wrap = wdFindAsk
        .Forward = True

        ' ��������� ������
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
    
     
    With rng.Find
        .ClearFormatting
        .Text = " ^p"   ' ���� ������ ����� ����� �������
        
        .Replacement.ClearFormatting
        .Replacement.Text = "^p" ' �������� �� ����� �����
        .Wrap = wdFindAsk
        .Forward = True
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
     With rng.Find
        .ClearFormatting
        .Text = "^p "    ' ���� ������ ����� ����� �������
        .Replacement.ClearFormatting
        .Replacement.Text = "^p" ' �������� �� ����� �����
        .Wrap = wdFindAsk
        .Forward = True
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
      With rng.Find
        .ClearFormatting
        .Text = vbFormFeed & " "
        .Replacement.ClearFormatting
        .Replacement.Text = vbFormFeed
        .Wrap = wdFindAsk
        .Forward = True

        rng.Find.Execute Replace:=wdReplaceAll
    End With
      With rng.Find
        .ClearFormatting
        .Text = Chr(9) & "^p"
        
        .Replacement.ClearFormatting
        .Replacement.Text = "^p" ' �������� �� ����� �����
        .Wrap = wdFindAsk
        .Forward = True
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
End Sub

Sub ReplacePages3()
    Dim doc As Document
    Set doc = ActiveDocument
  
    With selection.Find
        .ClearFormatting
        .Text = vbFormFeed  ' �������� ������ �������� ������
        .Replacement.ClearFormatting
        .Replacement.Text = vbFormFeed  ' ����� ������ �������� ������
        .Wrap = wdFindContinue
        .Forward = True

        ' ��������� ������
        selection.Find.Execute Replace:=wdReplaceAll
    End With
   

   ' MsgBox "������ ���������!"
End Sub

Sub DelExcessCarrets4()
Dim doc As Document
    Dim rng As Range
    Dim firstPageEnd As Long
    Dim firstPage As Range
   
    Set doc = ActiveDocument
    
    Set firstPage = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=2)
    
    firstPageEnd = firstPage.Start
    
    Set rng = doc.Range(Start:=firstPageEnd, End:=doc.Content.End)
    rng.Select
    rng.Find.ClearFormatting
    rng.Find.Replacement.ClearFormatting

MsgBox "�������� �������." & vbCrLf & "�� - �������������� ������ ��������; " & vbCrLf & "��� - ������������� ������ ��������"

    With rng.Find
        .Text = "^p^p^p^p^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    rng.Find.Execute Replace:=wdReplaceAll
    With rng.Find
        .Text = "^p^p^p^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    rng.Find.Execute Replace:=wdReplaceAll
    With rng.Find
        .Text = "^p^p^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    rng.Find.Execute Replace:=wdReplaceAll
    With rng.Find
        .Text = "^p^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    rng.Find.Execute Replace:=wdReplaceAll
    With rng.Find
        .Text = "^p^p"
        .Replacement.Text = "^p"
        .Forward = True
        .Wrap = wdFindAsk
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    rng.Find.Execute Replace:=wdReplaceAll
End Sub


Sub ReplaceHeaders5()
    Dim doc As Document
    Set doc = ActiveDocument
    
    With selection.Find
        .ClearFormatting
        .Text = "��������^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "��������^p"
        .Wrap = wdFindContinue
        .Forward = True
    
        selection.Find.Execute Replace:=wdReplaceAll
           
    End With
    With selection.Find
        .ClearFormatting
        .Text = "��������^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "��������^p"
        .Wrap = wdFindContinue
        .Forward = True

        
       selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    
    
    With selection.Find
        .ClearFormatting
        .Text = "����������^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "����������^p"
        .Wrap = wdFindContinue
        .Forward = True

        
        selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    With selection.Find
        .ClearFormatting
        .Text = "����������^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "����������^p"
        .Wrap = wdFindContinue
        .Forward = True

        
        selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    
    
     With selection.Find
        .ClearFormatting
        .Text = "����������^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "����������^p"
        .Wrap = wdFindContinue
        .Forward = True

         
        selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    With selection.Find
        .ClearFormatting
        .Text = "����������^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "����������^p"
        .Wrap = wdFindContinue
        .Forward = True

         
        selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    
  
    With selection.Find
        .ClearFormatting
        .Text = "������ �������������� ����������^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "������ �������������� ����������^p"
        .Wrap = wdFindContinue
        .Forward = True

         
        selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    With selection.Find
        .ClearFormatting
        .Text = "������ �������������� ����������^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "������ �������������� ����������^p"
        .Wrap = wdFindContinue
        .Forward = True

         
        selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    
     With selection.Find
        .ClearFormatting
        .Text = "������ ����������^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "������ �������������� ����������^p"
        .Wrap = wdFindContinue
        .Forward = True
        selection.Find.Execute Replace:=wdReplaceAll
    End With
    
     With selection.Find
        .ClearFormatting
        .Text = "C����� ����������^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "������ �������������� ����������^p"
        .Wrap = wdFindContinue
        .Forward = True
        selection.Find.Execute Replace:=wdReplaceAll
    End With
    
     With selection.Find
        .ClearFormatting
        .Text = "������ �������������� ����������^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "������ �������������� ����������^p"
        .Wrap = wdFindContinue
        .Forward = True
        selection.Find.Execute Replace:=wdReplaceAll
    End With
    
    With selection.Find
        .ClearFormatting
        .Text = "������ �������������� ����������^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "������ �������������� ����������^p"
        .Wrap = wdFindContinue
        .Forward = True
        selection.Find.Execute Replace:=wdReplaceAll
    End With
    
End Sub
Sub ReplaceBulletsWithDash6()
    Dim para As Paragraph
    Dim listFormat As listFormat
    With ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = ChrW(61485)
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleBullet
        .NumberPosition = CentimetersToPoints(1.25)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "Symbol"
        End With
        .LinkedStyle = ""
    End With
        
    ' �������� �� ���� ������� � ���������
    For Each para In ActiveDocument.Paragraphs
        Set listFormat = para.Range.listFormat
        
        ' ���������, ���� �� � ������ �������������� ������
      
        If listFormat.listType = wdListBullet Then
            ' �������� ������ �� ������ '-'
            para.Range.listFormat.RemoveNumbers
            para.Range.listFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdBulletGallery).ListTemplates(1), ContinuePreviousList:= _
        True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
        End If
    Next para
End Sub



Sub CheckListMarkers7()
    Dim para As Paragraph
    Dim errorMessage As String
    Dim bulletCharacter As String
    Dim firstChar As String
    Const ERROR_MSG As String = "� ������� ������" ' Error message constant
    Dim lastChar As String
    Dim nextlastChar As String
    ' Check if any lists exist in the document
    If ActiveDocument.Lists.Count = 0 Then
        
        Exit Sub
    End If

    ' Loop through each paragraph in the document
    For Each para In ActiveDocument.Paragraphs
        ' Check if the paragraph is part of a list
        If para.Range.listFormat.listType <> wdListNoNumbering Then
           ' bulletCharacter = para.Range.Characters(1).Text
            
            
            'lastChar = Trim(Right(para.Range.Text, 2))
          
            If Trim(Left(para.Range.Text, 1)) Like "[A-Z]" Or Trim(Left(para.Range.Text, 1)) Like "[�-�]" Then
                errorMessage = ERROR_MSG & ": ��������� ����� � ������."
                para.Range.Font.Color = wdColorRed
                para.Range.Comments.Add Range:=para.Range, Text:="������ �� ������ ���������� � ������� �����"
            End If
          
           If Not para.Next Is Nothing Then
                If (Right(para.Range.Text, 2) <> ";" & Chr(13) And para.Next.Range.listFormat.listType <> wdListNoNumbering) _
                Or (Right(para.Range.Text, 2) <> "." & Chr(13) And para.Next.Range.listFormat.listType = wdListNoNumbering) _
                Then
        
                        errorMessage = errorMessage & vbNewLine & ERROR_MSG & ": ���������� �� ������ ������������� ������."
                        para.Range.Font.Color = wdColorRed
                        para.Range.Comments.Add Range:=para.Range, Text:="�������� ������ �� ������ ������������� �� ����� ���������� ����� ���������� (��������� ������� ������ ������������� ������)"
                End If
           End If
        End If
    Next para

    ' Display an error message if any issues were found
    If errorMessage <> "" Then
        MsgBox errorMessage
    End If
End Sub

Sub SetAllIndentsToZero8()
    Dim doc As Document
    Dim para As Paragraph
    Set doc = ActiveDocument
    
    For Each para In doc.Paragraphs
        With para.Range.ParagraphFormat
            .FirstLineIndent = 0
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
    Next para
End Sub


Sub Numbering9()
  selection.Sections(1).Footers(1).PageNumbers.Add PageNumberAlignment:=1, firstPage:=False
  
End Sub


Sub Fields10()

selection.WholeStory
    
    With ActiveDocument.PageSetup
        
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(3)
        .RightMargin = CentimetersToPoints(1)
        
    End With

End Sub
Private Sub ChangeFont11()
  
    ActiveDocument.Content.Font.Name = "Times New Roman"
    ActiveDocument.Content.Font.Size = 14
    
End Sub


Sub SetLineSpacingToOnePointFive12()
    Dim doc As Document
    Dim rng As Range
    Dim firstPageEnd As Long
    Dim firstPage As Range
   
    Set doc = ActiveDocument
    
    Set firstPage = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=2)
    
    firstPageEnd = firstPage.Start
    
    Set rng = doc.Range(Start:=firstPageEnd, End:=doc.Content.End)
    
    rng.Select

    With rng.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpace1pt5
    End With
End Sub

Sub ParagraphIndent13()
    Dim para As Paragraph
    Dim doc As Document
    Dim rng As Range
    Dim firstPageEnd As Long
    Dim firstPage As Range

    Set doc = ActiveDocument
    
    ' ������� �� ������ ��������
    Set firstPage = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=2)
    firstPageEnd = firstPage.Start
    
    ' ���������� �������� �� ����� ������ �������� �� ����� ���������
    Set rng = doc.Range(Start:=firstPageEnd, End:=doc.Content.End)

    ' �������� �� ���� ���������� � ����������� ���������
    For Each para In rng.Paragraphs
        ' ���������, �������� �� �������� ��������� (�� �������� ��������)
        If para.Range.InlineShapes.Count = 0 And para.Range.ShapeRange.Count = 0 And para.Range.Information(wdWithInTable) = False Then
            ' ��������� ������ ������ � ��������� ����������
            para.Range.ParagraphFormat.FirstLineIndent = CentimetersToPoints(1.25)
        End If
    Next para
End Sub



Sub ResizeTablesToWindowWidth14()
   
    Dim tbl As Table
    Dim pageWidth As Single
    Dim tblWidth As Single
    Dim doc As Document
    
    Set doc = ActiveDocument
    pageWidth = doc.PageSetup.pageWidth
    pageHeight = doc.PageSetup.pageHeight
    
    For Each tbl In doc.Tables
        tbl.PreferredWidthType = wdPreferredWidthPercent
        tbl.PreferredWidth = 100
        tbl.Rows.Alignment = wdAlignRowCenter
    Next tbl
   
    
End Sub


Sub Reddots15()
    Dim para As Paragraph
    Dim messageShown As Boolean
    
    For Each para In ActiveDocument.Paragraphs
        
        If para.Range.listFormat.listType <> wdListNoNumbering Then
            
            If (Not (para.Previous Is Nothing)) And para.Previous.Range.listFormat.listType = wdListNoNumbering Then
          
                If Right(para.Previous.Range.Text, 2) <> ":" & Chr(13) Then
                    para.Previous.Range.Font.Color = wdColorRed ' ??????? ????
                    para.Previous.Range.Comments.Add Range:=para.Range, Text:="����� ������� ������ ������ ������ ���������"
                End If
            End If
        End If
        If Left(Trim(para.Range.Text), 1) = "-" Or _
        (IsNumeric(Left(Trim(para.Range.Text), 1)) And (Mid(para.Range.Text, 2, 1) = ".")) Or _
        (IsNumeric(Left(Trim(para.Range.Text), 1)) And (Mid(para.Range.Text, 2, 1) = ")")) Then
            ' �������� �������� �������
            para.Range.Font.Color = wdColorRed ' ������� ����
            para.Range.Comments.Add Range:=para.Range, Text:="�������� �� ������ ������� ����� ������"
            ' ���������, ���� �� ��������� ��� ��������
            If Not messageShown Then
                MsgBox ""
                messageShown = True ' ������������� ����, ��� ��������� ��������
            End If
        End If
    Next para
End Sub
Sub ReplaceWord16()

    Dim rng As Range
    Dim doc As Document

    Set doc = ActiveDocument

    ' ������������� �������� �� ���� ��������
    Set rng = doc.Content
    
    With rng.Find
        .Text = "���."
        
        .Replacement.Text = "�������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' ��������� ������ � ���������
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
   
     With rng.Find
        .Text = "(���."
        
        .Replacement.Text = "(�������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' ��������� ������ � ���������
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
    With rng.Find
        .Text = "[���."
        
        .Replacement.Text = "[�������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' ��������� ������ � ���������
        rng.Find.Execute Replace:=wdReplaceAll
    End With
     With rng.Find
      With rng.Find
        .Text = "^p��� "
        
        .Replacement.Text = "^p������� "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' ��������� ������ � ���������
        rng.Find.Execute Replace:=wdReplaceAll
    End With
        .Text = "����."
        
        .Replacement.Text = "�������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' ��������� ������ � ���������
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
     With rng.Find
        .Text = "(����."
        
        .Replacement.Text = "(�������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' ��������� ������ � ���������
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
    With rng.Find
        .Text = "[����."
        
        .Replacement.Text = "[�������"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' ��������� ������ � ���������
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    With rng.Find
        .Text = "^p���� "
        
        .Replacement.Text = "^p������� "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' ��������� ������ � ���������
        rng.Find.Execute Replace:=wdReplaceAll
    End With
End Sub

Sub AlignJustify17()
    Dim doc As Document
    Dim rng As Range
    Dim firstPageEnd As Long
    Dim firstPage As Range
   
    Set doc = ActiveDocument
    
    Set firstPage = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=2)
    
    firstPageEnd = firstPage.Start
    
    Set rng = doc.Range(Start:=firstPageEnd, End:=doc.Content.End)
    
    rng.Select
    With rng.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .Alignment = wdAlignParagraphJustify
    End With

    ' �������� �� ������� ������ � ���������
    'For Each para In ActiveDocument.Paragraphs
     '   If para.Range.Information(wdActiveEndAdjustedPageNumber) > 1 Then
      '   para.Alignment = wdAlignParagraphJustify ' ������������� ������������ �� ������
       ' End If
    'Next para --
     
    'MsgBox "��� ������ ��������� �� ������", vbInformation, "���������"
End Sub

Sub CenterAlignIfImageOrTable18()
    Dim tbl As Table
    Dim para As Paragraph
    Dim foundImage As Boolean
    Dim messageShown As Boolean
    Dim messageShownT As Boolean
    foundImage = False
    messageShown = False
    messageShownT = False
    ' �������� ����� ��� ��������� ���������
    For Each para In ActiveDocument.Paragraphs
        ' ���������, ���� �� � ��������� �����������
         If Not para Is Nothing Then
            
            If para.Range.InlineShapes.Count > 0 Then
                foundImage = True ' ������������� ����, ���� � ��������� ���� �����������
            ElseIf foundImage Then ' ���� ���������� �������� �������� �����������
         
                If (Left((para.Range.Text), 8) = "������� " And IsNumeric(Mid(para.Range.Text, 9, 1)) _
                And Mid(para.Range.Text, 10, 1) = "." And IsNumeric(Mid(para.Range.Text, 11, 1)) _
                And Mid(para.Range.Text, 12, 3) = Chr(32) & Chr(150) & Chr(32) _
                And (Mid(para.Range.Text, 15, 1) Like "[�-�]" Or Mid(para.Range.Text, 15, 1) Like "[A-Z]") _
                And (Right((para.Range.Text), 2) <> ";" & Chr(13) And Right((para.Range.Text), 2) <> "." & Chr(13))) _
                Then
                
                ' ����������� ������� �������� �� ������
                    
                    para.FirstLineIndent = CentimetersToPoints(0)
                    para.Alignment = wdAlignParagraphCenter
                    para.Range.InsertAfter vbCrLf
                    para.Previous.Range.InsertBefore vbCrLf
             Else
                ' ������������� ���� ������ � �������
                    para.Range.Font.Color = wdColorRed
                    para.Range.Comments.Add Range:=para.Range, Text:="����������� �������� �������. ������: ������� 1.1 - �������� (������������ ������� ����)"
                    If Not messageShown Then
                       ' MsgBox "��������� ������� ��������� �����������", vbExclamation
                        messageShown = True
                    End If
                End If
            foundImage = False ' ���������� ���� ����� ���������
        End If
        
        If para.Range.Tables.Count > 0 Then
        If Not para.Previous Is Nothing Then
        If para.Previous.Range.Tables.Count = 0 Then
                If (Left((para.Previous.Range.Text), 8) = "������� " And IsNumeric(Mid(para.Previous.Range.Text, 9, 1)) _
                And Mid(para.Previous.Range.Text, 10, 1) = "." And IsNumeric(Mid(para.Previous.Range.Text, 11, 1)) _
                And Mid(para.Previous.Range.Text, 12, 3) = Chr(32) & Chr(150) & Chr(32) _
                And (Mid(para.Previous.Range.Text, 15, 1) Like "[�-�]" Or Mid(para.Previous.Range.Text, 15, 1) Like "[A-Z]") _
                And (Right((para.Previous.Range.Text), 2) <> ";" & Chr(13) And Right((para.Previous.Range.Text), 2) <> "." & Chr(13))) _
                Then
                
                    para.Previous.FirstLineIndent = CentimetersToPoints(0)
                    para.Previous.Alignment = wdAlignParagraphLeft
                    para.Previous
                    para.Previous.Range.Select
                    selection.Collapse wdCollapseStart
                    'selection.Move Unit:=wdCharacter, Count:=-1
                    selection.TypeParagraph
                    Set tbl = para.Range.Tables(1)
                    tbl.Select
                    If Not Left((selection.Range.Next.Text), 1) = Chr(13) Then
                        selection.Collapse wdCollapseEnd
                    'selection.Move Unit:=wdCharacter, Count:=1
                        selection.TypeParagraph
                    End If
                Else
                ' ������������� ���� ������ � �������
                    para.Previous.Range.Font.Color = wdColorRed
                    para.Previous.Range.Comments.Add Range:=para.Range, Text:="����������� �������� �������. ������: ������� 1.1 - �������� (������������ ������� ����)"
                    
                    If Not messageShownT Then
                        MsgBox "��������� ������� ��������� �����������", vbExclamation
                        messageShownT = True
                    End If
                End If
            End If
        End If
        End If
        End If
        
    Next para

End Sub
Sub CenterAllImages19()
    Dim shape As shape
    Dim inlineShape As inlineShape
    
    'For Each shape In ActiveDocument.Shapes
     '   shape.Left = (ActiveDocument.PageSetup.pageWidth - shape.Width) / 2
    'Next shape

    For Each inlineShape In ActiveDocument.InlineShapes
        inlineShape.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Next inlineShape
End Sub
Sub FormatTables20()

    Dim tbl As Table
    Dim cel As Cell

    ' ���������� ��� ������� � ���������
    For Each tbl In ActiveDocument.Tables
        ' ���������� ��� ������ � ������� �������
        For Each cel In tbl.Range.Cells
            ' ������������� ��������� ����������� ��������
            cel.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            
            ' ������������ ������ �� ������ ����
            cel.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        Next cel
    Next tbl

    MsgBox "��� ������� ���������������!"
   ' MsgBox "���� ��������"
End Sub




Sub CheckHeaders21()
    Dim doc As Document
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If Left(para.Range.Text, 8) = "��������" Or Left(para.Range.Text, 9) = vbFormFeed & "��������" _
        Or Left((para.Range.Text), 10) = "����������" Or Left((para.Range.Text), 11) = vbFormFeed & "����������" _
            Or Left((para.Range.Text), 10) = "����������" Or Left((para.Range.Text), 11) = vbFormFeed & "����������" _
            Or (Left((para.Range.Text), 32) = "������ �������������� ����������") Or (Left((para.Range.Text), 33) = vbFormFeed & "������ �������������� ����������") Then
                para.Alignment = wdAlignParagraphCenter
                para.FirstLineIndent = CentimetersToPoints(0)
                
                'para.Next.Alignment = wdAlignParagraphJustify
        End If
    Next para
End Sub

Sub PageHeaders22()
    Dim doc As Document
    Dim para As Paragraph
    Set doc = ActiveDocument

    ' ����������� �� ���������� � ���������
    For Each para In doc.Paragraphs
        ' ���������, ���������� �� �������� � ������ ����
        If (Left(para.Range.Text, 9) = "��������" & Chr(13) _
            Or Left(para.Range.Text, 11) = "����������" & Chr(13) _
            Or Left(para.Range.Text, 11) = "����������" & Chr(13) _
            Or Left(para.Range.Text, 33) = "������ �������������� ����������" & Chr(13)) Then
            
            ' ���������, ���� �� ���������� ��������
            If Not para.Previous Is Nothing Then
                ' ���������, �� ������������� �� ���������� �������� �� ������ ��������
                If Right(para.Previous.Range.Text, 2) <> vbFormFeed & Chr(13) And Left(para.Range.Text, 1) <> vbFormFeed Then
                    ' ��������� ������ �������� ����� ������� ����������
                    
                    para.Range.InsertBefore vbFormFeed & "^p"
                    'para.Range.InsertParagraphAfter
                    
                End If
           
            End If
        End If
    Next para
End Sub






Sub ReplaceBulletsNumbers23()
Dim doc As Document
    Dim rng As Range
    Dim found As Boolean

    Set doc = ActiveDocument
    Set rng = doc.Content
    
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(1.25)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = wdUndefined
            .Italic = wdUndefined
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdUndefined
            .Size = wdUndefined
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = ""
        End With
        .LinkedStyle = ""
    End With
   For Each para In ActiveDocument.Paragraphs
        
        
        ' ��������, �������� �� ����� ������������
        If para.Range.listFormat.listType = wdListSimpleNumbering Then
            ' ������� ��������� � ��������� ������������� ������
            para.Range.listFormat.RemoveNumbers
            
        para.Range.listFormat.ApplyListTemplateWithLevel ListTemplate:= _
        ListGalleries(wdNumberGallery).ListTemplates(1), ContinuePreviousList:= _
        True, ApplyTo:=wdListApplyToWholeList, DefaultListBehavior:= _
        wdWord10ListBehavior
        End If
    Next para
End Sub
Sub RedName24()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If Left(para.Range.Text, 2) Like "[A-Z]" & "." Or Left(para.Range.Text, 2) Like "[�-�]" & "." Then
        
            para.Range.Font.Color = wdColorRed
            para.Range.Comments.Add Range:=para.Range, Text:="��� ������ ���������� � �������"
            'With para.Range.Comments.Add(Range:=para.Range, Text:="��� ������ ���������� � �������")
                
            'End With
            
        End If
    Next para
End Sub
Sub Formula25()
  Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        ' ���������, �������� �� ����� ������ "+"
        If InStr(para.Range.Text, "=") > 0 Then
            ' �������� ���� ����� ������� ������
            para.Range.Font.Color = wdColorRed
            para.Range.Comments.Add Range:=para.Range, Text:="�������� �� ������ �� ����� �������" & Chr(13) & "������:A+B=C              (1.1)" & Chr(13) & "(������� ������������� ����� �������->���������,������� ������������� �� ������ ����, ����� - �� �������)"
            
        End If
    Next para
End Sub


Sub BlackLiterature26()
    Dim rng As Range
    

    ' ������������� �������� �� ���� ��������
    Set rng = ActiveDocument.Content

    With rng.Find
        .Text = "������ �������������� ����������"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWholeWord = True
        
        If .Execute Then ' ���� ������� ����� "�����"
            ' ���������� �������� � ����� ���������� �����
            rng.Collapse wdCollapseStart
            
            ' ������������� ����� �������� �� ����� ����� "�����" �� ����� ���������
            Dim rngToHighlight As Range
            Set rngToHighlight = ActiveDocument.Content
            rngToHighlight.Start = rng.Start
            rngToHighlight.End = ActiveDocument.Content.End
            
            ' �������� �����
            rngToHighlight.Font.Color = -587137025
        
        End If
    End With
    MsgBox "���� ��������"
End Sub
