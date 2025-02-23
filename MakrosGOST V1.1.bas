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

MsgBox "Проверка пробелов." & vbCrLf & "ДА - редактирование первой страницы; " & vbCrLf & "НЕТ - игнорирование первой страницы"
With rng.Find
    .ClearFormatting
    .Text = "       "  ' Исходный символ переноса строки
    .Replacement.ClearFormatting
    .Replacement.Text = " "  ' Новый символ переноса строки
    .Wrap = wdFindAsk
    .Forward = True

    rng.Find.Execute Replace:=wdReplaceAll
End With
With rng.Find
    .ClearFormatting
    .Text = "      "  ' Исходный символ переноса строки
    .Replacement.ClearFormatting
    .Replacement.Text = " "  ' Новый символ переноса строки
    .Wrap = wdFindAsk
    .Forward = True

    rng.Find.Execute Replace:=wdReplaceAll
End With
    With rng.Find
        .ClearFormatting
        .Text = "    "  ' Исходный символ переноса строки
        .Replacement.ClearFormatting
        .Replacement.Text = " "  ' Новый символ переноса строки
        .Wrap = wdFindAsk
        .Forward = True

        ' Выполняем замену
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    With rng.Find
        .ClearFormatting
        .Text = "   "  ' Исходный символ переноса строки
        .Replacement.ClearFormatting
        .Replacement.Text = " "  ' Новый символ переноса строки
        .Wrap = wdFindAsk
        .Forward = True

        ' Выполняем замену
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    With rng.Find
        .ClearFormatting
        .Text = "  "  ' Исходный символ переноса строки
        .Replacement.ClearFormatting
        .Replacement.Text = " "  ' Новый символ переноса строки
        .Wrap = wdFindAsk
        .Forward = True

        ' Выполняем замену
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
    
     
    With rng.Find
        .ClearFormatting
        .Text = " ^p"   ' Ищем пробел перед новой строкой
        
        .Replacement.ClearFormatting
        .Replacement.Text = "^p" ' Заменяем на новый абзац
        .Wrap = wdFindAsk
        .Forward = True
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
     With rng.Find
        .ClearFormatting
        .Text = "^p "    ' Ищем пробел перед новой строкой
        .Replacement.ClearFormatting
        .Replacement.Text = "^p" ' Заменяем на новый абзац
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
        .Replacement.Text = "^p" ' Заменяем на новый абзац
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
        .Text = vbFormFeed  ' Исходный символ переноса строки
        .Replacement.ClearFormatting
        .Replacement.Text = vbFormFeed  ' Новый символ переноса строки
        .Wrap = wdFindContinue
        .Forward = True

        ' Выполняем замену
        selection.Find.Execute Replace:=wdReplaceAll
    End With
   

   ' MsgBox "Замена завершена!"
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

MsgBox "Проверка абзацев." & vbCrLf & "ДА - редактирование первой страницы; " & vbCrLf & "НЕТ - игнорирование первой страницы"

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
        .Text = "Введение^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "ВВЕДЕНИЕ^p"
        .Wrap = wdFindContinue
        .Forward = True
    
        selection.Find.Execute Replace:=wdReplaceAll
           
    End With
    With selection.Find
        .ClearFormatting
        .Text = "введение^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "ВВЕДЕНИЕ^p"
        .Wrap = wdFindContinue
        .Forward = True

        
       selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    
    
    With selection.Find
        .ClearFormatting
        .Text = "Содержание^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "СОДЕРЖАНИЕ^p"
        .Wrap = wdFindContinue
        .Forward = True

        
        selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    With selection.Find
        .ClearFormatting
        .Text = "содержание^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "СОДЕРЖАНИЕ^p"
        .Wrap = wdFindContinue
        .Forward = True

        
        selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    
    
     With selection.Find
        .ClearFormatting
        .Text = "Заключение^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "ЗАКЛЮЧЕНИЕ^p"
        .Wrap = wdFindContinue
        .Forward = True

         
        selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    With selection.Find
        .ClearFormatting
        .Text = "заключение^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "ЗАКЛЮЧЕНИЕ^p"
        .Wrap = wdFindContinue
        .Forward = True

         
        selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    
  
    With selection.Find
        .ClearFormatting
        .Text = "Список использованных источников^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ^p"
        .Wrap = wdFindContinue
        .Forward = True

         
        selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    With selection.Find
        .ClearFormatting
        .Text = "список использованных источников^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ^p"
        .Wrap = wdFindContinue
        .Forward = True

         
        selection.Find.Execute Replace:=wdReplaceAll
            
    End With
    
     With selection.Find
        .ClearFormatting
        .Text = "список литературы^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ^p"
        .Wrap = wdFindContinue
        .Forward = True
        selection.Find.Execute Replace:=wdReplaceAll
    End With
    
     With selection.Find
        .ClearFormatting
        .Text = "Cписок литературы^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ^p"
        .Wrap = wdFindContinue
        .Forward = True
        selection.Find.Execute Replace:=wdReplaceAll
    End With
    
     With selection.Find
        .ClearFormatting
        .Text = "список использованной литературы^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ^p"
        .Wrap = wdFindContinue
        .Forward = True
        selection.Find.Execute Replace:=wdReplaceAll
    End With
    
    With selection.Find
        .ClearFormatting
        .Text = "Список использованной литературы^p"
        .Replacement.ClearFormatting
        .Replacement.Text = "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ^p"
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
        
    ' Проходим по всем абзацам в документе
    For Each para In ActiveDocument.Paragraphs
        Set listFormat = para.Range.listFormat
        
        ' Проверяем, есть ли у абзаца форматирование списка
      
        If listFormat.listType = wdListBullet Then
            ' Заменяем маркер на символ '-'
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
    Const ERROR_MSG As String = "В списках ошибка" ' Error message constant
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
          
            If Trim(Left(para.Range.Text, 1)) Like "[A-Z]" Or Trim(Left(para.Range.Text, 1)) Like "[А-Я]" Then
                errorMessage = ERROR_MSG & ": Заглавная буква в списке."
                para.Range.Font.Color = wdColorRed
                para.Range.Comments.Add Range:=para.Range, Text:="Список не должен начинаться с большой буквы"
            End If
          
           If Not para.Next Is Nothing Then
                If (Right(para.Range.Text, 2) <> ";" & Chr(13) And para.Next.Range.listFormat.listType <> wdListNoNumbering) _
                Or (Right(para.Range.Text, 2) <> "." & Chr(13) And para.Next.Range.listFormat.listType = wdListNoNumbering) _
                Then
        
                        errorMessage = errorMessage & vbNewLine & ERROR_MSG & ": Примечание не должно заканчиваться точкой."
                        para.Range.Font.Color = wdColorRed
                        para.Range.Comments.Add Range:=para.Range, Text:="Элементы списка не должен заканчиваться на знаки препинания кроме последнего (последний элемент должен заканчиваться точкой)"
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
    
    ' Переход ко второй странице
    Set firstPage = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=2)
    firstPageEnd = firstPage.Start
    
    ' Определяем диапазон от конца первой страницы до конца документа
    Set rng = doc.Range(Start:=firstPageEnd, End:=doc.Content.End)

    ' Проходим по всем параграфам в определённом диапазоне
    For Each para In rng.Paragraphs
        ' Проверяем, является ли параграф текстовым (не содержит объектов)
        If para.Range.InlineShapes.Count = 0 And para.Range.ShapeRange.Count = 0 And para.Range.Information(wdWithInTable) = False Then
            ' Применяем отступ только к текстовым параграфам
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
                    para.Previous.Range.Comments.Add Range:=para.Range, Text:="Перед началом списка должно стоять двоеточие"
                End If
            End If
        End If
        If Left(Trim(para.Range.Text), 1) = "-" Or _
        (IsNumeric(Left(Trim(para.Range.Text), 1)) And (Mid(para.Range.Text, 2, 1) = ".")) Or _
        (IsNumeric(Left(Trim(para.Range.Text), 1)) And (Mid(para.Range.Text, 2, 1) = ")")) Then
            ' Выделяем параграф красным
            para.Range.Font.Color = wdColorRed ' Красный цвет
            para.Range.Comments.Add Range:=para.Range, Text:="Возможно вы хотели создать здесь список"
            ' Проверяем, было ли сообщение уже показано
            If Not messageShown Then
                MsgBox ""
                messageShown = True ' Устанавливаем флаг, что сообщение показано
            End If
        End If
    Next para
End Sub
Sub ReplaceWord16()

    Dim rng As Range
    Dim doc As Document

    Set doc = ActiveDocument

    ' Устанавливаем диапазон на весь документ
    Set rng = doc.Content
    
    With rng.Find
        .Text = "рис."
        
        .Replacement.Text = "Рисунок"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' Выполняем замену в документе
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
   
     With rng.Find
        .Text = "(рис."
        
        .Replacement.Text = "(Рисунок"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' Выполняем замену в документе
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
    With rng.Find
        .Text = "[рис."
        
        .Replacement.Text = "[Рисунок"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' Выполняем замену в документе
        rng.Find.Execute Replace:=wdReplaceAll
    End With
     With rng.Find
      With rng.Find
        .Text = "^pрис "
        
        .Replacement.Text = "^pРисунок "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' Выполняем замену в документе
        rng.Find.Execute Replace:=wdReplaceAll
    End With
        .Text = "табл."
        
        .Replacement.Text = "Таблица"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' Выполняем замену в документе
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
     With rng.Find
        .Text = "(табл."
        
        .Replacement.Text = "(Таблица"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' Выполняем замену в документе
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    
    With rng.Find
        .Text = "[табл."
        
        .Replacement.Text = "[Таблица"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' Выполняем замену в документе
        rng.Find.Execute Replace:=wdReplaceAll
    End With
    With rng.Find
        .Text = "^pтабл "
        
        .Replacement.Text = "^pТаблица "
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        
        ' Выполняем замену в документе
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

    ' Проходим по каждому абзацу в документе
    'For Each para In ActiveDocument.Paragraphs
     '   If para.Range.Information(wdActiveEndAdjustedPageNumber) > 1 Then
      '   para.Alignment = wdAlignParagraphJustify ' Устанавливаем выравнивание по ширине
       ' End If
    'Next para --
     
    'MsgBox "Все абзацы выровнены по ширине", vbInformation, "Завершено"
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
    ' Проходим через все параграфы документа
    For Each para In ActiveDocument.Paragraphs
        ' Проверяем, есть ли в параграфе изображение
         If Not para Is Nothing Then
            
            If para.Range.InlineShapes.Count > 0 Then
                foundImage = True ' Устанавливаем флаг, если в параграфе есть изображение
            ElseIf foundImage Then ' Если предыдущий параграф содержал изображение
         
                If (Left((para.Range.Text), 8) = "Рисунок " And IsNumeric(Mid(para.Range.Text, 9, 1)) _
                And Mid(para.Range.Text, 10, 1) = "." And IsNumeric(Mid(para.Range.Text, 11, 1)) _
                And Mid(para.Range.Text, 12, 3) = Chr(32) & Chr(150) & Chr(32) _
                And (Mid(para.Range.Text, 15, 1) Like "[А-Я]" Or Mid(para.Range.Text, 15, 1) Like "[A-Z]") _
                And (Right((para.Range.Text), 2) <> ";" & Chr(13) And Right((para.Range.Text), 2) <> "." & Chr(13))) _
                Then
                
                ' Выравниваем текущий параграф по центру
                    
                    para.FirstLineIndent = CentimetersToPoints(0)
                    para.Alignment = wdAlignParagraphCenter
                    para.Range.InsertAfter vbCrLf
                    para.Previous.Range.InsertBefore vbCrLf
             Else
                ' Устанавливаем цвет шрифта в красный
                    para.Range.Font.Color = wdColorRed
                    para.Range.Comments.Add Range:=para.Range, Text:="Неправильно подписан рисунок. Пример: Рисунок 1.1 - Название (используется среднее тире)"
                    If Not messageShown Then
                       ' MsgBox "Некоторые рисунки подписаны неправильно", vbExclamation
                        messageShown = True
                    End If
                End If
            foundImage = False ' Сбрасываем флаг после обработки
        End If
        
        If para.Range.Tables.Count > 0 Then
        If Not para.Previous Is Nothing Then
        If para.Previous.Range.Tables.Count = 0 Then
                If (Left((para.Previous.Range.Text), 8) = "Таблица " And IsNumeric(Mid(para.Previous.Range.Text, 9, 1)) _
                And Mid(para.Previous.Range.Text, 10, 1) = "." And IsNumeric(Mid(para.Previous.Range.Text, 11, 1)) _
                And Mid(para.Previous.Range.Text, 12, 3) = Chr(32) & Chr(150) & Chr(32) _
                And (Mid(para.Previous.Range.Text, 15, 1) Like "[А-Я]" Or Mid(para.Previous.Range.Text, 15, 1) Like "[A-Z]") _
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
                ' Устанавливаем цвет шрифта в красный
                    para.Previous.Range.Font.Color = wdColorRed
                    para.Previous.Range.Comments.Add Range:=para.Range, Text:="Неправильно подписан рисунок. Пример: Таблица 1.1 - Название (используется среднее тире)"
                    
                    If Not messageShownT Then
                        MsgBox "Некоторые таблицы подписаны неправильно", vbExclamation
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

    ' Перебираем все таблицы в документе
    For Each tbl In ActiveDocument.Tables
        ' Перебираем все ячейки в текущей таблице
        For Each cel In tbl.Range.Cells
            ' Устанавливаем одинарный межстрочный интервал
            cel.Range.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            
            ' Выравнивание текста по левому краю
            cel.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        Next cel
    Next tbl

    MsgBox "Все таблицы отформатированы!"
   ' MsgBox "ГОСТ соблюден"
End Sub




Sub CheckHeaders21()
    Dim doc As Document
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If Left(para.Range.Text, 8) = "ВВЕДЕНИЕ" Or Left(para.Range.Text, 9) = vbFormFeed & "ВВЕДЕНИЕ" _
        Or Left((para.Range.Text), 10) = "ЗАКЛЮЧЕНИЕ" Or Left((para.Range.Text), 11) = vbFormFeed & "ЗАКЛЮЧЕНИЕ" _
            Or Left((para.Range.Text), 10) = "СОДЕРЖАНИЕ" Or Left((para.Range.Text), 11) = vbFormFeed & "СОДЕРЖАНИЕ" _
            Or (Left((para.Range.Text), 32) = "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ") Or (Left((para.Range.Text), 33) = vbFormFeed & "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ") Then
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

    ' Итерируемся по параграфам в документе
    For Each para In doc.Paragraphs
        ' Проверяем, начинается ли параграф с нужных слов
        If (Left(para.Range.Text, 9) = "ВВЕДЕНИЕ" & Chr(13) _
            Or Left(para.Range.Text, 11) = "ЗАКЛЮЧЕНИЕ" & Chr(13) _
            Or Left(para.Range.Text, 11) = "СОДЕРЖАНИЕ" & Chr(13) _
            Or Left(para.Range.Text, 33) = "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" & Chr(13)) Then
            
            ' Проверяем, есть ли предыдущий параграф
            If Not para.Previous Is Nothing Then
                ' Проверяем, не заканчивается ли предыдущий параграф на разрыв страницы
                If Right(para.Previous.Range.Text, 2) <> vbFormFeed & Chr(13) And Left(para.Range.Text, 1) <> vbFormFeed Then
                    ' Вставляем разрыв страницы перед текущим параграфом
                    
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
        
        
        ' Проверка, является ли абзац нумерованным
        If para.Range.listFormat.listType = wdListSimpleNumbering Then
            ' Удаляем нумерацию и применяем маркированный список
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
        If Left(para.Range.Text, 2) Like "[A-Z]" & "." Or Left(para.Range.Text, 2) Like "[А-Я]" & "." Then
        
            para.Range.Font.Color = wdColorRed
            para.Range.Comments.Add Range:=para.Range, Text:="ФИО должно ничинаться с фамилии"
            'With para.Range.Comments.Add(Range:=para.Range, Text:="ФИО должно ничинаться с фамилии")
                
            'End With
            
        End If
    Next para
End Sub
Sub Formula25()
  Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        ' Проверяем, содержит ли абзац символ "+"
        If InStr(para.Range.Text, "=") > 0 Then
            ' Выделяем весь абзац красным цветом
            para.Range.Font.Color = wdColorRed
            para.Range.Comments.Add Range:=para.Range, Text:="Возможно вы хотели бы здесь формулу" & Chr(13) & "Пример:A+B=C              (1.1)" & Chr(13) & "(формула прописывается через вставка->уравнение,формула выравнивается по левому краю, номер - по правому)"
            
        End If
    Next para
End Sub


Sub BlackLiterature26()
    Dim rng As Range
    

    ' Устанавливаем диапазон на весь документ
    Set rng = ActiveDocument.Content

    With rng.Find
        .Text = "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWholeWord = True
        
        If .Execute Then ' Если найдено слово "текст"
            ' Перемещаем диапазон в конец найденного слова
            rng.Collapse wdCollapseStart
            
            ' Устанавливаем новый диапазон от конца слова "текст" до конца документа
            Dim rngToHighlight As Range
            Set rngToHighlight = ActiveDocument.Content
            rngToHighlight.Start = rng.Start
            rngToHighlight.End = ActiveDocument.Content.End
            
            ' Выделяем текст
            rngToHighlight.Font.Color = -587137025
        
        End If
    End With
    MsgBox "ГОСТ соблюден"
End Sub
