Dim LastLink        As String

Sub OnLoad(ribbon   As IRibbonUI)
End Sub

Private Sub insertNewParagraph()
    Call goToBeginningOfParagraph 
    Selection.TypeText Text:=vbNewLine
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
End Sub

Private Sub goToBeginningOfParagraph 
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveUp Unit:=wdParagraph, Count:=1
    
End Sub

Private Sub goToEndOfParagraph()
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.MoveDown Unit:=wdParagraph, Count:=1
End Sub

Private Sub refreshFields()
    Selection.Paragraphs(1).Range.Fields.Update
End Sub

Private Function askForLink()
    LastLink = InputBox(Prompt:="Ярлык для объекта", Title:="Задать ярлык")
    askForLink = LastLink
End Function

Private Function proposeLink()
    proposeLink = InputBox(Prompt:="Ярлык для вставки", Title:="Задать ярлык",Default:=LastLink)
End Function

Sub InputLink()
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, Text:="ref " & ProposeLink()
End Sub

Sub InsertAbstract(control As IRibbonControl)
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, Text:="seq cfig \r 0 \h"
    
    Selection.TypeText Text:="Отчёт выполнен в 1 части и содержит: страниц — "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False, Text:="numpages"
    Selection.TypeText Text:=", иллюстраций — "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty,PreserveFormatting:=False, Text:="ref totfig"
    Selection.TypeText Text:=", таблиц — "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="ref tottbl"
    Selection.TypeText Text:=", приложений — "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="ref totapx"
    Selection.TypeText Text:=", в отчёте использовано источников — "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="ref totbib"
    Selection.TypeText Text:="."
End Sub

Sub InsertEndingCounters(control As IRibbonControl)
    Selection.EndKey Unit:=wdStory
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="set totfig "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq cfig \c"
    Selection.EndKey Unit:=wdLine
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="set tottbl "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ctbl \c"
    Selection.EndKey Unit:=wdLine
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="set toteq "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ceq \c"
    Selection.EndKey Unit:=wdLine
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="set totbib "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq cbib \c"
    Selection.EndKey Unit:=wdLine
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="set totapx "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq capx \c"
    Selection.EndKey Unit:=wdLine
End Sub

Sub InsertFigureNameChapterNumber(control As IRibbonControl)
    
    Selection.TypeText Text:=vbNewLine
    
    Selection.Style = ActiveDocument.Styles("Название рисунка")
    Selection.TypeText Text:="Рисунок "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq fig \n"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq cfig \n \h"
    
    Dim Link        As String
    Link = askForLink()
    
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="set " & Link & " """
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq fig \c"
    Selection.TypeText Text:=""""
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    
    Selection.TypeText Text:=" — "
    
    Call refreshFields()
End Sub

Sub InsertFigureNameEndToEndNumber(control As IRibbonControl)
    
    Selection.TypeText Text:=vbNewLine
    
    Selection.Style = ActiveDocument.Styles("Название рисунка")
    Selection.TypeText Text:="Рисунок "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq wfig \n"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq cfig \n \h"
    
    Dim Link        As String
    Link = askForLink()
    
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="set " & Link & " """
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq wfig \c"
    Selection.TypeText Text:=""""
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    
    Selection.TypeText Text:=" — "
    
    Call refreshFields()
End Sub

Private Sub setTableBottomMargin()
    
    With Selection.Tables(1).Rows
        .WrapAroundText = TRUE
        .HorizontalPosition = wdTableCenter
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
        .DistanceLeft = CentimetersToPoints(0.32)
        .DistanceRight = CentimetersToPoints(0.32)
        .VerticalPosition = CentimetersToPoints(0)
        .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
        .DistanceTop = CentimetersToPoints(0)
        .DistanceBottom = CentimetersToPoints(0.5)
        .AllowOverlap = FALSE
    End With
    
End Sub

Sub InsertTableNameChapterNumber(control As IRibbonControl)
    
    Selection.Style = ActiveDocument.Styles("Название таблицы")
    Selection.TypeText Text:="Таблица "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq tbl \n"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ctbl \n \h"
    
    Dim Link        As String
    Link = askForLink()
    
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="set " & Link & " """
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq tbl \c"
    Selection.TypeText Text:=""""
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    
    Selection.TypeText Text:=" — "
    
    Call refreshFields()
End Sub

Sub InsertTableNameEndToEndNumber(control As IRibbonControl)
    
    Selection.Style = ActiveDocument.Styles("Название таблицы")
    Selection.TypeText Text:="Таблица "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq wtbl \n"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ctbl \n \h"
    
    Dim Link        As String
    Link = askForLink()
    
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="set " & Link & " """
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq wtbl \c"
    Selection.TypeText Text:=""""
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    
    Selection.TypeText Text:=" — "
    
    Call refreshFields()
End Sub

Sub InsertEquationChapterNumber(control As IRibbonControl)
    
    Call insertNewParagraph()
    Selection.Style = ActiveDocument.Styles("Уравнение")
    Selection.TypeText Text:="("
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq eq \n"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ceq \n \h"
    Selection.TypeText Text:=")"
    Selection.HomeKey Unit:=wdLine
    Selection.TypeText Text:=vbTab & vbTab
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    Call refreshFields()
End Sub

Sub InsertEquationEndToEndNumber(control As IRibbonControl)
    
    Call insertNewParagraph()
    Selection.Style = ActiveDocument.Styles("Уравнение")
    Selection.TypeText Text:="("
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq weq \n"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ceq \n \h"
    Selection.TypeText Text:=")"
    Selection.HomeKey Unit:=wdLine
    Selection.TypeText Text:=vbTab & vbTab
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    Call refreshFields()
End Sub

Sub InsertEquationDescription(control As IRibbonControl)
    
    Call goToEndOfParagraph()
    Selection.TypeText Text:=vbNewLine
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Style = ActiveDocument.Styles("Подпись уравнения")
    Selection.TypeText Text:=vbTab & vbTab & "—" & vbTab
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
    
End Sub

Sub InsertEquationLink(control As IRibbonControl)
    Selection.EndKey Unit:=wdLine
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    Dim Link        As String
    Link = askForLink()
    
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="set " & Link & " """
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq eq \c"
    Selection.TypeText Text:=""""
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    
    Call refreshFields()
End Sub

Sub InsertChapter(control As IRibbonControl)
    ClearFields 
    Call goToBeginningOfParagraph 
    Selection.Style = ActiveDocument.Styles("Раздел")
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ch \n"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq subch \r 0 \h"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq eq \r 0 \h"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq tbl \r 0 \h"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq fig \r 0 \h"
    Selection.TypeText Text:=" "
    
    Call refreshFields()
End Sub

Sub InsertSubChapter(control As IRibbonControl)
    ClearFields 
    Call goToBeginningOfParagraph 
    Selection.Style = ActiveDocument.Styles("Подраздел")
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq subch \n"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq pnt \r 0 \h"
    Selection.TypeText Text:=" "
    
    Call refreshFields()
End Sub

Sub InsertPoint(control As IRibbonControl)
    ClearFields 
    Call goToBeginningOfParagraph
    Selection.Style = ActiveDocument.Styles("Пункт")
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq subch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq pnt \n"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq subpnt \r 0 \h"
    Selection.TypeText Text:=" "
    
    Call refreshFields
End Sub

Sub InsertSubPoint(control As IRibbonControl)
    ClearFields
    Call goToBeginningOfParagraph 
    Selection.Style = ActiveDocument.Styles("Подпункт")
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False, Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq subch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq pnt \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq subpnt \n"
    Selection.TypeText Text:=" "
    
    Call refreshFields()
End Sub

Sub InsertAppendix(control As IRibbonControl)
    
    Dim Link        As String
    Link = askForLink()
    
    Selection.Style = ActiveDocument.Styles("Приложение")
    Selection.TypeText Text:="Приложение "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="symbol "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="= "
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="seq apx \n"
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:=" + 1039"
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:=" \u"
    Selection.Fields.Update 
    Selection.EndKey Unit:=wdLine
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq capx \n \h"
    
    Call refreshFields()
    
End Sub

Sub InsertBibliographyItem(control As IRibbonControl)
    
    Dim Link        As String
    Link = askForLink()
    
    Call insertNewParagraph()
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq bib \n"
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq cbib \n \h"
    
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False
    Selection.TypeText Text:="set " & Link & " """
    Selection.Fields.Add Range:=Selection.Range,Type:=wdFieldEmpty,PreserveFormatting:=False,Text:="seq bib \c"
    Selection.TypeText Text:=""""
    
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:=". "
    
    Call refreshFields()
    
End Sub

Sub FormatAsFigure(control As IRibbonControl)
    Selection.Style = ActiveDocument.Styles("Изображение рисунка")
    
End Sub

Sub FormatAsTable(control As IRibbonControl)
    Selection.Tables(1).Range.Style = ActiveDocument.Styles("Текст таблицы")
    Selection.Tables(1).Columns(1).Select
    Selection.Style = ActiveDocument.Styles("Боковик таблицы")
    Selection.Tables(1).Rows(1).Range.Style = ActiveDocument.Styles("Заголовок графы")
    Selection.Tables(1).Rows(1).HeadingFormat = TRUE
    Call setTableBottomMargin
End Sub

Sub FormatAsStructureElementStyle(control As IRibbonControl)
    
    Selection.Style = ActiveDocument.Styles("Заголовок структурного элемента")
    
End Sub

Sub FormatAsDefault(control As IRibbonControl)
    Selection.Style = ActiveDocument.Styles("Обычный")
    
End Sub

Sub FormatAsHighlightedText(control As IRibbonControl)
    Selection.Style = ActiveDocument.Styles("Выделить текст шрифтом")
    
End Sub

Sub ClearFields
    Selection.Paragraphs(1).Range.Select
    Selection.Fields.Unlink
    
End Sub