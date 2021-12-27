' ========================================
' EXPERIMENTAL
' ========================================

Option Explicit
Dim LastLink As String

' ========================================
' INTERNAL
' ========================================

Private Sub StartNewPar()

    Call GotoBeginPar
    Selection.TypeText Text:=vbNewLine
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
End Sub

Private Sub GotoEndPar()

  Selection.MoveLeft Unit:=wdCharacter, Count:=1
  Selection.MoveDown Unit:=wdParagraph, Count:=1
  
End Sub
Private Sub GotoBeginPar()

  Selection.MoveRight Unit:=wdCharacter, Count:=1
  Selection.MoveUp Unit:=wdParagraph, Count:=1
  
End Sub

Private Function AskLink() As String
    LastLink = InputBox(Prompt:="ßðëûê äëÿ îáúåêòà", Title:="Çàäàòü ÿðëûê")
    AskLink = LastLink
End Function

Private Function ProposeLink() As String
    ProposeLink = InputBox(Prompt:="ßðëûê äëÿ âñòàâêè", Title:="Çàäàòü ÿðëûê", Default:=LastLink)
End Function

Private Sub RefreshFields()
    Selection.Paragraphs(1).Range.Fields.Update
End Sub
' ========================================
' LINKS
' ========================================

Sub InputLink()
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="ref " & ProposeLink()
End Sub

' ========================================
' COUNTING
' ========================================

Sub InsAbstract()

    ' reset count fields
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq cfig \r 0 \h"
        
    Selection.TypeText Text:="Îò÷åò âûïîëíåí â 1 ÷àñòè è ñîäåðæèò: ñòðàíèö — "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="numpages"
    Selection.TypeText Text:=", èëëþñòðàöèé — "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="ref totfig"
    Selection.TypeText Text:=", òàáëèö — "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="ref tottbl"
    Selection.TypeText Text:=", ïðèëîæåíèé — "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="ref totapx"
    Selection.TypeText Text:=", â îò÷åòå èñïîëüçîâàíî èñòî÷íèêîâ — "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="ref totbib"
    Selection.TypeText Text:="."
    
End Sub

Sub InsEndCounters()
    Selection.EndKey Unit:=wdStory
    ' count figures
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="set totfig "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
            PreserveFormatting:=False, Text:="seq cfig \c"
    Selection.EndKey Unit:=wdLine
    ' count tables
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="set tottbl "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
            PreserveFormatting:=False, Text:="seq ctbl \c"
    Selection.EndKey Unit:=wdLine
    ' count equations
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="set toteq "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
            PreserveFormatting:=False, Text:="seq ceq \c"
    Selection.EndKey Unit:=wdLine
    ' count biblio
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="set totbib "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
            PreserveFormatting:=False, Text:="seq cbib \c"
    Selection.EndKey Unit:=wdLine
    ' count appendixes
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="set totapx "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
            PreserveFormatting:=False, Text:="seq capx \c"
    Selection.EndKey Unit:=wdLine

End Sub

' ========================================
' FIGURES
' ========================================

Sub InsFigName()

    Selection.TypeText Text:=vbNewLine

    Selection.Style = ActiveDocument.Styles("Íàçâàíèå ðèñóíêà")
    Selection.TypeText Text:="Ðèñóíîê "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq fig \n"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq cfig \n \h"
        
    Dim Link As String
    Link = AskLink()
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="set " & Link & " """
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq fig \c"
    Selection.TypeText Text:=""""
    Selection.MoveRight Unit:=wdCharacter, Count:=2
        
    Selection.TypeText Text:=" — "

    Call RefreshFields
End Sub

Sub InsFigNameWholeDoc()

    Selection.TypeText Text:=vbNewLine

    Selection.Style = ActiveDocument.Styles("Íàçâàíèå ðèñóíêà")
    Selection.TypeText Text:="Ðèñóíîê "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq wfig \n"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq cfig \n \h"
        
    Dim Link As String
    Link = AskLink()
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="set " & Link & " """
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq wfig \c"
    Selection.TypeText Text:=""""
    Selection.MoveRight Unit:=wdCharacter, Count:=2
        
    Selection.TypeText Text:=" — "

    Call RefreshFields
End Sub

Sub FormatFig()

    Selection.Style = ActiveDocument.Styles("Èçîáðàæåíèå ðèñóíêà")
    
End Sub

' ========================================
' TABLES
' ========================================

Private Sub SetTableBottomMargin()

    With Selection.Tables(1).Rows
        .WrapAroundText = True
        .HorizontalPosition = wdTableCenter
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionColumn
        .DistanceLeft = CentimetersToPoints(0.32)
        .DistanceRight = CentimetersToPoints(0.32)
        .VerticalPosition = CentimetersToPoints(0)
        .RelativeVerticalPosition = wdRelativeVerticalPositionParagraph
        .DistanceTop = CentimetersToPoints(0)
        .DistanceBottom = CentimetersToPoints(0.5)
        .AllowOverlap = False
    End With
    
End Sub


Sub InsTblName()

    Selection.Style = ActiveDocument.Styles("Íàçâàíèå òàáëèöû")
    Selection.TypeText Text:="Òàáëèöà "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq tbl \n"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ctbl \n \h"
    
    Dim Link As String
    Link = AskLink()
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="set " & Link & " """
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq tbl \c"
    Selection.TypeText Text:=""""
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    
    Selection.TypeText Text:=" — "
    
    Call RefreshFields
End Sub

Sub InsTblNameWholeDocument()

    Selection.Style = ActiveDocument.Styles("Íàçâàíèå òàáëèöû")
    Selection.TypeText Text:="Òàáëèöà "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq wtbl \n"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ctbl \n \h"
    
    Dim Link As String
    Link = AskLink()
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="set " & Link & " """
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq wtbl \c"
    Selection.TypeText Text:=""""
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    
    Selection.TypeText Text:=" — "
    
    Call RefreshFields
End Sub

Sub FormatTable()
    Selection.Tables(1).Range.Style = ActiveDocument.Styles("Òåêñò òàáëèöû")
    Selection.Tables(1).Columns(1).Select
    Selection.Style = ActiveDocument.Styles("Áîêîâèê òàáëèöû")
    Selection.Tables(1).Rows(1).Range.Style = ActiveDocument.Styles("Çàãîëîâîê ãðàôû")
    Selection.Tables(1).Rows(1).HeadingFormat = True
    Call SetTableBottomMargin
End Sub

' ========================================
' ---
' ========================================

Sub ApplyStrElement()

    Selection.Style = ActiveDocument.Styles("Çàãîëîâîê ñòðóêòóðíîãî ýëåìåíòà")
    
End Sub

Sub ApplyDefault()

    Selection.Style = ActiveDocument.Styles("Îáû÷íûé")
    
End Sub

Sub ApplyFontHighlight()

    Selection.Style = ActiveDocument.Styles("Âûäåëèòü òåêñò øðèôòîì")
    
End Sub

Sub ClearFields()

    Selection.Paragraphs(1).Range.Select
    Selection.Fields.Unlink
    
End Sub

' ========================================
' EQUATIONS
' ========================================

Sub InsEq()

    Call StartNewPar
    Selection.Style = ActiveDocument.Styles("Óðàâíåíèå")
    Selection.TypeText Text:="("
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq eq \n"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ceq \n \h"
    Selection.TypeText Text:=")"
    Selection.HomeKey Unit:=wdLine
    Selection.TypeText Text:=vbTab & vbTab
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    Call RefreshFields
End Sub

Sub InsEqWholeDocument()

    Call StartNewPar
    Selection.Style = ActiveDocument.Styles("Óðàâíåíèå")
    Selection.TypeText Text:="("
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq weq \n"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ceq \n \h"
    Selection.TypeText Text:=")"
    Selection.HomeKey Unit:=wdLine
    Selection.TypeText Text:=vbTab & vbTab
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    Call RefreshFields
End Sub

Sub InsEqDesc()
    
    Call GotoEndPar
    Selection.TypeText Text:=vbNewLine
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.Style = ActiveDocument.Styles("Ïîäïèñü óðàâíåíèÿ")
    Selection.TypeText Text:=vbTab & vbTab & "—" & vbTab
    Selection.MoveLeft Unit:=wdCharacter, Count:=3
    
End Sub

Sub InsEqLink()
    Selection.EndKey Unit:=wdLine
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    Dim Link As String
    Link = AskLink()
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="set " & Link & " """
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq eq \c"
    Selection.TypeText Text:=""""
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    
    Call RefreshFields
End Sub

' ========================================
' CHAPTERS
' ========================================

Sub InsCh()
    Call ClearFields
    Call GotoBeginPar
    Selection.Style = ActiveDocument.Styles("Ðàçäåë")
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ch \n"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq subch \r 0 \h"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq eq \r 0 \h"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq tbl \r 0 \h"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq fig \r 0 \h"
    Selection.TypeText Text:=" "
    
    Call RefreshFields
End Sub

Sub InsSubch()
    Call ClearFields
    Call GotoBeginPar
    Selection.Style = ActiveDocument.Styles("Ïîäðàçäåë")
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq subch \n"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq pnt \r 0 \h"
    Selection.TypeText Text:=" "
    
    Call RefreshFields
End Sub

Sub InsPnt()
    Call ClearFields
    Call GotoBeginPar
    Selection.Style = ActiveDocument.Styles("Ïóíêò")
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq subch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq pnt \n"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq subpnt \r 0 \h"
    Selection.TypeText Text:=" "
    
    Call RefreshFields
End Sub

Sub InsSubpnt()
    Call ClearFields
    Call GotoBeginPar
    Selection.Style = ActiveDocument.Styles("Ïîäïóíêò")
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq ch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq subch \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq pnt \c"
    Selection.TypeText Text:="."
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq subpnt \n"
    Selection.TypeText Text:=" "
    
    Call RefreshFields
End Sub

Sub InsApx()

    Dim Link As String
    Link = AskLink()

    Selection.Style = ActiveDocument.Styles("Ïðèëîæåíèå")
    Selection.TypeText Text:="Ïðèëîæåíèå "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="symbol "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="= "
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="seq apx \n"
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:=" + 1039"
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:=" \u"
    Selection.Fields.Update
    Selection.EndKey Unit:=wdLine
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq capx \n \h"
    
    Call RefreshFields
    
End Sub

' ========================================
' BIBLIO
' ========================================

Sub InsBib()

    Dim Link As String
    Link = AskLink()
    
    Call StartNewPar
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq bib \n"
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq cbib \n \h"
    
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False
    Selection.TypeText Text:="set " & Link & " """
    Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, _
        PreserveFormatting:=False, Text:="seq bib \c"
    Selection.TypeText Text:=""""
    
    Selection.MoveRight Unit:=wdCharacter, Count:=2
    Selection.TypeText Text:=". "
    
    Call RefreshFields
    
End Sub

' ========================================
' EXPERIMENTAL
' ========================================

Sub SplitTable()
    Selection.Tables(1).Split BeforeRow:=True
    Selection.MoveUp Unit:=wdLine, Count:=1
End Sub

