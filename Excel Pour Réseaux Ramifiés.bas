'Copyright (c) 2011 Roland Yonaba

'This software is provided 'as-is', without any express or implied
'warranty. In no event will the authors be held liable for any damages
'arising from the use of this software.

'Permission is granted to anyone to use this software for any purpose,
'including commercial applications, and to alter it and redistribute it
'freely, subject to the following restrictions:

'    1. The origin of this software must not be misrepresented; you must not
'    claim that you wrote the original software. If you use this software
'   in a product, an acknowledgment in the product documentation would be
'    appreciated but is not required.

'    2. Altered source versions must be plainly marked as such, and must not be
'    misrepresented as being the original software.

'   3. This notice may not be removed or altered from any source
'    distribution.

'Code: Roland Yonaba
'Design Advisor: Priva Kabré
'Purpose: Ramified Water Network Analysis
'Last Revision : September 2011

Private Const VERSION = "0.1"

' Calculates the very number of sections for the current network
' This is relevant to the user
Function GetNumberOfSections() As Integer
    Sheets("Configuration").Activate
    Dim n As Integer
    n = Cells(4, 4).value
    GetNumberOfSections = n
End Function

' Gets the last filled line number for the current network on the sheet
Function GetEndLine(ByVal col As Integer) As Integer
Sheets("Dimensionnement").Activate
Dim curLine As Integer
curLine = 5
    Do While Cells(curLine, col).value <> ""
    curLine = curLine + 1
    Loop
GetEndLine = curLine - 1
End Function

'Sets a background color a specific range
Sub Colorize(rangee As String)
    Range(rangee).Select
    With selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
End Sub

' Applies static formulas all over the number of lines specified.
' Assumes each line is a unique section of the network
Sub FillWithFormulas()
   
    'Clears the sheet
    Clear
    
    'Copies the pattern on the hidden "Sav" sheet
    Sheets("Sav").Visible = True
    Sheets("Sav").Select
    Range("B5:O6").Select
    selection.Copy
    
    'Pastes the pattern on the computing sheet
    Sheets("Dimensionnement").Select
    Range("B5").Select
    ActiveSheet.Paste
    Sheets("Sav").Select
    ActiveWindow.SelectedSheets.Visible = False
    
    Dim n As Integer
    Dim n_sections As Integer
    Dim n_size As Integer
    Dim head_section As String
    Dim mimimum_pressure As Double
    
    'Collect minimum pressure, head section name, number of sections
    'They are all user's input
    n = GetNumberOfSections()
    If n < 2 Then n = 2
    n_sections = n + (6 - 1)
    n_size = n + (5 - 1) + 1
    head_section = Cells(9, 4).value
    mimimum_pressure = Cells(10, 4).value
    
    'Prepares each column of the computing sheet
    Sheets("Dimensionnement").Activate
    
    Range("B6").Select
    selection.AutoFill Destination:=Range("B6:B" & n_sections), Type:=xlFillDefault
    
    Colorize ("C6:C" & n_sections)
        
    Range("D6").Select
    selection.AutoFill Destination:=Range("D6:D" & n_sections), Type:=xlFillDefault
    Colorize ("D6:D" & n_sections)
    
    Range("E6").Select
    selection.AutoFill Destination:=Range("E6:E" & n_sections), Type:=xlFillDefault
    Colorize ("E6:E" & n_sections)
    
    Range("F6").Select
    selection.AutoFill Destination:=Range("F6:F" & n_sections), Type:=xlFillDefault
    Range("F6:F" & n_sections).Select
    selection.NumberFormat = "0.00"

    Range("G6").Select
    selection.AutoFill Destination:=Range("G6:G" & n_sections), Type:=xlFillDefault
    
    Range("H6").Select
    selection.AutoFill Destination:=Range("H6:H" & n_sections), Type:=xlFillDefault
       
    Range("I6").Select
    selection.AutoFill Destination:=Range("I6:I" & n_sections), Type:=xlFillDefault
      
    Range("J6").Select
    selection.AutoFill Destination:=Range("J6:J" & n_sections), Type:=xlFillDefault

    Range("L6").Select
    selection.AutoFill Destination:=Range("L6:L" & n_sections), Type:=xlFillDefault
    
    Colorize ("M6:M" & n_sections)
    
    Range("N6").Select
    selection.AutoFill Destination:=Range("N6:N" & n_sections), Type:=xlFillDefault
    
    Range("O6").Select
    selection.AutoFill Destination:=Range("O6:O" & n_sections), Type:=xlFillDefault
    Range("O6:O" & n_sections).Select
    
    'Sets the minimum pressure validation condition
    With selection.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator _
        :=xlGreaterEqual, Formula1:="" & mimimum_pressure
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = False
    End With
    
    'Writes the head section name
    Cells(6, 2).value = head_section
    
    'Draws borders
    SetBorders ("B5:O" & n_size)
    
End Sub

'Draws Borders within and around cells on the computing sheet
Sub SetBorders(rangee As String)
    Range(rangee).Select
    selection.Borders(xlDiagonalDown).LineStyle = xlNone
    selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

'Checks if section names follows the pattern "Node_Node"
Function CheckSectionNames() As Boolean
    Dim n As Integer
    Dim i As Integer
    Dim status As Boolean
    
    status = True
    n = GetEndLine(2)
    For i = 6 To n Step 1
        If Len(Cells(i, 2).value) <> 3 Then
        MsgBox ("Renommez les tronçons correctement!")
        status = False
        Exit For
        Else
             If Mid(Cells(i, 2).value, 2, 1) <> "_" Then
             MsgBox ("Renommez les tronçons correctement!")
             status = False
             Exit For
             End If
        End If
    Next i
    CheckSectionNames = status
End Function

' Computes lambda value for each section of the network
Sub ComputeNetwork()
    Dim n As Integer
    Dim i As Integer
    Dim checked As Boolean
    
    'Checks section names
    checked = CheckSectionNames()
    
    If checked Then
        'Computes uphill and downhill hydraulics heads formulas
        checked = SetHydraulics()
        n = GetEndLine(2)
        
        Sheets("Dimensionnement").Activate
        
        'Computes lamba for section per section
        For i = 6 To n Step 1
            Range("F" & i).GoalSeek Goal:=0, ChangingCell:=Range("H" & i)
        Next i
        
        'Circles low head hydraulics
        If checked Then
        ActiveSheet.CircleInvalid
        End If
        
    End If
End Sub

' Clears the current sheet and prepares a new one
Sub Clear()
    Dim n As Integer
    n = GetEndLine(2)
    Range("B5:O" & n).Select
    selection.ClearContents
    selection.Borders(xlDiagonalDown).LineStyle = xlNone
    selection.Borders(xlDiagonalUp).LineStyle = xlNone
    selection.Borders(xlEdgeLeft).LineStyle = xlNone
    selection.Borders(xlEdgeTop).LineStyle = xlNone
    selection.Borders(xlEdgeBottom).LineStyle = xlNone
    selection.Borders(xlEdgeRight).LineStyle = xlNone
    selection.Borders(xlInsideVertical).LineStyle = xlNone
    selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A1").Select
    selection.Copy
    Range("B5:O" & n).Select
    selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub

'Sets cumulative hydraulic head formulas for column "L"
Function SetHydraulics()
    Dim n As Integer
    Dim it As Integer
    Dim i As Integer
    Dim cur_section_uphill_node As String
    Dim ref_uphill As Integer
    Dim status As Boolean
    
    status = True
    n = GetEndLine(2)
    
    For it = 7 To n Step 1
        cur_section_uphill_node = Mid(Cells(it, 2).value, 1, 1)
        ref_uphill = 0
        For i = 6 To n Step 1
            If i <> it And (Mid(Cells(i, 2).value, 3, 1) = cur_section_uphill_node) Then
                ref_uphill = i
            End If
        Next i
        If (ref_uphill = 0) Then
            MsgBox ("Erreur détectée dans la définition des tronçons!" & vbNewLine & " Tronçons non continus!")
            status = False
            Exit For
        Else
            Cells(it, 11).Formula = "=L" & ref_uphill
        End If
    Next it
    SetHydraulics = status
End Function

'Displays help
Sub Help()
Dim Result As VbMsgBoxResult
Dim HelpMsg As String
HelpMsg = "1. Remplissez d'abord la feuille  'Configuration' avec les paramètres demandés et dans les unités spécifiées. Le réseau à calculer " _
            & vbNewLine & " doit avoir au MINIMUM 2 tronçons." _
            & vbNewLine & vbNewLine & "2. Cliquez ensuite sur le bouton 'Appliquer' pour passer à la feuille 'Dimensionnement'. Les cellules à fond " _
            & vbNewLine & " tramé sont celles dans lequelles vous devrez fournir des valeurs." _
            & vbNewLine & vbNewLine & "3. Spécifiez d'abord les noms des tronçons. Ceux-ci doivent OBLIGATOIREMENT respecter la syntaxe " _
            & "'Noeud_Noeud' , soit par exemple A_B, B_C, etc. Les tronçons doivent être continus." _
            & vbNewLine & vbNewLine & "4. Spécifiez les longueurs et débits véhiculés dans chaque tronçon, ainsi que l'altitude du noeud " _
            & "aval." _
            & vbNewLine & vbNewLine & "5. Cliquez enfin sur le bouton 'Calculer le réseau'. Les pressions dynamiques inférieures à la " _
            & "pression minimale de service que vous aurez auparavant spécifié dans la feuille 'Configuation' seront alors entourées en rouge."
            
            
        Result = MsgBox(HelpMsg, vbOKOnly + vbInformation, "Comment utiliser ce logiciel ?")
    
End Sub

' Displays credits
Sub Credits()
    Dim res As VbMsgBoxResult
    res = MsgBox("Code : Roland Yonaba" _
                & vbNewLine & "E-mail : roland.yonaba@gmail.com " _
                & vbNewLine & "Version : " + VERSION _
                & vbNewLine & vbNewLine & "Release Year : 2011 ", vbOKOnly + vbInformation, "Credits")
End Sub





