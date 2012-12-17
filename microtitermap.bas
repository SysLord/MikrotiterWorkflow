Attribute VB_Name = "microtitermap"
' Copyright 2010 Christian Helmer
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program. If not, see <http://www.gnu.org/licenses/>.
'
' ====================================================================
'
' Author: Christian Helmer
' Date: 17.11.2010
' For Doro
'
' -----------------------------------------------------------------------------
'
' Synthesis plan for peptoidsequences with mikrotiter plate view
'
'
' Display all combinations of selected table.
' example:
' A J X
' B K Y
'   L Z   => AJX, AJY, AJZ, AKX, AKY...
'
' Usage: Select the table and run makro.
'
'

Option Explicit


' userdefined colors
Public colorTable As Range

' the used sheets
Public permList As Worksheet
Public permList2 As Worksheet
Public layout As Worksheet
Public workProcedure As Worksheet

Public colorSheet As Worksheet


' List  A F X
Public posSimpleList As Range

' Liste  A | F | X
Public posSimpleBlocks As Range

'blocks A A
'       F F ...
'       X Y
Public posBlocks As Range
Public posColoredBlocks As Range

Public blockIdx As Long
Public blockCount As Long
Public superblockIdx As Long

'blockdimensions
Public blockWidth As Long
Public blockHeight As Long
Public blockYDist As Long
Public blockXDist As Long

Public blockXTitles As Variant 'As String array
Public blockYTitles As Variant 'As String array

Public perBlockPageBreak As Boolean

Public colorp As Long
Public assocColor As New Collection

'height, width of selection
Public xlength As Long
Public ylength As Long

' Makro Entrypoint
Sub createMicrotiterMap()
Attribute createMicrotiterMap.VB_Description = "Syntheseplan für Peptoidsequenzen + Mikrotiterplatten-Ansicht"
Attribute createMicrotiterMap.VB_ProcData.VB_Invoke_Func = "l\n14"

'List of elements (= 'words')
Dim words() As String

Dim x As Long
Dim y As Long

Set permList = Worksheets("permutated list")
Set permList2 = Worksheets("permutated list 2")
Set layout = Worksheets("mikrotiter plate")
Set workProcedure = Worksheets("workflow")

Set colorSheet = Worksheets("user colors")


xlength = Selection.Columns.Count
ylength = Selection.Rows.Count

' one cell selected
If xlength < 2 Or ylength < 2 Then

    ' clears sheets if only one cell with content 'clear' is selected
    If CStr(Selection.Cells(1, 1)) = "clear" Then
        clearSheets
    End If
    
    Exit Sub
End If

' space for longest possible word
ReDim words(0 To xlength - 1)

clearSheets

Set colorTable = colorSheet.Range("C4")

' combinations in one column
Set posSimpleList = permList.Range("B2")

' combinations with each word in own column
Set posSimpleBlocks = permList2.Range("B2")

' raw view
Set posBlocks = layout.Range("B3")

' colored view in superblocks to be printed
Set posColoredBlocks = workProcedure.Range("B3")

workProcedure.Cells.EntireColumn.ColumnWidth = 4

' each block on new page
perBlockPageBreak = True

blockWidth = 12
blockHeight = 8

' labels
' arraylength needs to mach blockWidth and blockHeight
blockXTitles = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12")
blockYTitles = Array("A", "B", "C", "D", "E", "F", "G", "H")

' manually chosen colors
'colors = Array(3, 4, 7, 8, 15, 16, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 50)
' mixed
'colors = Array(3, 37, 36, 44, 46, 16, 42, 39, 33, 35, 45, 34, 4, 8, 47, 40, 38, 43, 41, 7, 50, 15)

' colorpointer
colorp = 0

' dumps all 57 color indices
' dumpColors colorTable

' ----------------------------------------------------------------------------------------

' associate color to each cell

Set assocColor = New Collection

For x = 1 To xlength
    For y = 1 To ylength
        If Not IsEmpty(Selection.Cells(y, x)) Then
            ' same color for same name
            If Not InCollection(assocColor, (CStr(Selection.Cells(y, x)))) Then
                Dim cellColor As Long
                cellColor = colorTable.Offset(colorp, 0).Interior.Color
                assocColor.Add cellColor, CStr(Selection.Cells(y, x))
                colorp = colorp + 1
            End If
        End If
    Next
Next

' superblock spacings
blockYDist = 4
blockXDist = 2

' superblock counter for positioning
blockIdx = 0
blockCount = 0
superblockIdx = 0

' main function
choose words, 0, 1


'optional: autofit colortables
'Range( _
'    posColoredBlocks, _
'    posColoredBlocks.Cells(0, (blockXDist * ylength + blockWidth * (ylength + 1))) _
').EntireColumn.AutoFit

' labels included:
Range( _
    posColoredBlocks.Offset(-1, -1), _
    posColoredBlocks.Cells(0, (blockXDist * ylength + blockWidth * (ylength + 1))) _
).EntireColumn.AutoFit


' pagebreaks around every colored block
If perBlockPageBreak = True Then
    'Columns
    For x = 1 To ylength + 1
        posColoredBlocks.Columns( _
            (blockXDist + blockWidth) * x).PageBreak = xlPageBreakManual
    Next x
    
End If

End Sub

' Recursive function to generate every combination in the prefix variable
Private Sub choose(prefix() As String, prefixp As Long, x As Long)

Dim y As Long


' yield prefix if end of selection is reached
If x > xlength Then
    yieldCombination prefix, x
    Exit Sub
End If

' recursion
' for every row
For y = 1 To ylength
    ' if not empty
    If Not IsEmpty(Selection.Cells(y, x)) Then
        ' add word to prefix
        prefix(prefixp) = Selection.Cells(y, x)
        ' recurse and use next column and next prefix position
        choose prefix, prefixp + 1, x + 1
    End If
Next

End Sub

' Output combination
'
' what blocks and superblocks are:
'
'    Cols
' _________
' |  1 2 3 ..
' |
' |A M .      <| row
' |  N .      <|______
'
' |B V .      <| row
' |  W .      <|______
'
' |C H .      <| row
' |  I .      <|______
' ...
'
'
'
'        Superblock 0
' ,-----------I---- ... ---.
' _______  _______
' |Block|  |Block|  ...
' |     |  |     |
' |     |  |     |
' ¯¯¯¯¯¯¯  ¯¯¯¯¯¯¯
'
'        Superblock 1
'            ...
'
Private Sub yieldCombination(prefix() As String, x As Long)

    Dim pos As Long
    Dim yoffs As Long
    Dim blockcol As Long
    Dim blockrow As Long
    
    Dim colorBlockIdx As Long
    
    '1. simple list
    posSimpleList.Offset(blockCount, 0) = Join(prefix, " ")
    
    '2. new column for every word
    For pos = 0 To UBound(prefix)
        posSimpleBlocks.Offset(blockCount, pos) = prefix(pos)
    Next
        
    '3. blocks
    blockcol = blockIdx Mod blockWidth
    blockrow = (blockIdx \ blockWidth) '* xlength
    
    ' calculate superblock y offset
    yoffs = superblockIdx * (xlength * blockHeight + blockYDist) + blockrow * xlength
    
    'Labels (need space to the left and above)
    'left labels
    If blockcol = 0 Then
        posBlocks.Offset(yoffs, blockcol - 1) = blockYTitles(blockrow)
        
        For colorBlockIdx = 0 To xlength - 1
            posColoredBlocks.Offset(yoffs, _
            blockcol - 1 + (blockWidth + blockXDist) * colorBlockIdx) _
                = blockYTitles(blockrow)
        Next
    End If
    
    ' top labels
    If blockrow = 0 Then
        posBlocks.Offset(yoffs - 1, blockcol) = blockXTitles(blockcol)
        For colorBlockIdx = 0 To xlength - 1
            posColoredBlocks.Offset(yoffs - 1, _
            blockcol + (blockWidth + blockXDist) * colorBlockIdx) _
                = blockXTitles(blockcol)
        Next
    End If
        
    'words
    For pos = 0 To xlength - 1 'UBound(prefix)
        posBlocks.Offset(yoffs + pos, blockcol) = prefix(pos)
        
        For colorBlockIdx = 0 To xlength - 1
            If pos = colorBlockIdx Then
                posColoredBlocks.Offset( _
                    yoffs + pos, _
                    blockcol + (blockWidth + blockXDist) * colorBlockIdx) _
                    .Interior.Color = assocColor(prefix(pos))
            Else
                'Optional: shrink unimportant entries
                'posColoredBlocks.Offset(yoffs + pos, _
                    blockcol + (blockWidth + blockXDist) * colorBlockIdx).Font.Size = 8
            End If
            
            posColoredBlocks.Offset(yoffs + pos, _
                blockcol + (blockWidth + blockXDist) * colorBlockIdx) = prefix(pos)
        Next
    Next
        
    'Borders
    With posBlocks
        BorderAround .Range( _
            .Cells(yoffs - 1, blockcol), _
            .Cells(yoffs + xlength - 2, blockcol)), 1
    End With
    
    With posColoredBlocks
        For colorBlockIdx = 0 To xlength - 1
            BorderAround .Range( _
                .Cells(yoffs - 1, _
                    blockcol + (blockWidth + blockXDist) * colorBlockIdx), _
                .Cells(yoffs + xlength - 2, _
                    blockcol + (blockWidth + blockXDist) * colorBlockIdx)), 1
        Next
    End With
    
    
    'Block Indices
    blockIdx = blockIdx + 1
    blockCount = blockCount + 1
    
    If blockIdx >= blockWidth * blockHeight Then
        superblockIdx = superblockIdx + 1
        blockIdx = 0
        
        ' pagebreak here necause of new superblock
        If perBlockPageBreak = True Then
            Dim breakyoffs As Long
            
            breakyoffs = superblockIdx * (xlength * blockHeight + blockYDist) - 1
            posColoredBlocks.Offset(breakyoffs, 0).Rows(0).PageBreak = xlPageBreakManual
        End If
        
    End If
    
End Sub

' Reset sheets
' Note: workProcedure is deleted, so old references
' like posColoredBlock are not valid anymore.
Private Sub clearSheets()

        With permList
            .Cells.ClearContents
            .ResetAllPageBreaks
        End With
        With permList2
            .Cells.ClearContents
            .ResetAllPageBreaks
        End With
        With layout
            .Cells.Clear
            .ResetAllPageBreaks
        End With
        With workProcedure
            .Cells.Clear
            .Cells.EntireColumn.ColumnWidth = 4
            
            ' does not work
            '.UsedRange.PageBreak = xlPageBreakNone
            ' works, but resets pagebreaks
            .ResetAllPageBreaks
            
            
        End With

End Sub

'********************************************************
'*                       Helpers                        *
'********************************************************

'thank you stackoverflow.com
Private Function InCollection(col As Collection, key As String) As Boolean
  Dim var As Variant
  Dim errNumber As Long

  InCollection = False
  Set var = Nothing

  Err.Clear
  On Error Resume Next
    var = col.item(key)
    errNumber = CLng(Err.Number)
  On Error GoTo 0

  '5 is not in, 0 and 438 represent incollection
  If errNumber = 5 Then ' it is 5 if not in collection
    InCollection = False
  Else
    InCollection = True
  End If

End Function

' border around range with given colorindex
Private Sub BorderAround(r As Range, cidx As Integer)
    Dim myBorders() As Variant, item As Variant
    myBorders = Array( _
        xlEdgeLeft, _
        xlEdgeTop, _
        xlEdgeBottom, _
        xlEdgeRight)
    For Each item In myBorders
        With r.Borders(item)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = cidx 'xlAutomatic
        End With
    Next item
    
End Sub

' dump colors
Private Sub dumpColors(r As Range)

Dim i As Integer
Dim Color As Variant
i = 0

For i = 0 To 56
    r.Offset(i, 3).Interior.ColorIndex = i
Next i

End Sub
