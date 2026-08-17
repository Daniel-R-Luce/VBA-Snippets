VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' @TITLE: Tiny StringBuilder
' @OVERVIEW:
'   A lightweight utility for building multi-line strings such as SQL queries,
'   logs, or formatted text blocks. Designed to make complex string assembly
'   more human-readable and maintainable than raw concatenation.
'
' @METHODS:
'   - Append: add text to the current line in the StringBuilder
'   - AppendLine: add text with automatic line breaks
'   - ToString: return the full assembled string
'   - Debugging properties: Length, LineCount, IsEmpty
'
' @NOTES:
'   - Internally uses a Collection for efficient line storage
'   - No Clear method: encourages fresh instances to avoid state errors
'   - Minimal surface area for ease of use by VBA beginners
'================================================================================

Private lines As Collection

Private Sub Class_Initialize()
    Set lines = New Collection
End Sub


Public Sub Append(ByVal text As String)
    If lines.Count = 0 Then
        lines.Add text
    Else
        Dim lastVal As String
        lastVal = lines(lines.Count)
        lines.Remove lines.Count
        lines.Add lastVal & text
    End If
End Sub

Public Sub AppendLine(ByVal text As String)
    lines.Add text
End Sub

Public Function ToString() As String
    Dim output() As String
    Dim i As Long
    
    ReDim output(1 To lines.Count)
    
    For i = 1 To lines.Count
        output(i) = lines(i)
    Next i
    
    ToString = Join(output, vbCrLf)
End Function

Public Property Get Length() As Long
    Length = Len(Me.ToString)
End Property

Public Property Get LineCount() As Long
    LineCount = lines.Count
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = (lines.Count = 0)
End Property

