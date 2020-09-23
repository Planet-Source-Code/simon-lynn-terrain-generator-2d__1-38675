VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Terrain Generator"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   294
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##########################
'#                        #
'#  2D Terrain Generator  #
'#     by Simon Lynn      #
'#                        #
'##########################

' Anyone can use this code in their projects if they find it useful.

Const ITERATIONS = 8                     ' Increase this value to increase the size of the map
Const ROW = (2 ^ ITERATIONS) + 1         ' The width of the map when drawn. Used in calculations.
Const SLOPE = 15                         ' This value affects how "jagged" the terrain is

Dim Heights(ROW - 1, ROW - 1) As Single  ' Array of height values
Dim R, G, B, Bound                       ' Variant array variables

' SetPixelV API call is faster than PSet
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Sub Form_Load()

Dim Seed As Single ' Seed value for the random number generator

    ' This array defines the boundaries for different colors
    Bound = Array(-10, 0, 15, 25, 30, 35, 40, 45)
    
    ' These arrays define the terrain colours for the heights boundaries.
    R = Array(20, 23, 15, 20, 136, 145, 91, 88, 142)
    G = Array(14, 15, 85, 113, 136, 88, 55, 88, 142)
    B = Array(171, 199, 9, 11, 56, 41, 26, 88, 142)
    
    ' Change these values to create different effects. For a more defined
    ' coastline, make the first two Bound values the same. If you create
    ' more color boundaries, the terrain can become more detailed.

    Seed = Timer   ' Generate the map's seed based on the system timer.
    Randomize Seed ' It should be possible to obtain previous maps by using their seed values here
    Caption = "Terrain Generator    Seed = " & Seed

End Sub

Private Sub Form_Resize()

    Form_Click

End Sub

Private Sub Form_Click()

    Generate_Heights
    Draw

End Sub

Private Sub Generate_Heights()

' This sub generates the height map which is placed in the Heights array. The
' heights define which color is shown when the map is drawn.
'
' The method it uses is quite difficult to explain, but here goes...
'
' The process of generation is best described by the following progression:
'
'   Initial    |  i = 0      |  i = 1
'   state:     |             |
'   1       1  |  1   2   1  |  1 3 2 3 1
'              |             |  3 3 3 3 3
'              |  2   2   2  |  2 3 2 3 2      etc.
'              |             |  3 3 3 3 3
'   1       1  |  1   2   1  |  1 3 2 3 1
'
' Initially, the four corners are given random height values. Then, in each
' iteration height values are created for the points that are halfway between
' the points in the last iteration, and the points in the middle of four values.
' So, when i = 0, the "2"s in the diagram represent the added height values, and
' when i = 1, the "3"s in the diagram are the next lot of added height values.
' This is repeated until the required number of iterations is reached. The
' height values for the new points are determined by taking the average of
' surrounding points and adding a random displacement to this. The amount of
' displacement from the average is determined by the SLOPE constant, and also
' by the iteration counter, i. The higher the value of i, the lower the
' displacement, so that the first few iterations have a large influence on the
' shape of the terrain but later iterations do not, producing only minor
' variations on the present state of the terrain. This makes the terrain less
' jagged and more rounded.
'
' Using this method, it is theoretically possible to zoom in on the map
' indefinetely to uncover new levels of detail on the same map.

Dim i       As Long     ' Iteration counter
Dim x       As Long     ' Coordinate
Dim y       As Long     '            counters
Dim Step    As Long     ' Used in calculations in the main loop

    Erase Heights

    ' Inistialise the heights in each of the corners
    Heights(0, 0) = Rnd * -10
    Heights(0, ROW - 1) = Rnd * -10
    Heights(ROW - 1, 0) = Rnd * -10
    Heights(ROW - 1, ROW - 1) = Rnd * -10
        
    ' The main iteration loop. See above for explanation.
    ' There may be a better algorithm for this, but this is the fasted one
    ' that I could come up with.
    For i = 0 To ITERATIONS - 1
        Step = (ROW - 1) / (2 ^ (i + 1))
        For y = 0 To ROW - 1 Step Step
            For x = 0 To ROW - 1 Step Step
                If (y / Step) Mod 2 = 0 Then
                    If Heights(x, y) = 0 Then
                        Heights(x, y) = (Heights(x - Step, y) + Heights(x + Step, y)) / 2 + (Rnd * SLOPE - (SLOPE / 2)) * ((ITERATIONS - i) / (i + 1))
                    End If
                Else
                    If (x / Step) Mod 2 = 0 Then
                        Heights(x, y) = (Heights(x, y - Step) + Heights(x, y + Step)) / 2 + (Rnd * SLOPE - (SLOPE / 2)) * ((ITERATIONS - i) / (i + 1))
                    Else
                        Heights(x, y) = (Heights(x - Step, y - Step) + Heights(x - Step, y + Step) + Heights(x + Step, y - Step) + Heights(x + Step, y + Step)) / 4 + (Rnd * SLOPE - (SLOPE / 2)) * ((ITERATIONS - i) / (i + 1))
                    End If
                End If
            Next
        Next
    Next
    
End Sub
    
Private Sub Draw()

' Draws the terrain onto the form based on the values in the height map. The
' displayed colours for the pixels corresponding to values in the height map
' are based on which two height boundaries the height value is between. The
' actual colour displayed at the pixel comes from a gradient between the
' boundary colours, so that the terrain blends together and looks more
' realistic.
'
' The height map could also be used to generate 3D terrain.

Dim Color As Long

    Cls ' Clear the form before drawing
    
    ' Iterate through all points in the height map
    For y = 0 To ROW - 1
        For x = 0 To ROW - 1
            
            ' Check the height value against each boundary to see where it lies
            For i = 0 To UBound(Bound)
                If Heights(x, y) <= Bound(i) Then Exit For
            Next

            If i = UBound(Bound) + 1 Then
                Color = RGB(R(i), G(i), B(i))
            ElseIf i = 0 Then
                Color = RGB(R(0), G(0), B(0))
            Else
                ' Calculate the gradient between the two boundary colors
                Color = RGB(R(i) + (R(i + 1) - R(i)) * ((Heights(x, y) - Bound(i - 1)) / (Bound(i) - Bound(i - 1))), G(i) + (G(i + 1) - G(i)) * ((Heights(x, y) - Bound(i - 1)) / (Bound(i) - Bound(i - 1))), B(i) + (B(i + 1) - B(i)) * ((Heights(x, y) - Bound(i - 1)) / (Bound(i) - Bound(i - 1))))
            End If
            
            SetPixelV hdc, ScaleWidth \ 2 - ROW \ 2 + x, ScaleHeight \ 2 - ROW \ 2 + y, Color
            
        Next
    Next
        
End Sub
