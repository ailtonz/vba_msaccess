Attribute VB_Name = "Bas_BarCode_Generator"
Option Compare Database   'Use database order for string comparisons
Option Explicit

' (c) 1993-1995 James I. Mercanti, MicroDoctor, Baton Rouge, LA.  USA
'
' Permission granted for public use and royalty-free distribution only
' if my Copyright notice is included with this module code for credit.
' No other mention of source or credits is required. All rights reserved.
'
' I ask for no money, no contributions, just credit for my work.
'
' I'm working on more symbologies including 2of5, 128, PDF-417, etc.
'
' Send me an e-note with your postal address for demos and more info.
' email: microdoc@ix.netcom.com
'
'
' NOTE FOR PURISTS: I have elected to use standard English object
' identifiers for those not familiar with L-R naming conventions.
'
'
' TO USE THIS CODE:
'
'   1 - Create Report with a TextBox control. (example named Barcode)
'       Make sure the Visible property is set to "No".
'   2 - Set On-Print property of section to [Event Procedure]
'       by clicking on the [...] and selecting "Code Builder"
'   3 - Confirm that the following code matches yours...
'
'      Sub Detail1_Print (Cancel As Integer, PrintCount As Integer)
'
'         Result = MD_Barcode39(Barcode, Me)
'
'      End Sub
'
'   4 - NOTE: The name of the section is "Detail1" for example only!
'       Your section might show a different name. Ditto for "Barcode".
'
'   5 - NOTE: To use on sub-forms, the Report name should be hard-coded
'       into the function. i.e. Rpt = Reports!MainForm!SubForm.Report.
'       The easy method is to just avoid using sub-forms and sub-reports.

Function MD_Barcode39(Ctrl As Control, Rpt As Report)
    
    On Error GoTo ErrorTrap_BarCode39
    
    Dim Nbar As Single, Wbar As Single, Qbar As Single, NextBar As Single
    Dim CountX As Single, CountY As Single, CountR As Single
    Dim Parts As Single, Pix As Single, Color As Long, BarStamp As Variant
    Dim Stripes As String, OneStripe As String, Barcode As String
    Dim Mx As Single, my As Single, Sx As Single, Sy As Single
    Const White = 16777215: Const Black = 0
    Const Nratio = 20, Wratio = 55, Qratio = 35
    Sx = Ctrl.Left: Sy = Ctrl.Top: Mx = Ctrl.Width: my = Ctrl.Height
    Barcode = Ctrl
    Parts = (Len(Barcode) + 2) * ((6 * Nratio) + (3 * Wratio) + (1 * Qratio))
    Pix = (Mx / Parts):
    Nbar = (20 * Pix): Wbar = (55 * Pix): Qbar = (35 * Pix)
    NextBar = Sx
    Color = White
    BarStamp = "*" & UCase(Barcode) & "*"
    For CountX = 1 To Len(BarStamp)
        Stripes = MD_BC39(Mid$(BarStamp, CountX, 1))
        For CountY = 1 To 9
            OneStripe = Mid$(Stripes, CountY, 1)
            If Color = White Then Color = Black Else Color = White
            Select Case OneStripe
                Case "1"
                    Rpt.Line (NextBar, Sy)-Step(Wbar, my), Color, BF
                    NextBar = NextBar + Wbar    'WideBar
                Case "0"
                    Rpt.Line (NextBar, Sy)-Step(Nbar, my), Color, BF
                    NextBar = NextBar + Nbar    'NarrowBar
            End Select
        Next CountY
        If Color = White Then Color = Black Else Color = White
        Rpt.Line (NextBar, Sy)-Step(Qbar, my), Color, BF
        NextBar = NextBar + Qbar     'Intermediate Quiet Bar
    Next CountX
    
Exit_BarCode39:
    Exit Function

ErrorTrap_BarCode39:
    Resume Exit_BarCode39

End Function

Function MD_BC39(CharCode As String) As String
    
    On Error GoTo ErrorTrap_BC39

    ReDim BC39(90)

    BC39(32) = "011000100" ' space
    BC39(36) = "010101000" ' $
    BC39(37) = "000101010" ' %
    BC39(42) = "010010100" ' * Start/Stop
    BC39(43) = "010001010" ' +
    BC39(45) = "010000101" ' |
    BC39(46) = "110000100" ' .
    BC39(47) = "010100010" ' /
    BC39(48) = "000110100" ' 0
    BC39(49) = "100100001" ' 1
    BC39(50) = "001100001" ' 2
    BC39(51) = "101100000" ' 3
    BC39(52) = "000110001" ' 4
    BC39(53) = "100110000" ' 5
    BC39(54) = "001110000" ' 6
    BC39(55) = "000100101" ' 7
    BC39(56) = "100100100" ' 8
    BC39(57) = "001100100" ' 9
    BC39(65) = "100001001" ' A
    BC39(66) = "001001001" ' B
    BC39(67) = "101001000" ' C
    BC39(68) = "000011001" ' D
    BC39(69) = "100011000" ' E
    BC39(70) = "001011000" ' F
    BC39(71) = "000001101" ' G
    BC39(72) = "100001100" ' H
    BC39(73) = "001001100" ' I
    BC39(74) = "000011100" ' J
    BC39(75) = "100000011" ' K
    BC39(76) = "001000011" ' L
    BC39(77) = "101000010" ' M
    BC39(78) = "000010011" ' N
    BC39(79) = "100010010" ' O
    BC39(80) = "001010010" ' P
    BC39(81) = "000000111" ' Q
    BC39(82) = "100000110" ' R
    BC39(83) = "001000110" ' S
    BC39(84) = "000010110" ' T
    BC39(85) = "110000001" ' U
    BC39(86) = "011000001" ' V
    BC39(87) = "111000000" ' W
    BC39(88) = "010010001" ' X
    BC39(89) = "110010000" ' Y
    BC39(90) = "011010000" ' Z
    
    MD_BC39 = BC39(Asc(CharCode))

Exit_BC39:
    Exit Function

ErrorTrap_BC39:
    MD_BC39 = ""
    Resume Exit_BC39

End Function

