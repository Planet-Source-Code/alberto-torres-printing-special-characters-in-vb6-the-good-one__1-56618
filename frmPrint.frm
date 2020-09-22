VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrint 
   AutoRedraw      =   -1  'True
   Caption         =   "Printing"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdlPrint 
      Left            =   1665
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStop 
      Height          =   690
      Left            =   2182
      Picture         =   "frmPrint.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   675
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   690
      Left            =   877
      Picture         =   "frmPrint.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   675
      Width           =   735
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Font_Name As String = "Arial"
Private MyString    As String
Public i            As Long

Private Sub cmdPrint_Click()

        MyString = "abcdefghijklmnopqrstuvwxzABCDEFGHIJKLMNOPQRSTUVWXYZ" & vbCrLf
  
        On Error GoTo ErrorHandler

'       Stablishing the CommomDialog Control
        With cdlPrint
            .CancelError = True
            On Error Resume Next
            
            If (Err.Number = cdlCancel) Or (Err.Number = 32755) Or (cdlCancel = True) Then
            '   The User Canceled. Do nothing.
                On Error GoTo 0
                Exit Sub
            ElseIf Err.Number <> 0 Then
            '   Unexpected error. Report it.
                GoTo ErrorHandler
                Exit Sub
            Else
                .Flags = (cdlPDReturnDC = True) And (cdlPDSelection = True) And _
                         (cdlPDHidePrintToFile = True) And _
                         (cdlPDDisablePrintToFile = True) And _
                         (cdlPDSelection = True) And _
                         (cdlPDAllPages = True)
                
                .ShowPrinter
                Printer.Orientation = .Orientation
            End If
        End With
    
    '  Printing in Printer
         
        Printer.Font.Bold = True
        Printer.Print vbTab & "Alberto Torres Klinger"
        Printer.Print vbTab & "Mechanical Engineer" & vbCrLf
        Printer.Print vbTab & "Allowable Stress Design" & vbCrLf
        Printer.Print vbTab & Me.Caption & vbCrLf
        Printer.Print vbTab & Format(Now, " mmm dd, yyyy  (dddd)" & vbTab & "hh:mm:ss") & vbCrLf
        Printer.Font.Bold = False
        '
       
        Printer.Print "Alberto"
        sSymbol Chr$(&HD6)  '   Alt + 251
        sSymbol Chr$(&HF2)  '   S
        For i = &H40 To &H7E
            sSymbol Chr$(i)
            Printer.Print " ";
        Next i
        
        sSymbol MyString
        Printer.Print vbCrLf
        
        'Printer.Font.Name = "GreekC"
        sGreekC "Abc"
        sGreekC MyString
        Printer.Print vbCrLf
        
        'Printer.Font.Name = "GreekS"
        sGreekS "Abc"
        sGreekS MyString
        Printer.Print vbCrLf
        
        sGreekC "s"
        Printer.Print " = ";
        sSymbol Chr$(&HFD)
        sGreekC " f"
        Printer.Print " ";
        sSymbol Chr$(&HFD)
        
        
        Printer.NewPage
        Printer.EndDoc
        Printer.KillDoc
           
        Exit Sub
ErrorHandler:
        MsgBox "Error: " & Format$(Err.Number) & _
               "Selecting Printer:" & Printer.DeviceName & vbCrLf & vbCrLf & _
               Err.Description & vbCrLf & vbCrLf & _
               Err.Source, vbCritical, Me.Caption
        Resume Next
        Exit Sub

End Sub

Private Sub cmdStop_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Show
    AutoRedraw = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure     : sSymbol
' DateTime      : 10/7/2004 10:32
' Author        : Alberto Torres Klinger
'
' Description   : To Printer a single character or a long string in "Symbol".
' Inputs        : a string.
' Outputs       : a string in "Symbol".
'
' Comments      :
'---------------------------------------------------------------------------------------

Public Sub sSymbol(char As String)
    
    With Printer
        .Font.Name = "Symbol"
        .Font.Bold = True
    
        Printer.Print char;
        .Font.Bold = False
    End With
    
    Printer.Font.Name = Font_Name
    Printer.Print "";
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure     : sGreekS
' DateTime      : 10/7/2004 11:00
' Author        : Alberto Torres Klinger
'
' Description   : To Printer a single character or a long string in "GreekS".
' Inputs        : a string.
' Outputs       : a string in GreekC.
'
' Comments      :
'---------------------------------------------------------------------------------------

Public Sub sGreekS(char As String)
    
    With Printer
        .Font.Name = "GreekS"
        .Font.Bold = True
    
        Printer.Print char;
        .Font.Bold = False
    End With
    
    Printer.Font.Name = Font_Name
    Printer.Print "";
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure     : sGreekC
' DateTime      : 10/7/2004 11:00
' Author        : Alberto Torres Klinger
'
' Description   : To Printer a single character or a long string in "GreekC".
' Inputs        : a string.
' Outputs       : a string in GreekC.
'
' Comments      :
'---------------------------------------------------------------------------------------

Public Sub sGreekC(char As String)
    
    With Printer
        .Font.Name = "GreekC"
        .Font.Bold = True
    
        Printer.Print char;
        .Font.Bold = False
    End With
    
    Printer.Font.Name = Font_Name
    Printer.Print "";
    
End Sub
