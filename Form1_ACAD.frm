VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Tag + Value"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3975
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   225
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read the ATTRIBUTES from AutoCAD"
      Height          =   555
      Left            =   45
      TabIndex        =   0
      Top             =   4680
      Width           =   5685
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Extract()
    Form1.Cls
    Dim sheet As Object
    Dim shapes As Object
    Dim elem As Object
    Dim excel As Object
    Dim Max As Integer
    Dim Min As Integer
    Dim NoOfIndices As Integer
    Dim excelSheet As Object
    Dim RowNum As Integer
    Dim Array1 As Variant
    Dim Count As Integer
    Dim Teller As Integer
    Dim Teller1 As Integer
    
    Screen.MousePointer = vbHourglass
    procOpenDrawing
    Set Doc = acad.ActiveDocument
    Set mspace = Doc.ModelSpace
    RowNum = 1
    Dim Header As Boolean
    Header = False
    Teller = 0
    Teller1 = 0
    Text1.Text = ""
    For Each elem In mspace
        With elem
            If StrComp(.EntityName, "AcDbBlockReference", 1) = 0 Then
                If .HasAttributes Then
                    Teller = Teller + 1
                    Array1 = .GetAttributes
                     Text1.Text = Text1.Text & vbNewLine & " ****** Read the TAGS ***** " & vbNewLine
                    For Count = LBound(Array1) To UBound(Array1)
                        If Header = False Then
                            If StrComp(Array1(Count).EntityName, "AcDbAttribute", 1) = 0 Then
                                 Text1.Text = Text1.Text & Array1(Count).TagString & vbNewLine
                            End If
                        End If
                    Debug.Print
                    Next Count
                    RowNum = RowNum + 1
                    
                    Text1.Text = Text1.Text & vbNewLine & " ****** Read the VALUE ***** " & vbNewLine
                    
                    For Count = LBound(Array1) To UBound(Array1)
                        Teller1 = Teller1 + 1
                        If Count = 0 Then
                            Text1.Text = Text1.Text & Array1(Count).TextString & vbNewLine
                        Else
                             Text1.Text = Text1.Text & Array1(Count).TextString & vbNewLine
                        End If
                    Next Count
                    Debug.Print
                    Header = True
                End If
            End If
        End With
    Next elem
    NumberOfAttributes = RowNum - 1
    If NumberOfAttributes > 0 Then
    
    Else
        MsgBox "No attributes found in the current drawing"
    End If
    Set acad = Nothing
    Me.SetFocus
    Screen.MousePointer = vbNormal
End Sub

Private Sub Auto_Close()
    Set excelSheet = Nothing
End Sub

Private Sub Command1_Click()
    Extract
End Sub

Sub procOpenDrawing()
        Set acad = Nothing
        On Error Resume Next
        Set acad = GetObject(, "AutoCAD.Application")
        If Err <> 0 Then
            Set acad = CreateObject("AutoCAD.Application") 'If AutoCAD is closed...
        End If
        acad.Visible = True
        Set Doc = acad.ActiveDocument
        Set mspace = Doc.ModelSpace
        dwgName = App.Path & "\Drawing.dwg"
        If Dir(dwgName) <> "" Then
            Doc.Open dwgName
        Else
            MsgBox "File: " & dwgName & vbCrLf & vbCrLf & " can't find the file...", vbInformation, "Message..."
        End If
        Unload frmAutocad

End Sub

