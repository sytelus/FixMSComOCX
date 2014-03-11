VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtReport 
      Height          =   1095
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   1920
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fix"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtDir 
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Text            =   "C:\temp1\psc\"
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Dir with slash:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    'Fix VBP files
    Dim sVBPFileName As String
    
    Const sQUOTE As String = """"
    
    sVBPFileName = Dir(txtDir.Text & "*.vbp")
    Do While sVBPFileName <> vbNullString
        Call ReplaceClassIDLine(txtDir.Text & sVBPFileName, _
            Array("MSCOMCTL.OCX", "COMCT332.OCX", "MSCOMCT2.OCX"), _
            Array("Object={831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0; MSCOMCTL.OCX", "Object={38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0; COMCT332.OCX", "Object={FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0; MSCOMCT2.OCX"))
        sVBPFileName = Dir
    Loop
    Const a = """a""bas""S"""
    sVBPFileName = Dir(txtDir.Text & "*.frm")
    Do While sVBPFileName <> vbNullString
        Call ReplaceClassIDLine(txtDir.Text & sVBPFileName, _
            Array("MSCOMCTL.OCX", "COMCT332.OCX", "MSCOMCT2.OCX"), _
            Array("Object = ""{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0""; ""MSCOMCTL.OCX""", "Object = ""{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0""; ""COMCT332.OCX""", "Object = ""{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0""; ""MSCOMCT2.OCX"""))
        sVBPFileName = Dir
    Loop

    sVBPFileName = Dir(txtDir.Text & "*.ctl")
    Do While sVBPFileName <> vbNullString
        Call ReplaceClassIDLine(txtDir.Text & sVBPFileName, _
            Array("MSCOMCTL.OCX", "COMCT332.OCX", "MSCOMCT2.OCX"), _
            Array("Object = ""{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0""; ""MSCOMCTL.OCX""", "Object = ""{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0""; ""COMCT332.OCX""", "Object = ""{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#2.0#0""; ""MSCOMCT2.OCX"""))
        sVBPFileName = Dir
    Loop

    txtReport.Text = txtReport.Text & "Done" & vbCrLf
    
End Sub

Private Sub ReplaceClassIDLine(ByVal vsFileName As String, ByVal vaSearchString As Variant, ByVal vaReplacementLine As Variant)
    Dim sFileContent As String
    Dim sLine As String
    Dim lFileHandle As Integer
    Dim lFindStringIndex As Long
    
    lFileHandle = FreeFile
    Open vsFileName For Input Access Read As #lFileHandle
    Do While Not EOF(lFileHandle)
        Line Input #lFileHandle, sLine
        For lFindStringIndex = LBound(vaSearchString) To UBound(vaReplacementLine)
            If InStr(1, sLine, vaSearchString(lFindStringIndex), VbCompareMethod.vbTextCompare) <> 0 Then
                sLine = vaReplacementLine(lFindStringIndex)
                txtReport.Text = txtReport.Text & vsFileName & vbCrLf
            End If
        Next lFindStringIndex
        sFileContent = sFileContent & sLine & vbCrLf
    Loop
    Close #lFileHandle
    
    lFileHandle = FreeFile
    Open vsFileName For Binary Access Write As #lFileHandle
        Put #lFileHandle, , sFileContent
    Close #lFileHandle
End Sub
