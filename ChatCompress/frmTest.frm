VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dictionary Compression"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCompFlags 
      Caption         =   "Compare Flags"
      Height          =   420
      Left            =   915
      TabIndex        =   9
      Top             =   5025
      Width           =   1620
   End
   Begin VB.CommandButton cmdCompress1 
      Caption         =   "Compress"
      Height          =   450
      Left            =   6330
      TabIndex        =   8
      Top             =   4965
      Width           =   1875
   End
   Begin VB.CommandButton test 
      Caption         =   "benchmark sentence (x1000)"
      Height          =   285
      Left            =   6885
      TabIndex        =   7
      Top             =   2625
      Width           =   2520
   End
   Begin VB.CommandButton Command 
      Caption         =   "Load Dicts"
      Height          =   375
      Left            =   6720
      TabIndex        =   6
      Top             =   0
      Width           =   2655
   End
   Begin VB.TextBox txtDecomp 
      Height          =   2010
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   2895
      Width           =   9420
   End
   Begin VB.TextBox txtOutput 
      Height          =   705
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1830
      Width           =   9420
   End
   Begin VB.TextBox txtInput 
      Enabled         =   0   'False
      Height          =   1020
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   435
      Width           =   9420
   End
   Begin VB.Label lbl2 
      Caption         =   "Decompressed:"
      Height          =   285
      Left            =   210
      TabIndex        =   5
      Top             =   2610
      Width           =   2850
   End
   Begin VB.Label lblCompressed 
      Caption         =   "Compressed String:"
      Height          =   315
      Left            =   165
      TabIndex        =   4
      Top             =   1530
      Width           =   3270
   End
   Begin VB.Label lbl1 
      Caption         =   "Type here after loading dictionary"
      Height          =   270
      Left            =   150
      TabIndex        =   3
      Top             =   90
      Width           =   3645
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long
Public strCOutput As String
Public strUOutput As String

Private Sub cmdTest_Click()
txtDecomp.Text = mdc_DecompressText(txtOutput.Text)
End Sub

Private Sub cmdCompFlags_Click()
Dim i As Long
Dim f1 As String
'Dim f2 As String
'
'For i = 0 To 254
'f1 = Chr(0)
'f2 = Chr(1)
'If f1 = f2 Then MsgBox i
'MsgBox Len(f1)
'Next
For i = 1 To 11
f1 = f1 & Chr(i) & vbNewLine
Next

MsgBox f1
End Sub

Private Sub cmdCompress1_Click()
Dim sOut As String
Dim sIn As String
Dim sDecompress As String

sIn = txtInput.Text
sOut = modDictCompress.mdc_CompressText(sIn)

sDecompress = mdc_DecompressText(sOut)
End Sub

Private Sub Command_Click()
Dim i As Long
'load word list from dictionary files
mdc_loadDicts App.Path & "\dict1.txt", App.Path & "\dict2.txt", App.Path & "\dict3.txt", App.Path & "\dict4.txt", App.Path & "\dict5.txt"
txtInput.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'properly unload all created arrays, dictionaries, etc.
    UnloadCompressor
End Sub

Private Sub test_Click()
Dim i As Long
Dim t As Long
Dim t2 As Long
Dim s1 As String
Dim texxxt As String
texxxt = Me.txtInput.Text
DoEvents

t = GetTickCount
For i = 1 To 1000
s1 = mdc_CompressText(texxxt)
Next
t2 = GetTickCount

i = t2 - t

MsgBox i / 1000 & " Seconds"
End Sub

Private Sub txtInput_Change()
On Error Resume Next
strCOutput = mdc_CompressText(txtInput.Text)
Me.txtOutput.Text = strCOutput

txtDecomp.Text = mdc_DecompressText(strCOutput)

Me.Caption = "Original: " & Len(txtInput.Text) & "  Compressed: " & Len(strCOutput) & "  (Difference: " & Len(txtInput.Text) - Len(strCOutput) & " Chars)" & "  Ratio: " & (Len(strCOutput) / (Len(txtInput.Text)) * 100) & "%"
End Sub
