VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Addition"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3075
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add!"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox tbxNmb2 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      ToolTipText     =   "Second operand"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox tbxNmb1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      ToolTipText     =   "First operand"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblNmb2 
      Caption         =   "Second operand"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblNmb1 
      Caption         =   "First operando"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Author:  José Antonio Barranquero Fernández
'Version: 1.0
'Date:    02/04/2017
'Remark:  Simple application adds two numbers

Dim num1 As Integer
Dim num2 As Integer
Dim res As Integer

'Subroutine (event) that handles the Add button click event
Private Sub btnAdd_Click()
On Error Resume Next
	Let num1 = CInt(tbxNmb1.Text)
	Let num2 = CInt(tbxNmb2.Text)
	Let res = num1 + num2
	MsgBox "Result is: " & Str(res), vbOKOnly, "Result"
End Sub

'Subroutine (event) that handles the Clear button click event
Private Sub btnClear_Click()
	tbxNmb1.Text = ""
	tbxNmb2.Text = ""
	Let num1 = 0
	Let num2 = 0
	Let res = 0
End Sub
