VERSION 5.00
Begin VB.Form FRMLOGIN 
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDSALIR 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton CMDLOGIN 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox TXTNOM 
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox TXTCED 
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   2
      Top             =   3600
      Width           =   4455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CEDULA"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CEDULA"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   960
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10800
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   10800
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BAZAR Jessica"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[!] INGRESE EL NOMBRE Y NUMERO DE CEDULA DEL PROPIETARIO"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   8895
   End
End
Attribute VB_Name = "FRMLOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDLOGIN_Click()
    If TXTNOM = "Jessica" And TXTCED = "1234567890" Then
        FRMLOGIN.Hide
        FRMINV.Show
    ElseIf TXTNOM = "" And TXTCED = "" Then
    MsgBox "Llenar todos los campos indicados.", vbInformation, "Dialogo"
    ElseIf TXTNOM = "" Then
    MsgBox "Llenar el campo de nombre", vbInformation, "Dialogo"
    ElseIf TXTCED = "" Then
    MsgBox "Llenar el campo de cedula", vbInformation, "Dialogo"
    ElseIf Not (IsNumeric(TXTCED.Text)) Then
    MsgBox "Llenar el campo de cedula correcta con numeros", vbInformation, "Dialogo"
    TXTCED = ""
    Else
    MsgBox "Un campo ingresado es incorrecto", vbCritical, "Dialogo"
    TXTNOM = ""
    TXTCED = ""
    End If
End Sub

Private Sub CMDSALIR_Click()
End
End Sub

Private Sub Form_Load()

End Sub
