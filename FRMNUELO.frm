VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FRMNUELO 
   Caption         =   "PROPIEDADES DE USUARIO"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "ELIMINAR"
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
      TabIndex        =   15
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDITAR"
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
      Left            =   3240
      TabIndex        =   14
      Top             =   5160
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2760
      Top             =   6240
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\JULIO\Desktop\PROYECTOFINAL\DATA\BASEINV.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\JULIO\Desktop\PROYECTOFINAL\DATA\BASEINV.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "DUENO"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   11
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   10
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GUARDAR"
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
      Left            =   1680
      TabIndex        =   8
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton CMDCREAR 
      Caption         =   "CREAR"
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
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton CMDREG 
      Caption         =   "REGRESAR"
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
      Left            =   6840
      TabIndex        =   6
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox TXTCEDN 
      DataField       =   "CEDULA"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   1
      Top             =   2520
      Width           =   3615
   End
   Begin VB.TextBox TXTNOMN 
      DataField       =   "NOMBRE"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3600
      TabIndex        =   0
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Line Line4 
      X1              =   6480
      X2              =   6480
      Y1              =   5040
      Y2              =   6000
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   8640
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIOS REGISTRADOS"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   13
      Top             =   600
      Width           =   3615
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   8640
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   8760
      Y1              =   3720
      Y2              =   3720
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
      Left            =   840
      TabIndex        =   5
      Top             =   360
      Width           =   3735
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   10800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10800
      Y1              =   1200
      Y2              =   1200
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
      Left            =   840
      TabIndex        =   4
      Top             =   2640
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
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
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
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   2415
   End
End
Attribute VB_Name = "FRMNUELO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CON As New ADODB.Connection
Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1

Private Sub CMDCREAR_Click()
Adodc1.Recordset.AddNew

End Sub

Private Sub CMDREG_Click()
If MsgBox("Esta seguro que desea regresar al formulario de login?", vbQuestion + vbYesNo) = vbYes Then
FRMNUELO.Hide
FRMLOGIN.Show
    End If
End Sub

Private Sub Command1_Click()
If TXTNOMN.Text = "" Or TXTCEDN.Text = "" Then
    
    MsgBox "Llenar todos los campos del nuevo usuario.", vbInformation, "Dialogo"
    Adodc1.Recordset.Delete
    Else
    MsgBox "El nuevo usuario ha sido registrado.", vbInformation, "Dialogo"
    End If
End Sub

Private Sub Command2_Click()
     
Adodc1.Recordset.Fields("NOMBRE") = TXTNOMN.Text
Adodc1.Recordset.Fields("CEDULA") = TXTCEDN.Text
Adodc1.Recordset.Update
MsgBox "El registro ha sido actualizado.", vbInformation, "Dialogo"
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveLast

End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MovePrevious

End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MoveFirst

End Sub

Private Sub Command7_Click()
If MsgBox("Esta seguro que desea eliminar este registro del usuario?", vbQuestion + vbYesNo) = vbYes Then
        Adodc1.Recordset.Delete

    End If
End Sub

Private Sub Form_Load()
Dim CN As New ADODB.Connection
Dim rs As New ADODB.Recordset
Adodc1.LockType = adLockReadOnly
rs.LockType = adLockOptimistic

End Sub
