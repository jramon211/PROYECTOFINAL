VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FRMINV 
   Caption         =   "INVENTARIO"
   ClientHeight    =   12855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11835
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   ScaleHeight     =   12855
   ScaleWidth      =   11835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMDGUARD 
      Caption         =   "GUARDAR"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   22
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "VENTAS"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   21
      Top             =   11880
      Width           =   2295
   End
   Begin VB.TextBox TXTNUMP 
      DataField       =   "NOMBRE"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      Top             =   2520
      Width           =   3615
   End
   Begin VB.TextBox TXTCAN 
      DataField       =   "CANTIDAD"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox TXTCOS 
      DataField       =   "PRECIO"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton CMDMOD 
      Caption         =   "AGREGAR"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton CMDGUA 
      Caption         =   "EDITAR"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   12
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton CMDBUS 
      Caption         =   "ELIMINAR"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      TabIndex        =   11
      Top             =   4440
      Width           =   1695
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
      Height          =   735
      Left            =   9720
      TabIndex        =   10
      Top             =   11880
      Width           =   735
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
      Height          =   735
      Left            =   8760
      TabIndex        =   9
      Top             =   11880
      Width           =   735
   End
   Begin VB.CommandButton Command2 
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
      Height          =   735
      Left            =   7800
      TabIndex        =   8
      Top             =   11880
      Width           =   735
   End
   Begin VB.CommandButton Command1 
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
      Height          =   735
      Left            =   6960
      TabIndex        =   7
      Top             =   11880
      Width           =   735
   End
   Begin VB.CommandButton CMDBUSCAR 
      Caption         =   "BUSCAR"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8880
      TabIndex        =   6
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton CMDSAL 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9600
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   6255
      Left            =   240
      TabIndex        =   4
      Top             =   5400
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11033
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   11520
      Top             =   12000
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
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
      RecordSource    =   "INVENTARIO"
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
   Begin VB.TextBox TXTIDPRO 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   3
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Line Line6 
      X1              =   6600
      X2              =   6600
      Y1              =   11760
      Y2              =   12600
   End
   Begin VB.Line Line5 
      X1              =   6480
      X2              =   6480
      Y1              =   11760
      Y2              =   12600
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
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COSTO"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   18
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "DATOS DE LOS PRODUCTOS"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   840
      TabIndex        =   17
      Top             =   2040
      Width           =   4215
   End
   Begin VB.Line Line4 
      X1              =   8040
      X2              =   8040
      Y1              =   1920
      Y2              =   5040
   End
   Begin VB.Line Line3 
      X1              =   7920
      X2              =   7920
      Y1              =   1920
      Y2              =   5040
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   10920
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IDPRODUCTOS"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   2
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10920
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "Myriad Hebrew"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      TabIndex        =   1
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
End
Attribute VB_Name = "FRMINV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDBUS_Click()
If MsgBox("Esta seguro que desea eliminar un registro?", vbQuestion + vbYesNo) = vbYes Then
        Adodc1.Recordset.Delete

    End If
End Sub

Private Sub CMDBUSCAR_Click()
Adodc1.Refresh
DataGrid1.Refresh
Adodc1.Recordset.Find "idproducto=" & Val(TXTIDPRO.Text)

End Sub

Private Sub CMDGUA_Click()
     
Adodc1.Recordset.Fields("NOMBRE") = TXTNUMP.Text
Adodc1.Recordset.Fields("CANTIDAD") = TXTCAN.Text
Adodc1.Recordset.Fields("PRECIO") = TXTCOS.Text
Adodc1.Recordset.Update
MsgBox "El registro ha sido actualizado.", vbInformation, "Dialogo"
End Sub

Private Sub CMDGUARD_Click()
If TXTNUMP.Text = "" Or TXTCAN.Text = "" Or TXTCOS.Text = "" Then
    
    MsgBox "Llenar todos los campos de datos de los productos", vbInformation, "Dialogo"
    Adodc1.Recordset.Delete
    Else
    MsgBox "El registro ha sido guardado.", vbInformation, "Dialogo"
    End If
End Sub

Private Sub CMDMOD_Click()
    Adodc1.Recordset.AddNew
    
End Sub
'
Private Sub CMDSAL_Click()
If MsgBox("Esta seguro que desea cerrar el formulario?", vbQuestion + vbYesNo) = vbYes Then
            End
    End If
End Sub
'BUTONES DE MOVIMIENTO INICIO

Adodc1.Recordset.MoveFirst

End Sub

Private Sub Command2_Click()

Adodc1.Recordset.MovePrevious

End Sub

Private Sub Command3_Click()

Adodc1.Recordset.MoveNext

End Sub

Private Sub Command4_Click()

Adodc1.Recordset.MoveLast

End Sub
'BUTONES DE MOVIMIENTO END

Private Sub Command5_Click()
FRMVENTAS.Show
FRMINV.Hide

End Sub

Private Sub DataGrid1_DblClick()
If DataGrid1.ApproxCount < 1 Then
MsgBox "no ha seleccionado ningun registro", vbExclamation
Exit Sub
Else
      TXTIDPRO.Text = DataGrid1.Columns(0).Text
     TXTNUMP.Text = DataGrid1.Columns(1).Text
     TXTCAN.Text = DataGrid1.Columns(2).Text
     TXTCOS.Text = DataGrid1.Columns(3).Text
     'TXTIDPROV.Text = DataGrid1.Columns(4).Text
    
    
End If

End Sub

Private Sub Form_Load()

Dim CN As New ADODB.Connection
Dim rs As New ADODB.Recordset
Adodc1.LockType = adLockReadOnly
rs.LockType = adLockOptimistic


End Sub
