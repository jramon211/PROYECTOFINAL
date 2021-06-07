VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form FRMVENTAS 
   Caption         =   "Form1"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "CERRAR INVENTARIO"
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
      Left            =   10680
      TabIndex        =   21
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MOSTRAR INVENTARIO"
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
      Left            =   10680
      TabIndex        =   20
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LIMPIAR"
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
      Left            =   7440
      TabIndex        =   19
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox TXTID 
      DataField       =   "IDPRODUCTOS"
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
      Left            =   9240
      TabIndex        =   18
      Top             =   5400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   720
      Top             =   8040
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "VENTAS"
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
   Begin VB.CommandButton CMDREG 
      Caption         =   "REGRESAR"
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
      Left            =   10680
      TabIndex        =   17
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox TXTCAN2 
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
      Left            =   3480
      TabIndex        =   15
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton CMDB1 
      Caption         =   "BUSCAR ID"
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
      Left            =   7440
      TabIndex        =   14
      Top             =   1920
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "VENTAS.frx":0000
      Height          =   1575
      Left            =   840
      TabIndex        =   13
      Top             =   6120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2778
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
   Begin VB.TextBox TXTTOT 
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
      Left            =   3480
      TabIndex        =   11
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox TXTIDPRO 
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
      Left            =   3480
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox TXTCAN 
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
      Left            =   3480
      TabIndex        =   3
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox TXTCOS 
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
      Left            =   3480
      TabIndex        =   2
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton CMDAGR 
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
      Left            =   7440
      TabIndex        =   1
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton CMDELI 
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
      Left            =   7440
      TabIndex        =   0
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Line Line3 
      X1              =   10440
      X2              =   10440
      Y1              =   120
      Y2              =   7680
   End
   Begin VB.Label Label7 
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
      Left            =   1200
      TabIndex        =   16
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
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
      Left            =   1200
      TabIndex        =   12
      Top             =   4800
      Width           =   1815
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
      TabIndex        =   10
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ventas"
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
      TabIndex        =   9
      Top             =   120
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ID PRODUCTO"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK"
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
      Left            =   1200
      TabIndex        =   7
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COSTO/u"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   10320
      Y1              =   1080
      Y2              =   1080
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
      Left            =   2040
      TabIndex        =   5
      Top             =   1560
      Width           =   4215
   End
End
Attribute VB_Name = "FRMVENTAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim CN As New ADODB.Connection
 Private WithEvents rs As ADODB.Recordset
Attribute rs.VB_VarHelpID = -1
Private Sub CMDB1_Click()
TXTCAN2.Enabled = True

rs.Requery
rs.Find "idproducto=" & Val(TXTIDPRO.Text)

If rs.EOF Then
            MsgBox "No se encontro ningun registro", vbInformation, "Eliminar registro"
            Exit Sub 'Termina el procedimiento
        Else
           TXTCAN.Text = rs.Fields("CANTIDAD")
           TXTCOS.Text = rs.Fields("PRECIO")
           TXTID.Text = Val(TXTIDPRO.Text)
            Command1.Enabled = True
            CMDAGR.Enabled = True
            CMDELI.Enabled = True
End If
 
End Sub

Private Sub CMDAGR_Click()

If Val(TXTCAN2.Text) > Val(TXTCAN.Text) Then
    MsgBox "El cantidad de productos pedidos es mayor al stock en este momento.", vbInformation, "Dialogo"
    Exit Sub
Else
    rs.Fields("CANTIDAD") = Val(TXTCAN.Text) - Val(TXTCAN2.Text)
    rs.Update
End If



TXTTOT.Text = Val(TXTCAN2.Text) * Val(TXTCOS.Text)
Adodc1.Recordset.AddNew

'Le añadi estas lineas, como te dije al momento de poner un AddNew debo especificar los campos y con que informacion voy _
a llenarlos. El problema de porque no nos salio antes es que en tu proyecto tienes como datasource de los textbox el adodc _
por lo que al momento de poner AddNew los textbox borran su contenido y ya no podemos extraer la infomacion de ahi, solo le quite eso _
y le añadi estas lineas.

Adodc1.Recordset("CEDULA") = (a)
Adodc1.Recordset("IDPRODUCTOS") = (TXTIDPRO.Text)
Adodc1.Recordset("CANTIDAD") = (TXTCAN2.Text)
Adodc1.Recordset("TOTAL") = (TXTTOT.Text)
Adodc1.Recordset("FECHA") = (Date)
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc1.Recordset.MoveLast

FRMINV.Adodc1.Refresh
TXTCAN.Text = Val(TXTCAN.Text) - Val(TXTCAN2.Text)

End Sub

Private Sub CMDELI_Click()
If MsgBox("Esta seguro que desea eliminar un registro?", vbQuestion + vbYesNo) = vbYes Then
        rs.Fields("CANTIDAD") = Val(TXTCAN.Text) + Val(TXTCAN2.Text)
        rs.Update
        Adodc1.Recordset.Delete
        FRMINV.Adodc1.Refresh
    End If
End Sub

Private Sub CMDREG_Click()
If MsgBox("Esta seguro que desea regresar al formulario de inventario?", vbQuestion + vbYesNo) = vbYes Then
FRMVENTAS.Hide
FRMINV.Show
    End If
End Sub

Private Sub Command1_Click()
TXTIDPRO.Text = ""
TXTID.Text = ""
TXTCAN.Text = ""
TXTCOS.Text = ""
TXTCAN2.Text = ""
TXTTOT.Text = ""
End Sub

Private Sub Command2_Click()
FRMINV.Show
End Sub

Private Sub Command3_Click()
FRMINV.Hide
End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
    'Abrimos la base de datos "agenda.mdb".
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data\BASEINV.mdb;Persist Security Info=False"

    '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\JULIO\Desktop\PROYECTOFINAL\DATA\BASEINV.mdb;Persist Security Info=False"
    rs.Source = "INVENTARIO" 'Especificamos la fuente de datos. En este caso la tabla "contactos".
    rs.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    rs.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
    rs.Open "select * from INVENTARIO", CN 'Abrimos el Recordset y lo llenamos con una consulta SQL.
    'Cargamos los datos en las cajas de texto.
    rs.MoveFirst 'Nos movemos al principio del Recordset.
    TXTCAN.Enabled = False
    TXTCOS.Enabled = False
    TXTCAN2.Enabled = False
    TXTTOT.Enabled = False
     Command1.Enabled = False
      CMDAGR.Enabled = False
    CMDELI.Enabled = False
    
End Sub

