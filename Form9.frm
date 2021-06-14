VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12360
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   6465
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   360
      Top             =   840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   360
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton Command7 
      Caption         =   "Ver Facturas False"
      Height          =   495
      Left            =   7920
      TabIndex        =   10
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Buscar Fatura"
      Height          =   495
      Left            =   10200
      TabIndex        =   5
      Top             =   5760
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar Factura"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7920
      TabIndex        =   3
      Top             =   3000
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   1320
      TabIndex        =   2
      Top             =   1920
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4683
      _Version        =   393216
      Enabled         =   -1  'True
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
            LCID            =   12298
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
            LCID            =   12298
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
   Begin VB.CommandButton Command2 
      Caption         =   "Añadir producto"
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver Productos"
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Image Image3 
      Height          =   840
      Left            =   8520
      Picture         =   "Form9.frx":4305A
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   10560
      Picture         =   "Form9.frx":4A160
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   6360
      Picture         =   "Form9.frx":4B6C9
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   960
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "F"
      Height          =   375
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Show
    Form1.inicio
    Form1.Label8.Caption = 1
    Form9.Hide
End Sub

Private Sub Command2_Click()
    Form6.Show
    Form6.Command3.Visible = True
    Form6.Command2.Enabled = False
    Form6.Image1.Picture = LoadPicture(App.Path & "\img\No_Picture.jpg")
    Form9.Hide
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    CFact
    With Fact
        .Find "Id_F='" & Label1.Caption & "'"
        !Valido = "False"
        .UpdateBatch
    End With
    CDFact
        With DFact
        If .State = 1 Then .Close
        .Open "select * from Detalle_Factura where [Id_F]like '" & Label1.Caption & "'", base, adOpenStatic, adLockBatchOptimistic
            For i = 1 To .RecordCount
                If .EOF Or .BOF Then Exit Sub
                a = !Id_P_FK
                b = !Talla
                c = !Cantidad
                CTP
                With TP
                    .Find "Id_Producto='" & a & "'"
                    If b = "S" Then !Talla_S = Val(!Talla_S) + Val(c)
                    If b = "M" Then !Talla_M = Val(!Talla_M) + Val(c)
                    If b = "G" Then !Talla_G = Val(!Talla_G) + Val(c)
                    .UpdateBatch
                End With
                .MoveNext
            Next i
        End With
        Label5.Caption = "T"
    If Label2.Caption = "F" Then carga3
    If Label2.Caption = "T" Then carga2
End Sub

Private Sub Command4_Click()
    Label2.Caption = "F"
    Form10.Show
    Form10.Text1.Text = ""
End Sub

Private Sub Command5_Click()
    Form7.Show
    Form9.Hide
End Sub

Private Sub Command7_Click()
    Set DataGrid1.DataSource = Nothing
    If Command7.Caption = "Ver Facturas False" Then
        Adodc2.CursorLocation = adUseClient
        Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\base\base.mdb;Persist Security Info=False"
        X = "False"
        Adodc2.RecordSource = "select * from Factura where [Valido]like '" & X & "'"
        Set DataGrid1.DataSource = Adodc2
        Command7.Caption = "Ver Facturas True"
    Else
        carga3
        Command7.Caption = "Ver Facturas False"
    End If
End Sub

Private Sub DataGrid1_Click()
    If DataGrid1.ApproxCount < 1 Then Exit Sub
    Command3.Enabled = True
    If Label2.Caption = "T" Then
        With Fact
            Label1.Caption = !Id_F
        End With
    Else
        If Label5.Caption <> "T" Then
            With Adodc1.Recordset
                Label1.Caption = !Id_F
            End With
        Else
            With Fact
                Label1.Caption = !Id_F
            End With
        End If
    End If
End Sub

Sub carga()
    Adodc1.CursorLocation = adUseClient
    Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\base\base.mdb;Persist Security Info=False"
    X = "True"
    Adodc1.RecordSource = "select * from Factura where [Valido]like '" & X & "'"
    Set DataGrid1.DataSource = Adodc1
End Sub

Sub carga2()
    With Fact
        If .State = 1 Then .Close
        X = Label3.Caption
        Y = "True"
        .Open "select * from Factura where [Id_C]like '" & X & "' and [Valido]like '" & Y & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
    Set DataGrid1.DataSource = Fact
End Sub

Sub carga3()
    With Fact
        If .State = 1 Then .Close
        Y = "True"
        .Open "select * from Factura where [Valido]like '" & Y & "'", base, adOpenStatic, adLockBatchOptimistic
    End With
    Set DataGrid1.DataSource = Fact
End Sub

Private Sub Form_Load()
    carga
End Sub

Private Sub Image1_Click()

End Sub
