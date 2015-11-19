VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTraza 
   Caption         =   "Generacion de volcado"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16215
   Icon            =   "frmITraza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   16215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameEscribir 
      BorderStyle     =   0  'None
      Height          =   8655
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   16185
      Begin VB.Frame Frame1 
         Height          =   585
         Left            =   11520
         TabIndex        =   9
         Top             =   150
         Width           =   2565
         Begin VB.OptionButton Option2 
            Caption         =   "Ver todo"
            Height          =   255
            Left            =   1200
            TabIndex        =   11
            Top             =   210
            Width           =   1035
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Hoy "
            Height          =   255
            Left            =   210
            TabIndex        =   10
            Top             =   210
            Width           =   1035
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   9210
         TabIndex        =   7
         Text            =   "Combo1"
         Top             =   270
         Width           =   1875
      End
      Begin VB.CommandButton cmdAccCRM 
         Height          =   495
         Index           =   2
         Left            =   15270
         Picture         =   "frmITraza.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Eliminar"
         Top             =   240
         Width           =   585
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6615
         Left            =   270
         TabIndex        =   5
         Top             =   990
         Width           =   15615
         _ExtentX        =   27543
         _ExtentY        =   11668
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   2190
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   270
         Width           =   6045
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   13560
         TabIndex        =   1
         Top             =   7890
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   14760
         TabIndex        =   2
         Top             =   7890
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Línea"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   8550
         TabIndex        =   8
         Top             =   300
         Width           =   1785
      End
      Begin VB.Label Label2 
         Caption         =   "Lectura Código EAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   330
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmTraza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Private WithEvents frmC As frmCal
Private NoEncontrados As String




Dim Sql As String
Dim VariasEntradas As String


Dim Albaran As Long
Dim FecAlbaran As String
Dim Socio As String
Dim Campo As String
Dim Variedad As String
Dim TipoEntr As String
Dim KilosNet As String
Dim Calidad(20) As String
Dim TotalArray As Integer

Private WithEvents frmMens As frmMensajes 'Registros que no ha entrado con error
Attribute frmMens.VB_VarHelpID = -1


Private Function EliminarVolcado()
Dim Sql As String

    On Error GoTo eEliminarVolcado

    EliminarVolcado = False

    Sql = "delete from trzlineas_cargas where idpalet = " & DBSet(Me.ListView1.SelectedItem.Text, "N")
    Conn.Execute Sql
    
    EliminarVolcado = True
    Exit Function
    
eEliminarVolcado:
    MsgBox "Error en eliminar Volcado: " & vbCrLf & vbCrLf & Err.Description, vbExclamation
End Function

Private Sub cmdAccCRM_Click(Index As Integer)

    If Me.ListView1.SelectedItem Then
        If MsgBox("¿ Desea eliminar el volcado seleccionado ?" & vbCrLf & "Id.Palet: " & Me.ListView1.SelectedItem.Text, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
            If EliminarVolcado Then
                MsgBox "Proceso realizado correctamente.", vbExclamation
                CargarListaVolcados
            End If
        End If
        PonerFoco Text2
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Rc As Byte
Dim Mens As String
Dim Palet As String

    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
        
    If DatosOk Then
        If InsertarVolcado(Palet) Then
            MsgBox "Se ha realizado el volcado del IdPalet " & Palet & " con éxito.", vbExclamation
            Text2.Text = ""
            Combo1.ListIndex = 0
            CargarListaVolcados
        End If
    End If
    
    PonerFoco Text2
        
        
End Sub

Private Function InsertarVolcado(Palet As String) As Boolean
Dim Sql As String
Dim Rs As ADODB.Recordset
    On Error GoTo eInsertarVolcado

    InsertarVolcado = False

    Sql = "select * from trzpalets where crfid = " & DBSet(Trim(Text2.Text), "T")

    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not Rs.EOF Then
        Sql = "insert into trzlineas_cargas(linea, idpalet, fechahora, fecha, tipo) values (" & Combo1.ListIndex + 1 & ","
        Sql = Sql & DBSet(Rs!Idpalet, "N") & ", " & DBSet(Now, "FH") & "," & DBSet(Now, "F") & ",0)"
        
        Conn.Execute Sql
        
        Palet = Rs!Idpalet
    End If
    
    InsertarVolcado = True
    Exit Function

eInsertarVolcado:
    MsgBox "Error insertando Volcado: " & vbCrLf & vbCrLf & Err.Description, vbExclamation

End Function


Private Function DatosOk() As Boolean
Dim Sql As String
Dim b As Boolean

    DatosOk = False
    b = True
    
    If Text2.Text = "" Then
        MsgBox "Debe leer de la pistola. Revise.", vbExclamation
        PonerFoco Text2
        b = False
    Else
        ' comprobamos que no se haya volcado ya
        Sql = "select count(*) from trzlineas_cargas inner join trzpalets on trzlineas_cargas.idpalet = trzpalets.idpalet "
        Sql = Sql & " where trzpalets.crfid = " & DBSet(Trim(Text2.Text), "T")
    
        If TotalRegistros(Sql) <> 0 Then
            MsgBox "Este palet ya ha sido volcado. Revise", vbExclamation
            PonerFoco Text2
            b = False
        End If
        
        If b Then
            Sql = "select count(*) from trzpalets where crfid = " & DBSet(Trim(Text2.Text), "T")
            If TotalRegistros(Sql) = 0 Then
                MsgBox "Esta etiqueta no ha sido asignada a ningún palet.", vbExclamation
                PonerFoco Text2
                b = False
            End If
        End If
        ' debemos introducir la linea de volcado
        If b Then
            If Combo1.ListIndex = -1 Then
                MsgBox "Debe seleccionar una línea de volcado. Revise.", vbExclamation
                Combo1.SetFocus
                b = False
            End If
        End If
        
    End If
    DatosOk = b
End Function





Private Sub Form_Activate()
    Combo1.ListIndex = 0
    
    Option1.Value = True
End Sub

Private Sub Form_Load()
    
    FrameEscribir.visible = False
    Limpiar
    
    Caption = "Generación Volcado de Palet"
    FrameEscribir.visible = True
 

 
 
    CargaCombo
 
    CargarCabecera
 
    CargarListaVolcados
    

End Sub

Private Sub CargarCabecera()
    
    ListView1.ColumnHeaders.Clear
    
    ListView1.ColumnHeaders.Add , , "IdPalet", 1000.0631
    ListView1.ColumnHeaders.Add , , "Fecha", 1500.2522, 2
    ListView1.ColumnHeaders.Add , , "Hora", 1000.2522, 2
    ListView1.ColumnHeaders.Add , , "Código", 1100.2522, 2
    ListView1.ColumnHeaders.Add , , "Socio", 3600.2522
    ListView1.ColumnHeaders.Add , , "Código", 1100.2522, 2
    ListView1.ColumnHeaders.Add , , "Variedad", 2200.2522
    ListView1.ColumnHeaders.Add , , "Albarán", 1200.2522, 2
    ListView1.ColumnHeaders.Add , , "Cajas", 1400.2522, 1
    ListView1.ColumnHeaders.Add , , "Kilos", 1400.2522, 1

End Sub

Private Sub CargarListaVolcados()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim It As ListItem

    Sql = "select trzlineas_cargas.idpalet, date(fechahora) fecha, time(fechahora) hora, trzpalets.codsocio, rsocios.nomsocio, trzpalets.codvarie, variedades.nomvarie, "
    Sql = Sql & " trzpalets.numnotac, trzpalets.numcajones, trzpalets.numkilos "
    Sql = Sql & " from trzlineas_cargas, trzpalets, rsocios, variedades "
    Sql = Sql & " where trzlineas_cargas.idpalet = trzpalets.idpalet "
    Sql = Sql & " and trzpalets.codsocio = rsocios.codsocio "
    Sql = Sql & " and trzpalets.codvarie = variedades.codvarie "
    
    If Option1.Value Then
        'Sql = Sql & " and date(fechahora) >= date_sub(curdate(), interval 6 day) "
        Sql = Sql & " and date(fechahora) = " & DBSet(Now, "F")
    End If
    
    Sql = Sql & " order by 2 desc, 3 desc "
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    
    ListView1.ListItems.Clear
    TotalArray = 0
    While Not Rs.EOF
        Set It = ListView1.ListItems.Add
            
        It.Text = Format(DBLet(Rs!Idpalet, "N"), "0000000")
        It.SubItems(1) = DBLet(Rs!fecha, "F")
        It.SubItems(2) = Format(Rs!hora, "hh:mm:ss")
        It.SubItems(3) = Format(DBLet(Rs!codsocio, "N"), "000000")
        It.SubItems(4) = DBLet(Rs!nomsocio, "T")
        It.SubItems(5) = Format(DBLet(Rs!codvarie, "N"), "000000")
        It.SubItems(6) = DBLet(Rs!nomvarie, "T")
        It.SubItems(7) = Format(DBLet(Rs!numnotac, "N"), "0000000")
        It.SubItems(8) = Format(DBLet(Rs!numcajones, "N"), "###,##0")
        It.SubItems(9) = Format(DBLet(Rs!numkilos, "N"), "###,##0")
        
        Rs.MoveNext
        TotalArray = TotalArray + 1
        If TotalArray > 300 Then
            TotalArray = 0
            DoEvents
        End If
    Wend
    Rs.Close
    
End Sub





Private Sub Limpiar()
Dim T As Control
    For Each T In Me.Controls
        If TypeOf T Is TextBox Then
            T.Text = ""
        End If
    Next
        
End Sub

Private Function TransformaComasPuntos(Cadena) As String
Dim cad As String
Dim j As Integer
    
    j = InStr(1, Cadena, ",")
    If j > 0 Then
        cad = Mid(Cadena, 1, j - 1) & "." & Mid(Cadena, j + 1)
    Else
        cad = Cadena
    End If
    TransformaComasPuntos = cad
End Function

Private Sub frmC_Selec(vFecha As Date)
'    Text4.Text = Format(vFecha, "dd/mm/yyyy")
End Sub



Public Sub IncrementarProgresNew(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
'    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    PBar.Value = PBar.Value + Veces
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Option1_Click()
    CargarListaVolcados
End Sub

Private Sub Option2_Click()
    CargarListaVolcados
End Sub



Private Sub Text2_GotFocus()
    ConseguirFoco Text2, 3
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub KEYpress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
    
    End If
End Sub

Private Sub CargaCombo()
Dim Ini As Integer
Dim Fin As Integer
Dim i As Integer
Dim Rs As ADODB.Recordset
Dim Sql As String


    ' *** neteje els combos, els pose valor i seleccione el valor per defecte ***
    Combo1.Clear
    
    'tipo de hectareas
    Combo1.AddItem "Línea 1"
    Combo1.ItemData(Combo1.NewIndex) = 0
    Combo1.AddItem "Línea 2"
    Combo1.ItemData(Combo1.NewIndex) = 1
    Combo1.AddItem "Línea 3"
    Combo1.ItemData(Combo1.NewIndex) = 2
    
End Sub

Public Sub ConseguirFoco(ByRef Text As TextBox, Modo As Byte)
'Acciones que se realizan en el evento:GotFocus de los TextBox:Text1
'en los formularios de Mantenimiento
On Error Resume Next

    If (Modo <> 0 And Modo <> 2) Then
        If Modo = 1 Then 'Modo 1: Busqueda
            Text.BackColor = vbYellow
        End If
        Text.SelStart = 0
        Text.SelLength = Len(Text.Text)
    End If
    If Err.Number <> 0 Then Err.Clear
End Sub

