Attribute VB_Name = "libTraza"
Option Explicit


Public mConfig As CFGControl
Public Conn As Connection


Public MiXL As Object  ' Variable que contiene la referencia
    ' de Microsoft Excel.
Public ExcelNoSeEjecutaba As Boolean   ' Indicador para liberación final .
Public ExcelSheet As Object
Public wrk As Excel.Workbook

Public BaseDatos As String

Public EsImportaci As Byte
Public NombreHoja As String

Public Const ValorNulo = "Null"
Public Const FormatoFecha = "yyyy-mm-dd"
Public Const FormatoHora = "hh:mm:ss"
Public Const FormatoImporte = "#,###,###,##0.00"
Public Const FormatoPrecio = "##,##0.000"
Public Const FormatoPorcen = "##0.00"

Dim Rc As Byte


Public Usuario As Long


Public Sub Main()
    Dim I As Integer
    'Vemos si ya se esta ejecutando
    If App.PrevInstance Then
        MsgBox "Ya se está ejecutando el programa de Vocaldo (Tenga paciencia).", vbCritical
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    
    Set mConfig = New CFGControl
    If mConfig.Leer = 1 Then
        MsgBox "No configurado"
        End
    End If
    
    If AbrirConexion(mConfig.BaseDatos) Then
        frmTraza.Show vbModal
    End If

End Sub

Public Function RecuperaValor(ByRef Cadena As String, Orden As Integer) As String
Dim I As Integer
Dim j As Integer
Dim cont As Integer
Dim cad As String

    I = 0
    cont = 1
    cad = ""
    Do
        j = I + 1
        I = InStr(j, Cadena, "|")
        If I > 0 Then
            If cont = Orden Then
                cad = Mid(Cadena, j, I - j)
                I = Len(Cadena) 'Para salir del bucle
                Else
                    cont = cont + 1
            End If
        End If
    Loop Until I = 0
    RecuperaValor = cad
End Function

Public Function AbrirConexion(BaseDatos As String) As Boolean
Dim cad As String

    
    AbrirConexion = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
                       
    cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & Trim(BaseDatos) & ";SERVER=" & mConfig.SERVER & ";"
    cad = cad & ";UID=" & mConfig.User
    cad = cad & ";PWD=" & mConfig.password
'++monica: tema del vista
    cad = cad & ";Persist Security Info=true"
    
    Conn.ConnectionString = cad
    Conn.Open
    If Err.Number <> 0 Then
        MsgBox "Error en la cadena de conexion" & vbCrLf & BaseDatos, vbCritical
        End
    Else
        AbrirConexion = True
    End If
End Function


Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
'Para cuando recupera Datos de la BD
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    DBLet = ""
                Case "N"    'Numero
                    DBLet = 0
                Case "F"    'Fecha
                     '==David
'                    DBLet = "0:00:00"
                     '==Laura
'                     DBLet = "0000-00-00"
                      DBLet = ""
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function

Public Function DBSet(vData As Variant, Tipo As String, Optional EsNulo As String) As Variant
'Establece el valor del dato correcto antes de Insertar en la BD
Dim cad As String

        If IsNull(vData) Then
            DBSet = ValorNulo
            Exit Function
        End If

        If Tipo <> "" Then
            Select Case Tipo
                Case "T"    'Texto
                    If vData = "" Then
                        If EsNulo = "N" Then
                            DBSet = "''"
                        Else
                            DBSet = ValorNulo
                        End If
                    Else
                        cad = (CStr(vData))
                        NombreSQL cad
                        DBSet = "'" & cad & "'"
                    End If
                    
                Case "N"    'Numero
                    If vData = "" Or vData = 0 Then
                        If EsNulo <> "" Then
                            If EsNulo = "S" Then
                                DBSet = ValorNulo
                            Else
                                DBSet = 0
                            End If
                        Else
                            DBSet = 0
                        End If
                    Else
                        cad = CStr(ImporteFormateado(CStr(vData)))
                        DBSet = TransformaComasPuntos(cad)
                    End If
                    
                Case "F"    'Fecha
'                     '==David
''                    DBLet = "0:00:00"
'                     '==Laura
                    If vData = "" Then
                        If EsNulo = "S" Then
                            DBSet = ValorNulo
                        Else
                            DBSet = "'1900-01-01'"
                        End If
                    Else
                        DBSet = "'" & Format(vData, FormatoFecha) & "'"
                    End If
                    
                Case "FH" 'Fecha/Hora
                    If vData = "" Then
                        If EsNulo = "S" Then DBSet = ValorNulo
                    Else
                        DBSet = "'" & Format(vData, "yyyy-mm-dd hh:mm:ss") & "'"
                    End If
                    
                Case "H" 'Hora
                    If vData = "" Then
                    Else
                        DBSet = "'" & Format(vData, "hh:mm:ss") & "'"
                    End If
                    
                Case "B"  'Boolean
                    If vData Then
                        DBSet = 1
                    Else
                        DBSet = 0
                    End If
            End Select
        End If
End Function


'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef Cadena As String)
Dim j As Integer
Dim I As Integer
Dim Aux As String
    j = 1
    Do
        I = InStr(j, Cadena, "'")
        If I > 0 Then
            Aux = Mid(Cadena, 1, I - 1) & "\"
            Cadena = Aux & Mid(Cadena, I)
            j = I + 2
        End If
    Loop Until I = 0
End Sub

'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256,98
'   Tiene que venir numérico
Public Function ImporteFormateado(Importe As String) As Currency
Dim I As Integer

    If Importe = "" Then
        ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            I = InStr(1, Importe, ".")
            If I > 0 Then Importe = Mid(Importe, 1, I - 1) & Mid(Importe, I + 1)
        Loop Until I = 0
        ImporteFormateado = Importe
    End If
End Function

'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(Cadena As String) As String
Dim I As Integer
    Do
        I = InStr(1, Cadena, ",")
        If I > 0 Then
            Cadena = Mid(Cadena, 1, I - 1) & "." & Mid(Cadena, I + 1)
        End If
    Loop Until I = 0
    TransformaComasPuntos = Cadena
End Function

Public Function Round2(Number As Variant, Optional NumDigitsAfterDecimals As Long) As Variant
Dim Ent As Integer
Dim cad As String

  ' Comprobaciones

  If Not IsNumeric(Number) Then
    Err.Raise 13, "Round2", "Error de tipo. Ha de ser un número."
    Exit Function
  End If

  If NumDigitsAfterDecimals < 0 Then
    Err.Raise 0, "Round2", "NumDigitsAfterDecimals no puede ser negativo."
    Exit Function
  End If

  ' Redondeo.

  cad = "0"
  If NumDigitsAfterDecimals <> 0 Then cad = cad & "." & String(NumDigitsAfterDecimals, "0")
  Round2 = Val(TransformaComasPuntos(Format(Number, cad)))

End Function

Public Sub PonerFoco(ByRef Text As TextBox)
On Error Resume Next
    Text.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Public Function TotalRegistros(vSQL As String) As Long
'Devuelve el valor de la SQL
'para obtener COUNT(*) de la tabla
Dim RS As ADODB.Recordset

    On Error Resume Next

    Set RS = New ADODB.Recordset
    RS.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    TotalRegistros = 0
    If Not RS.EOF Then
        If RS.Fields(0).Value > 0 Then TotalRegistros = RS.Fields(0).Value  'Solo es para saber que hay registros que mostrar
    End If
    RS.Close
    Set RS = Nothing

    If Err.Number <> 0 Then
        TotalRegistros = 0
        Err.Clear
    End If
End Function

'''
'''Mezcala una linea del fichero de texto sobre el fichero de EXCEL
'''Public Function MezclaFicheros(ValorLinea As String, NumeroLinea As Integer) As Byte
'''Dim Col As Integer
'''Dim aux As String
'''Dim j As Integer
'''Dim inicio As Integer
'''Dim Cadena As String
'''
'''
'''On Error GoTo ErrorMezcla
'''MezclaFicheros = 1
'''inicio = 1
'''Col = 2  'Porque empieza en la 2
'''Do
'''    j = InStr(inicio, ValorLinea, "|")
'''    If j > 0 Then
'''        Cadena = Mid(ValorLinea, inicio, j - inicio)
'''        aux = ComasAPuntos(Cadena)
'''        ExcelSheet.Cells(NumeroLinea, Col) = aux
'''        inicio = j + 1
'''        Col = ColumnaProxima(Col)
'''    End If
'''Loop Until j = 0
'''
'''MezclaFicheros = 0
'''Exit Function
'''ErrorMezcla:
'''
'''End Function
