Attribute VB_Name = "ModuloExcel"
Option Explicit

' Declara las rutinas API necesarias:
Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
                    ByVal lpWindowName As Long) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
                    ByVal wParam As Long, _
                    ByVal lParam As Long) As Long

Public Function GetExcel() As Byte
Dim Ind As Integer
Dim I As Integer
Dim aux As String

' Prueba para ver si hay una copia de Microsoft Excel ejecut�ndose.
    On Error Resume Next    ' Inicializa la interceptaci�n del error.
' La llamada a la funci�n Getobject sin el primer argumento devuelve una
' referencia a una instancia de la aplicaci�n . Si no se est� ejecutando,
' se produce un error . Observe que se utiliza la coma como el primer marcador del
' argumento.
    Set MiXL = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        ExcelNoSeEjecutaba = True
        Set MiXL = CreateObject("Excel.Application")
    End If
    
Err.Clear   ' Borra el objeto Err si se produce un error.


' Comprueba Excel. Si se est� ejecutando Excel,
' lo introduce en la tabla Running Object .
'    DetectExcel


' Establece la variable de objeto para hacer referencia al archivo que desea ver.
   'Dim xl As Excel.Application
    

    MiXL.Workbooks.Open (NombreHoja)
    aux = NombreHoja
    Do
        I = InStr(1, aux, "\")
        If I > 0 Then aux = Mid(aux, I + 1)
    Loop Until I = 0
    Ind = 1
    For I = 1 To MiXL.Workbooks.Count
        If MiXL.Workbooks(I).Name = aux Then
            Ind = I
            Exit For
        End If
    Next I
    
    
    Set wrk = MiXL.Workbooks.Item(Ind)
    
   
    
    
    Set ExcelSheet = wrk.Sheets(1)
    ExcelSheet.Activate

    
    
If Err.Number <> 0 Then
    GetExcel = 1
    Else
        GetExcel = 0
End If

End Function

Sub DetectExcel()
' El procedimiento dectecta un Excel en ejecuci�n y lo registra.
    Const WM_USER = 1024
    Dim hWnd As Long
' Si se est� ejecutando Excel esta llamada API devuelve el controlador .
    hWnd = FindWindow("XLMAIN", 0)
    If hWnd = 0 Then    ' 0 quiere decir que Excel no se est� ejecutando .

Exit Sub
    Else
    ' Excel se est� ejecutando por lo que se utiliza la funci�n API SendMessage
    ' para introducirlo en la tabla Running Object.
        SendMessage hWnd, WM_USER + 18, 0, 0
    End If
End Sub


Public Sub CerrarExcel()

' Si no se est� ejecutando esta copia de Microsoft Excel cuando
' comenz�, ci�rrela utilizando el m�todo Quit de la propiedad Application.
' Observe que cuando intenta salir de Microsoft Excel, la barra de t�tulo
' de Microsoft Excel parpadea y Microsoft Excel muestra un mensaje
' pregunt�ndole si desea guardar los archivos cargados.

    wrk.Save
    wrk.Close
   

    If ExcelNoSeEjecutaba Then
        MiXL.Application.Quit
    End If

    Set ExcelSheet = Nothing
    Set wrk = Nothing
    Set MiXL = Nothing  ' Libera la referencia a la
    
End Sub


Public Function AbrirEXCEL() As Byte

If GetExcel = 1 Then
    CerrarExcel
    GoTo ErrorAbrirExcel
End If
'Si queremos que se vea descomentamos  esto
MiXL.Application.Visible = True
MiXL.Parent.Windows(1).Visible = True

Exit Function
ErrorAbrirExcel:
   MsgBox "Abriendo excel: " & Err.Description, vbExclamation
End Function





Private Function ColumnaProxima(ByRef Columna As Integer) As Integer
Dim aux As Integer
Select Case Columna
Case 4, 6, 8, 10, 15
    aux = Columna + 2
Case Else
    aux = Columna + 1
End Select
ColumnaProxima = aux
End Function


Private Function ComasAPuntos(Cad As String) As String
Dim I
Do
    I = InStr(1, Cad, ",")
    If I > 0 Then _
        Cad = Mid(Cad, 1, I - 1) & "." & Mid(Cad, I + 1)
Loop Until I = 0
ComasAPuntos = Cad
End Function
