Option Compare Database
Const Patron = "Net"
Dim Buffer As String

Private Sub Form_Load()
    NETComm.CommPort = 6
    NETComm.Settings = "9600,n,8,1"
    NETComm.RThreshold = 1
    If NETComm.PortOpen = False Then NETComm.PortOpen = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    NETComm.PortOpen = False
End Sub

Private Sub NETComm_OnComm()
    If NETComm.CommEvent = NETComm_EV_RECEIVE Then
        Buffer = Buffer & NETComm.InputData
        If InStr(Buffer, Patron) Then
            ProcessData Buffer
            Buffer = ""
        End If
    End If
End Sub

Sub ProcessData(Buffer As String)
On Error GoTo handler
    ini = InStr(Buffer, Patron)
    Forms![Impresion Etiquetas].pesoNeto = Mid(Buffer, ini - 8, 4)
    'Debug.Print Buffer
    'Debug.Print Mid(Buffer, 44, 4)
    Call Imprime
Exit Sub

handler:
    MsgBox "error numero " & Err.Number
    MsgBox "Buffer: " & Buffer
End Sub

Private Sub paquetes_AfterUpdate()
    deCuantos.Caption = "/ " & paquetes
End Sub
