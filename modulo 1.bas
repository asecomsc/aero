Option Compare Database

Sub Imprime()
Dim Fm As Form
Set Fm = Forms("Impresion Etiquetas")
        Fm.paquete = Fm.paquete + 1  'set en form, valor predeterminado = 0
        tara = Val(Fm.piezas) * 0.03 + 0.15  '30 gr de cada cono y 150 gr de la bolsa
        miNeto = Val(Fm.pesoNeto) - tara  'bruto - tara

        Open "etiqueta.txt" For Output As #1
        Print #1, "^XA~TA080~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR4,4~SD15^JUS^LRN^CI0^XZ"
        Print #1, "^XA"
        Print #1, "^MMT"
        Print #1, "^PW812"
        Print #1, "^LL1218"
        Print #1, "^LS0"
        Print #1, "^FT155,1021^A0B,102,100^FH\^FDHilaturas Providencia^FS"
        Print #1, "^FT642,175^A0B,39,38^FH\^FD" & Fm.piezas & "^FS"
        Print #1, "^FT640,281^A0B,28,28^FH\^FDPiezas:" & "^FS"
        Print #1, "^FT643,494^A0B,39,38^FH\^FD" & Format(Val(tara), ".000") & "^FS" 'tara
        Print #1, "^FT642,836^A0B,39,38^FH\^FD" & Format(Val(Fm.pesoNeto), "0.000") & "^FS" 'peso bruto, err confund� neto con bruto
        Print #1, "^FT640,575^A0B,28,28^FH\^FDTara:^FS"
        Print #1, "^FT642,994^A0B,28,28^FH\^FDPeso Bruto:^FS"
        Print #1, "^FT565,825^A0B,45,45^FH\^FD" & Fm.paquete & "^FS"   '1 /20 '& " / " & Fm.paquetes
        Print #1, "^FT557,994^A0B,28,28^FH\^FDPaquete:^FS"
        Print #1, "^FT503,827^A0B,45,45^FH\^FD" & Fm.te�ido & "^FS"
        Print #1, "^FT497,994^A0B,28,28^FH\^FDTe\A4ido:^FS"
        Print #1, "^FT359,828^A0B,45,45^FH\^FD" & Fm.cbColor.Column(1) & "^FS"
        Print #1, "^FT347,992^A0B,28,28^FH\^FDColor:^FS"
        Print #1, "^FT440,830^A0B,56,55^FH\^FD" & Format(Val(miNeto), "0.000") & "^FS"
        Print #1, "^FT442,991^A0B,28,28^FH\^FDNeto:^FS"
        Print #1, "^FT410,992^A0B,28,28^FH\^FDPeso^FS"
        Print #1, "^FT288,836^A0B,45,45^FH\^FD" & Fm.cbModelo.Column(1) & "^FS"
        Print #1, "^FT281,994^A0B,28,28^FH\^FDModelo:^FS"
        Print #1, "^FO675,47^GB0,1110,1^FS"
        Print #1, "^FT762,994^A0B,23,24^FH\^FDTELEFONOS:      OFNA:  449  915 46 14       CEL:  449 123 83 96^FS"
        Print #1, "^FO210,65^GB0,1110,1^FS"
        Print #1, "^FT734,1072^A0B,23,24^FH\^FDPlanta Rinconada No 6-A.  La Providencia Tanque de los Jim\82nez. Aguascalientes, Ags^FS"
        Print #1, "^BY3,3,81^FT496,492^BCB,,Y,N"
        Print #1, "^FD>;" & Format(Fm.cbModelo, "000") & Format(Fm.cbColor, "000") & Format(Val(miNeto * 100), "0000") & Fm.te�ido & "^FS"
        Print #1, "^PQ1,0,1,Y^XZ"
        Close #1
        retval = Shell("cmd /c copy etiqueta.txt \\bascula-pc\zebra")
        If Val(Fm.paquete) = Val(Fm.paquetes) Then anyval = MsgBox("Impresion Terminada...", vbInformation, "Impresora Zebra")
Set ob = Nothing
End Sub
