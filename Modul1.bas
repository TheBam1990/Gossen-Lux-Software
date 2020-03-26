Attribute VB_Name = "Modul1"
Sub timer()
Application.OnTime Now() + TimeValue("00:00:01"), "ZeileAusTXTlesen2"
End Sub


Sub ZeileAusTXTlesen2()
 ' liest 1 Zeile aus einer Textdatei
 Dim intFF As Integer
 Dim strDatei As String
 
 strDatei = "C:\Daten\wert.txt"
 intFF = FreeFile
 
    Open strDatei For Input As #intFF       ' Öffnet Textdatei zum Lesen
    Line Input #intFF, strDatei             ' Liest die 1. Zeile aus
    ActiveCell = strDatei                  ' Der Wert wird aus dem Zwischenspeicher an die _
Zelle A1 übergeben
    Close #intFF                            ' schließt die Textdatei
    timer
    
 End Sub



