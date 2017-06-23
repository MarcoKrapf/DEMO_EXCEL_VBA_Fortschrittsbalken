# DEMO_EXCEL_VBA_Fortschrittsbalken
## Demo-Code eines Fortschrittbalkens in Excel VBA

Option Explicit

'Beschreibung
'------------
    'Diese Demo zeigt, wie ein Fortschrittsbalken
    'selbst erstellt und angewendet werden kann.
    
    'Für ein Debugging im Einzelschrittmodus (F8)
    'sollte die Variable "lngAnzahlSchleifendurchlaeufe"
    'auf einen niedrigen Wert gesetzt werden.

'Vorbereitung
'------------
    'UserForm anlegen (hier "frmFortschrittsbalken") mit den Elementen,
    'die im Code angesprochen werden (hier ein Label mit dem Namen
    '"lblFortschrittBalken" ohne Text, aber mit einer Hintergrundfarbe
    'und ein Label mit dem Namen "lblFortschrittProzent" zur Anzeige
    'des Fortschritts in Prozent als Text.
    'Wichtig: Die "ShowModal"-Eigenschaft des UserForms muss auf "False" stehen
    
'Code
'----
Sub DemoFortschrittsbalken()
    
    'Variablen
    Dim lngAnzahlSchleifendurchlaeufe As Long 'Wie oft die Schleife durchlaufen werden soll
    Dim i As Long 'Zählvariable für die Schleife
    Dim intBalkenLaenge As Integer 'Länge, die der Fortschrittsbalken am Ende haben soll
    Dim dblBalkenAnteil As Double 'Breite des Fortschrittsbalkens pro Schleifendurchlauf in Pixel
    Dim dblBalkenAktuell As Double 'Aktuelle Breite des Fortschrittsbalkens in Pixel
    
    'UserForm mit dem Fortschrittsbalken anzeigen
    frmFortschrittsbalken.Show
    
    'Stückelung des Balkens berechnen
        'In dieser Demo wird der Wert für die Anzahl der Schleifendurchläufe
        'statisch auf 50000 gesetzt. In der Realität wird der Wert dynamisch
        'je nach Aktion berechnet, die in einer Schlaufe ausgeführt wird
        '(z.B. Anzahl der Zeilen, die durchlaufen werden).
        lngAnzahlSchleifendurchlaeufe = 50000 'Für Debugging mit F8 z.B. auf 50 setzen
        'Die Zahl 200 ist die Breite in Pixel, die der Balken in dieser
        'Demo am Ende erreichen soll.
        intBalkenLaenge = 200 'Kann je nach Wunsch z.B. auf 100 gesetzt werden
        
        'Berechnung der Breite in Pixel, die der Balken pro
        'Schleifendurchlauf länger wird.
        dblBalkenAnteil = intBalkenLaenge / lngAnzahlSchleifendurchlaeufe

    'Schleife durchlaufen und Fortschrittsbalken verlängern
    For i = 1 To lngAnzahlSchleifendurchlaeufe
    
        'Neue Balkenlänge berechnen
        dblBalkenAktuell = dblBalkenAktuell + dblBalkenAnteil
        
        'Fortschrittsbalken aktualisieren
        With frmFortschrittsbalken
            'Aktuelle Breite des Balkens (mit CInt in ganze Pixel konvertiert)
            .lblFortschrittBalken.Width = CInt(dblBalkenAktuell)
            'Anzeige des Prozentanteils (mit CInt in ganze Prozentzahl
            'konvertiert, CInt kann zum Debugging auch entfernt werden)
            .lblFortschrittProzent.Caption = CInt(100 / intBalkenLaenge * dblBalkenAktuell)
        End With
        
        'Fortschrittsbalken neu zeichnen
        DoEvents 'WICHTIG!
        
    Next i
    
End Sub
