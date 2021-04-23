Attribute VB_Name = "mdl_Translation"
Option Explicit

Public strLabel(7) As String
Public strScreentip(5) As String
Public strSupertip(5) As String
Public strError(5) As String
Public strfrmInfo(0) As String

Public Sub germanText()
    strLabel(0) = "inoRound Runden"
    strLabel(1) = "Runden"
    strLabel(2) = "Aufrunden"
    strLabel(3) = "Abrunden"
    strLabel(4) = "Zahlen runden"
    strLabel(5) = "Info Menu"
    strLabel(6) = "Info"
    strLabel(7) = "Rundung entfernen"
    
    strSupertip(0) = "Markierte Zahl(en) werden kaufmännsich gerundet."
    strSupertip(1) = "Markierte Zahl(en) werden aufgerundet."
    strSupertip(2) = "Markierte Zahl(en) werden abgerundet."
    strSupertip(3) = "Wenn in einer Zelle eine Zahl steht, wird diese gerundet."
    strSupertip(4) = "Fügen sie ein Zahlbeispiel mit der Anzahl der benötigten Stellen vor oder hinter dem Komma ein."
    strSupertip(4) = "Fügen sie ein Zahlbeispiel mit der Anzahl der benötigten Stellen vor oder hinter dem Komma ein."
    strSupertip(5) = "Enfernt vorhandene Rundungen. Die Formatierung muss ggf. angepasst werden."
    
    strScreentip(0) = "Kaufmännsich runden"
    strScreentip(1) = "Aufrunden"
    strScreentip(2) = "Abrunden"
    strScreentip(3) = "Zahlen runden"
    strScreentip(4) = "Anzahl der Rundungsstellen auswählen"
    strScreentip(5) = "Rundung entfernen"
    
    strError(0) = "Eingabehinweis"
    strError(1) = "Der eingegeben Text ist keine Zahl."
    strError(2) = "Bearbeitungshinweis"
    strError(3) = "Es dauert ein bisschen."
    strError(4) = "Bearbeite Eintrag "
    strError(5) = " von "
    
    strfrmInfo(0) = "Der Quellcode is OpenSource unter AGPLv3 und verfügbar auf "
End Sub

Public Sub englishText()
    strLabel(0) = "inoRound Round"
    strLabel(1) = "Round"
    strLabel(2) = "Round up"
    strLabel(3) = "Round down"
    strLabel(4) = "Round numbers"
    strLabel(5) = "Info Menu"
    strLabel(6) = "Info"
    strLabel(7) = "Remove round"
    
    strSupertip(0) = "Use flexible round for all marked cells."
    strSupertip(1) = "Use round up for all marked cells."
    strSupertip(2) = "Use round down for all marked cells."
    strSupertip(3) = "If a cell holds a numeric value it the number will be rounded."
    strSupertip(4) = "Add a formatted example with the amount of digits before or after the decimal separator."
    strSupertip(5) = "Remomes round. May be the formatting need to be adjusted."
    
    strScreentip(0) = "Flexible round"
    strScreentip(1) = "Round up"
    strScreentip(2) = "Round down"
    strScreentip(3) = "Round numbers"
    strScreentip(4) = "Define number of digits to get rounded"
    strScreentip(5) = "Remove round"
    
    strError(0) = "Eingabehinweis"
    strError(1) = "Der eingegeben Text ist keine Zahl."
    strError(2) = "Editing note"
    strError(3) = "The process will take a while"
    strError(4) = "Process entry "
    strError(5) = " of "
    
    strfrmInfo(0) = "Source code is OpenSource under AGPLv3 and available at "
End Sub

