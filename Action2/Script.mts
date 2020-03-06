' Définition de variables, notamment la variable qui contient l'adresse du fichier d'export
Set Window_MyFlight = WpfWindow("Micro Focus MyFlight Sample")
sheetname = "Global"
fichier_export = "C:\Users\recette\Documents\dtexp.xlsx"

' Supprime le fichier d'export si il existe
Set fso = createobject("Scripting.filesystemobject")
If fso.FileExists(fichier_export) = true Then
	fso.DeleteFile(fichier_export)
End If

' Importe le feuillet 1 du fichier excel dans data à l'onglet 'sheetname'
datatable.importsheet "C:\Users\recette\Documents\JDD.xlsx",1,sheetname
' Compte le nombre de ligne et le store dans n
n = datatable.GetSheet(sheetname).GetRowCount

' Boucle qui fait défiler toutes les valeurs possibles de villes de départ
For i = 1 To n Step 1
	Datatable.SetCurrentRow(i)
	Window_MyFlight.WpfComboBox("fromCity").Select datatable("From")
Next @@ hightlight id_;_2054110840_;_script infofile_;_ZIP::ssf2.xml_;_



' Sélectionne la ligne de la datatable qui contient la valeur qui nous intéresse
DataTable.SetCurrentRow(5)
' Store dans une variable la valeur issue de la datatable
cityChoice = DataTable.Value("From", "Global")
' Sélectionne la valeur dans le champ ville de départ de l'application
Window_MyFlight.WpfComboBox("fromCity").Select cityChoice

' Sélectionne la ville d'arrivée puis valide
Window_MyFlight.WpfComboBox("toCity").Select "San Francisco" @@ hightlight id_;_2054113624_;_script infofile_;_ZIP::ssf4.xml_;_
Window_MyFlight.WpfImage("WpfImage").Click 9,17 @@ hightlight id_;_2054113816_;_script infofile_;_ZIP::ssf5.xml_;_

' Clique sur l'image du calendrier puis sélection de la date
Window_MyFlight.WpfImage("WpfImage").Click
Window_MyFlight.WpfCalendar("lu").SetDate "12-Mar-2020"

Window_MyFlight.WpfComboBox("Class").Select "First" @@ hightlight id_;_2054117560_;_script infofile_;_ZIP::ssf8.xml_;_
Window_MyFlight.WpfComboBox("numOfTickets").Select "2"
Window_MyFlight.WpfButton("FIND FLIGHTS").Click @@ hightlight id_;_2064775712_;_script infofile_;_ZIP::ssf27.xml_;_

' Export de la datatable dans un fichier excel
datatable.Export "C:\Users\recette\Documents\dtexp.xlsx"
