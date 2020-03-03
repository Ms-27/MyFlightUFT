' Défini la variable nom pour la commande
Dim name
name = "Pawan"

' Saisie du nom et validation de la commande
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").Set name @@ hightlight id_;_1920641024_;_script infofile_;_ZIP::ssf1.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Click @@ hightlight id_;_2054130904_;_script infofile_;_ZIP::ssf2.xml_;_

' Utilisation d'une regex pour le message de validation (dans le texte de l'objet dans le repository)
WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 91 completed").Click 52,23 @@ hightlight id_;_1969518864_;_script infofile_;_ZIP::ssf3.xml_;_

' Store dans deux variable les valeurs du message de validation
' d'un côté la valuer de run de l'autre la valeur de test
Dim order_ref, order_ref_test
order_ref = WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 91 completed").GetROProperty("text")
order_ref_test = WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 91 completed").GetTOProperty("text")

' Rentre les données en sortie dans un feuillet excel de notre fichier déjà existant
'' Instance d'application Excel
Set excelObj = CreateObject("Excel.Application")
'' Variable avec le chemin du fichier dans lequel on va écrire
FilePath = "C:\Users\recette\Documents\JDDexp.xlsx"
'' Défini le workbook actif
excelObj.Workbooks.Open(FilePath)
'' Détermine le feuillet dans lequel on écrit
Set ExcelSheet = excelObj.ActiveWorkbook.Worksheets("Feuil3")
'' Renseigne les valeurs dans des cellules
ExcelSheet.cells(1,1).value = name
ExcelSheet.cells(1,2).value = order_ref
'' Sauvegarde le fichier excel
excelObj.ActiveWorkbook.Save

' Vide les instances
Set excelObj = nothing
Set ExcelSheet = nothing


' Ecrit dans un fichier plat
'' Instance de Filesystemobject
Set fso = createobject("Scripting.FileSystemObject")
'' Création du fichier plat, écrase le fichier si il existe déjà
Set stream = fso.CreateTextFile("C:\Users\recette\Documents\flatfile.txt", true)
	'' Ecrit dans le fichier
	stream.WriteLine("Valeur de test	Nom: " & name & " - Etat de la commande: " & order_ref_test)
	stream.WriteLine("Valeur de run		Nom: " & name & " - Etat de la commande: " & order_ref)
stream.Close

'' Vide les instances
Set stream = nothing

''' A revoir
'Set file = fso.OpenTextFile("C:\Users\recette\Documents\flatfile.txt", ForReading, true)
'Do while file.AtEndofStrean <> True
'data = file.ReadLine()
'msgbox data
'Loop
'
Set fso = nothing
