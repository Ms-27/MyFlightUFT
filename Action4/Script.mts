' Défini la variable nom pour la commande
Dim name
name = "Pawan"

' Saisie du nom et validation de la commande
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").Set name @@ hightlight id_;_1920641024_;_script infofile_;_ZIP::ssf1.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Click @@ hightlight id_;_2054130904_;_script infofile_;_ZIP::ssf2.xml_;_

' Utilisation d'une regex pour le message de validation (dans le texte de l'objet dans le repository)
WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 91 completed").Click 52,23 @@ hightlight id_;_1969518864_;_script infofile_;_ZIP::ssf3.xml_;_

' Store dans deux variable les valeurs du message de validation
' d'un côté la valeur de run, de l'autre la valeur de test
Dim order_ref, order_ref_test
order_ref = WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 91 completed").GetROProperty("text")
order_ref_test = WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 91 completed").GetTOProperty("text")



' Rentre les données en sortie dans une feuille excel d'un fichier

'' Variable avec le chemin du fichier dans lequel on va écrire
Dim FilePath
FilePath = "C:\Users\recette\Documents\Orderexp.xlsx"

'' Instancie un objet de système de fichier
Set fso = CreateObject("Scripting.FileSystemObject")
'' Vérifie l'éxistence du fichier, si non création
If (fso.FileExists(FilePath)) Then
	'' Set une instance excel
	Set excelObj = CreateObject("Excel.Application")
	excelObj.Visible = true
	'' Ouvre le fichier
	excelObj.Workbooks.Open FilePath
	
	'' Instancie un objet feuille excel
	Set resultSheetObj = excelObj.ActiveWorkbook.Worksheets(1)
	
Else
	'' Set une instance excel
	Set excelObj = CreateObject("Excel.Application")
	excelObj.Visible = true
	'' Ajoute un workbook
	excelObj.Workbooks.Add
	'' Sauvegarde le fichier
	excelObj.ActiveWorkbook.SaveAs FilePath
	
	'' Instancie un objet feuille excel
	Set resultSheetObj = excelObj.ActiveWorkbook.Worksheets(1)
	'' Nomme la feuille
	resultSheetObj.Name = "Order"
End If

'' Renseigne les valeurs dans des cellules
resultSheetObj.cells(1,1).value = name
resultSheetObj.cells(1,2).value = order_ref

'' Sauvegarde le fichier Excel et quit
excelObj.ActiveWorkbook.Save
excelObj.Workbooks.Close
excelObj.Quit

' Vide les instances
Set excelObj = nothing
Set resultSheetObj = nothing



' Ecrit dans un fichier plat
'' Défini une variable qui contient le chemin du fichier
Dim FlatFilePath
FlatFilePath = "C:\Users\recette\Documents\flatfile.txt"
'' Instance de Filesystemobject
Set fso = createobject("Scripting.FileSystemObject")
'' Création du fichier plat, écrase le fichier si il existe déjà
Set stream = fso.CreateTextFile(FlatFilePath, true)
	'' Ecrit dans le fichier
	stream.WriteLine("Valeur de test	Nom: " & name & " - Etat de la commande: " & order_ref_test)
	stream.WriteLine("Valeur de run		Nom: " & name & " - Etat de la commande: " & order_ref)
	stream.Close

' Vide les instances
Set stream = nothing
Set fso = nothing



' Lecture d'un fichier plat
'' Instance de Filesystemobject
Set fso = CreateObject("Scripting.FileSystemObject")

'' Instancie l'objet fichier plat
Set flatFile = fso.OpenTextFile(FlatFilePath, 1, True)

'' Boucle qui renvoit chaque ligne puis son numéro de ligne dans le Output
Do Until flatFile.AtEndOfStream
	print flatFile.ReadLine & " - line: " & flatFile.Line
Loop
'' Fermeture du fichier
flatFile.Close

' Vide les instances et variables
Set fso = nothing
Set flatFile = nothing
FlatFilePath = Empty
