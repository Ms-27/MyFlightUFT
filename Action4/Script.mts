' Saisie du nom et validation de la commande
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").Set "Manjoya" @@ hightlight id_;_1920641024_;_script infofile_;_ZIP::ssf1.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Click @@ hightlight id_;_2054130904_;_script infofile_;_ZIP::ssf2.xml_;_

' Utilisation d'une regex pour le message de validation (dans le texte de l'objet dans le repository)
WpfWindow("Micro Focus MyFlight Sample").WpfObject("Order 91 completed").Click 52,23 @@ hightlight id_;_1969518864_;_script infofile_;_ZIP::ssf3.xml_;_
