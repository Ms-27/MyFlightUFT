' Démarrage de l'application à tester FlightGUI
SystemUtil.Run "C:\Program Files (x86)\Micro Focus\Unified Functional Testing\samples\Flights Application\FlightsGUI.exe" @@ hightlight id_;_1835870_;_script infofile_;_ZIP::ssf48.xml_;_

' Utilisation de paramètre d'action pour le login et le mot de passe
''' Reste à régler le problème de typage des variables
Dim username 'As String
Dim password 'As Password
username = parameter.Item("Username")
password = parameter.Item("Password")

' Saisie du login, du mot de passe et validation
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("agentName").Set username @@ hightlight id_;_1920672944_;_script infofile_;_ZIP::ssf50.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").SetSecure password @@ hightlight id_;_1920646592_;_script infofile_;_ZIP::ssf51.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click @@ hightlight id_;_2054115496_;_script infofile_;_ZIP::ssf52.xml_;_

' Vide les variables
username = Empty
password = Empty

