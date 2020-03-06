' Démarrage de l'application à tester: FlightGUI
SystemUtil.Run "C:\Program Files (x86)\Micro Focus\Unified Functional Testing\samples\Flights Application\FlightsGUI.exe" @@ hightlight id_;_1835870_;_script infofile_;_ZIP::ssf48.xml_;_

' Utilisation de paramètre d'action pour le login et le mot de passe
username = parameter.Item("Username")
password = parameter.Item("Password")

' Utilise une variable pour un objet
Set Window_MyFlight = WpfWindow("Micro Focus MyFlight Sample")

' Saisie du login, du mot de passe et validation
Window_MyFlight.WpfEdit("agentName").Set username @@ hightlight id_;_2064789248_;_script infofile_;_ZIP::ssf53.xml_;_
Window_MyFlight.WpfEdit("password").SetSecure password

' Checkpoint sur le bouton OK puis clique
Set Ok_Button = Window_MyFlight.WpfButton("OK")
Ok_Button.Check CheckPoint("OK")
Ok_Button.Click @@ hightlight id_;_2064790352_;_script infofile_;_ZIP::ssf55.xml_;_

' Vide les variables
username = Empty
password = Empty
Set Window_MyFlight = nothing
Set Ok_Button = nothing
