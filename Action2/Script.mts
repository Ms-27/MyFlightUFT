' Effectue une recherche de vol en changeant les 5 paramètres
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("fromCity").Select "Paris" @@ hightlight id_;_2054110840_;_script infofile_;_ZIP::ssf2.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("toCity").Select "San Francisco" @@ hightlight id_;_2054113624_;_script infofile_;_ZIP::ssf4.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfImage("WpfImage").Click 9,17 @@ hightlight id_;_2054113816_;_script infofile_;_ZIP::ssf5.xml_;_

	' Clique sur l'image du calendrier puis sélection de la date
WpfWindow("Micro Focus MyFlight Sample").WpfImage("WpfImage").Click
WpfWindow("Micro Focus MyFlight Sample").WpfCalendar("lu").SetDate "12-Mar-2020"

WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("Class").Select "First" @@ hightlight id_;_2054117560_;_script infofile_;_ZIP::ssf8.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("numOfTickets").Select "2" @@ hightlight id_;_2054121688_;_script infofile_;_ZIP::ssf12.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click @@ hightlight id_;_2054123272_;_script infofile_;_ZIP::ssf13.xml_;_
