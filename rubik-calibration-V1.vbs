'''''''''''''''''''''''''''
' Rubik-bot by Rubikshow ''
' Tout droit réservé !   ''
'''''''''''''''''''''''''''

dim name
name = "RUBIKALI V1"
welcome = "Bienvenue sur le rubikali"
set ObVoice=CreateObject("SAPI.SpVoice")
ObVoice.speak welcome


do
	choix = inputbox ("Tape E pour calibrer l'extrudeur, Q pour quitter", name) 
		if (choix = "E") then
			msgbox("Marquez le filament a l'entree de l'extrudeur")

			msgbox("Faites chauffer votre buse et attendez que la temperature soit stable")

			msgbox("Deplacez l'axe Z a 100mm au dessus du plateau environ")

			msgbox("Executez le G-code suivant : G92 E0")

			msgbox("Executez le G-code suivant : G1 E100 F100")

			msgbox("Attendez que l'extrudeur s'arrette")

			msgbox("Marquez le filament a l'entree de l'extrudeur")

			msgbox("Sortez le filament et mesurez la distance entre les deux")

			dim mesure
			dim stepperunit
			dim result

			mesure = inputbox("Entrez la valeur obtenu en mm :" + chr(13) + "*Indiquez la separation par une virgule" + chr(13) + "*Exemple : 94,5" + chr(13) + "*Contre exemple : 94.5",name)

			stepperunit = inputbox("Quel est votre step par mm actuel sur le moteur de l'extrudeur :" + chr(13) + "*Indiquez la separation par une virgule" + chr(13) + "*Exemple : 94,5" + chr(13) + "*Contre exemple : 94.5",name)

			result = 100/mesure
			result = result*stepperunit

			msgbox("Votre nouveau step par mm est le suivant : " & result)

			wscript.quit
		end if

		if (choix = "Q") then
			wscript.quit
		end if
loop