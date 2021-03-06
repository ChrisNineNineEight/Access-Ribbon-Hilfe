﻿Access-Ribbon-Hilfe
===================

Wenn man auf den Supertip eines Elementes eines Access Ribbons klickt oder auf "Weitere Infos" klickt, öffnet sich standardmäßig die Access Hilfe mit einem nichtssagenden Text.
Die hier erhältliche ACCDB enthält ein Modul, das sich in die Ereignisskette einklinkt, den Klick auf den Tooltip abfängt und den Namen des Elements auf dem Ribbon zurück gibt.
Mit dieser Information kann dann jeder weiter verfahren, wie er lustig ist. Zum Beispiel kann man eine eigene Hilfedatei oder auch die Hilfen auf der eigenen Webseite aufrufen.

WICHTIG:
1) DAS VBA CODE FENSTER NICHT (!!) ÖFFNEN, WÄHREND DIE EVENT HOOKS LAUFEN! DAS GIBT UNSCHÖNES GEFLACKER UND ACCESS MUSS NEU GESTARTET WERDEN! WENN IHR ETWAS AM CODE ÄNDERN WOLLT, DIE DATENBANK IMMER MIT SHIFT TASTE GEDRÜCKT STARTEN!

2) KOMPRIMIEREN UND REPARIEREN *NUR* MIT GEDRÜCKTER SHIFT-TASTE DURCHFÜHREN. SONST STÜRZT ACCESS AB!

3) ES WIRD MIT WIN32 CALLBACKS GEARBEITET. HIER AUF JEDEN FALL DAS "ON ERROR RESUME NEXT" BENUTZEN!

4) Es sind so wenige Event Hooks eingerichtet wie möglich. Das Event EVENT\_OBJECT_HIDE wird Access beim Beenden zum Beispiel definitiv abstürzen lassen. Also seid vorsichtig, was ihr alles abfragt.

5) Dieses Beispiel wurde in Access 2013 erstellt. Das Abfangen des Hilfe Buttons auf der Backstage in Access 2010 funktioniert mit diesem Beispiel nicht. Da ich kein Access 2010 besitze, kann ich dieses Problem jedoch nicht beheben.


Aufbau der ACCDB:

- AutoExec Makro zum Starten der Event Hooks
- AutoKeys Makro, um das Drücken auf F1 abzufangen, wenn die Maus sich über einem Ribbon Element befindet
- modAutoExec  zum Starten der Event Hooks (wird vom AutoExec Makro aufgerufen)
- modRibbonHelp, das den kommentierten Code enthält
- CloseForm Formular, das genutzt wird, damit man auch auf der Backstage "Beenden" und "Schließen" benutzen kann. Direktes Schließen lässt Access abstürzen.


Ablauf bei Interaktion mit der Maus:

- Benutzer bewegt die Maus über ein Element auf dem Ribbon
- Der Tooltip zu diesem Element wird geöffnet. Dadurch wird abgefragt, welches Element sich gerade unter der Maus befindet.
- Beim Klick auf den Tooltip wird abgefragt, ob es sich um einen Tooltip handelt und der Pfad des Elements wird zurück gegeben.
- RibbonHelpTooltip_Click wird aufgerufen. (HIER DANN WEITERFÜHRENDEN CODE REIN)


Ablauf bei Drücken von F1:

- Die Maus muss sich über einem Element auf dem Ribbon oder dessen Tooltip befinden.
- Beim Drücken von F1 wird der Pfad dieses Elements abgefragt und zurück gegeben.
- RibbonHelpTooltip_F1 wird aufgerufen. (HIER DANN WEITERFÜHRENDEN CODE REIN)

Klicks auf die "Hilfe" Buttons auf Front- und Backstage: (Fragezeichen oben rechts neben dem "Minimieren" Button)

- Der Button auf der Frontstage wird ganz normal über einen Callback im Ribbon XML abgefangen: <command idMso="Help" onAction="OnActionHelpButton"/>
- Der Button auf der Backstage wird so abgefangen wie auch die Tooltips der Ribbon Elemente


Probleme:

- Das Entfernen der Event Hooks funktioniert aus unerfindlichen Gründen nicht. Deswegen kann das Event EVENT\_OBJECT_HIDE nicht benutzt werden, damit man mit bekommt, wenn ein Tooltip wieder verschwindet. Wenn hier jemand eine Lösung hat, kann er mir diese gerne zukommen lassen.
