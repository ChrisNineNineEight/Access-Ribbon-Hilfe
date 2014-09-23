Access-Ribbon-Hilfe
===================

Wenn man auf den Supertip eines Elementes eines Access Ribbons klickt oder auf "Weitere Infos" klickt, öffnet sich standardmäßig die Access Hilfe mit einem nichtssagenden Text.
Die hier erhältliche ACCDB enthält ein Modul, das sich in die Ereignisskette einklinkt, den Klick auf den Tooltip abfängt und den Namen des Elements auf dem Ribbon zurück gibt.
Mit dieser Information kann dann jeder weiter verfahren, wie er lustig ist. Zum Beispiel kann man eine eigene Hilfedatei oder auch die Hilfen auf der eigenen Webseite aufrufen.

WICHTIG:
1) DAS VBA CODE FENSTER NICHT (!!) ÖFFNEN, WÄHREND DIE EVENT HOOKS LAUFEN! DAS GIBT UNSCHÖNES GEFLACKER UND ACCESS MUSS NEU GESTARTET WERDEN! WENN IHR ETWAS AM CODE ÄNDERN WOLLT, DIE DATENBANK IMMER MIT SHIFT TASTE GEDRÜCKT STARTEN!

2) ES WIRD MIT WIN32 CALLBACKS GEARBEITET. HIER AUF JEDEN FALL DAS "ON ERROR RESUME NEXT" BENUTZEN!

3) Es sind so wenige Event Hooks eingerichtet wie möglich. Das Event EVENT\_OBJECT_HIDE wird Access beim Beenden zum Beispiel definitiv abstürzen lassen. Also seid vorsichtig, was ihr alles abfragt.


Aufbau der ACCDB:

- AutoExec Makro zum Starten der Event Hooks
- AutoKeys Makro, um das Drücken auf F1 abzufangen, wenn die Maus sich über einem Ribbon Element befindet
- modAutoExec  zum Starten der Event Hooks (wird vom AutoExec Makro aufgerufen)
- modRibbonHelp, das den auskommentierten Code enthält


Ablauf bei Interaktion mit der Maus:

- Benutzer bewegt die Maus über ein Element auf dem Ribbon
- Der Tooltip zu diesem Element wird geöffnet. Dadurch wird abgefragt, welches Element sich gerade unter der Maus befindet.
- Beim Klick auf den Tooltip wird abgefragt, ob es sich um einen Tooltip handelt und der Pfad des Elements wird zurück gegeben.
- RibbonHelpTooltip_Click wird aufgerufen. (HIER DANN WEITERFÜHRENDEN CODE REIN)


Ablauf bei Drücken von F1:

- Die Maus muss sich über einem Element auf dem Ribbon befinden.
- Beim Drücken von F1 wird der Pfad dieses Elements abgefragt und zurück gegeben.
- RibbonHelpTooltip_F1 wird aufgerufen. (HIER DANN WEITERFÜHRENDEN CODE REIN)


Probleme:

- Das Entfernen der Event Hooks funktioniert aus unerfindlichen Gründen nicht. Deswegen kann das Event EVENT\_OBJECT_HIDE nicht benutzt werden, das gebraucht wird, damit die Funktionsweise der F1 Taste so ist wie im Original. Wenn hier jemand eine Lösung hat, kann er mir diese gerne zukommen lassen.
