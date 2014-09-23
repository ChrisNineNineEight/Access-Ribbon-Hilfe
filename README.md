Access-Ribbon-Hilfe
===================

Wenn man auf den Supertip eines Elementes eines Access Ribbons klickt oder auf "Weitere Infos" klickt, öffnet sich standardmäßig die Access Hilfe mit einem nichtssagenden Text.
Die hier erhältliche ACCDB enthält ein Modul, das sich in die Ereignisskette einklinkt, den Klick auf den Tooltip abfängt und den Namen des Elements auf dem Ribbon zurück gibt.
Mit dieser Information kann dann jeder weiter verfahren, wie er lustig ist. Zum Beispiel kann man eine eigene Hilfedatei oder auch die Hilfen auf der eigenen Webseite aufrufen.

WICHTIG:
1) DAS VBA CODE FENSTER NICHT (!!) ÖFFNEN, WÄHREND DIE EVENT HOOKS LAUFEN! DAS GIBT UNSCHÖNES GEFLACKER UND ACCESS MUSS NEU GESTARTET WERDEN!

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


Ablauf bei Drücken von F1:

- Die Maus muss sich über einem Element auf dem Ribbon befinden.
- Beim Drücken von F1 wird der Pfad dieses Elements abgefragt und zurück gegeben.
