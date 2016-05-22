Starten / Stoppen eines ListenerPorts: controlListenerPort <stage> <node> <server> <action>
Starten / Stoppen eines AppServers: controlAppServer <stage> <node> <server> <action>

<stage> := {T|A|I|P}
<server> := Name eines Application Servers innerhalb von <node>
<node> := Names eines Knotens innerhalb der jew. Cell
<action> := {start|stop}

T, A, I, P kennzeichnen die jeweiligen Stages:
T = Test
A = Abnahme
I = RE2 / Integration
P = Produktion


BEISPIELE:

Stoppen von WTSR01A in WTNODEA:
controlAppServer T WTNODEA WTSR01A stop

Stoppen des DaieListenerPorts innerhalb von WTSR01A:
controlListenerPort T WTNODEA WTSR01A stop
