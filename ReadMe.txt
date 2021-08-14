================================================================================
 MICROSOFT FOUNDATION CLASS-BIBLIOTHEK: JTLSteuer-Projekt�bersicht
================================================================================

Der Anwendungs-Assistent hat diese JTLSteuer-Anwendung f�r Sie 
erstellt. Diese Anwendung zeigt nicht nur die Grundlagen der Verwendung von
Microsoft Foundation Classes, sondern dient auch als Ausgangspunkt f�r das
Schreiben Ihrer Anwendung.

Diese Datei enth�lt eine Zusammenfassung dessen, was Sie in jeder der Dateien
finden, aus denen Ihre JTLSteuer�Anwendung besteht.

JTLSteuer.vcxproj
    Dies ist die Hauptprojektdatei f�r VC++-Projekte, die mithilfe eines 
    Anwendungs-Assistenten erstellt werden. 
    Sie enth�lt Informationen �ber die Version von Visual C++, in der die Datei 
    erzeugt wurde, sowie Informationen �ber die Plattformen, Konfigurationen und 
    Projektfunktionen, die mit dem Anwendungs-Assistenten ausgew�hlt wurden.

JTLSteuer.vcxproj.filters
    Dies ist die Filterdatei f�r VC++-Projekte, die mithilfe eines 
    Anwendungs-Assistenten erstellt werden. 
    Sie enth�lt Informationen �ber die Zuordnung zwischen den Dateien im 
    Projekt und den Filtern. Diese Zuordnung wird in der IDE zur Darstellung der 
    Gruppierung von Dateien mit �hnlichen Erweiterungen unter einem bestimmten 
    Knoten verwendet (z. B. sind CPP-Dateien dem Filter "Quelldateien" 
    zugeordnet).

JTLSteuer.h
    Dies ist die Hauptheaderdatei f�r die Anwendung. Sie enth�lt weitere
    projektspezifische Header (einschlie�lich "Resource.h") und deklariert
    die CJTLSteuerApp-Anwendungsklasse.

JTLSteuer.cpp
    Dies ist die Hauptquelldatei der Anwendung, die die CJTLSteuerApp-
    Anwendungsklasse enth�lt.

JTLSteuer.rc
    Dies ist eine Auflistung aller Microsoft Windows-Ressourcen, die das
    Programm verwendet. Sie enth�lt die Symbole, Bitmaps und Cursor, die im 
    Unterverzeichnis "RES" gespeichert werden. Diese Datei kann direkt in 
    Microsoft Visual C++ bearbeitet werden. Ihre Projektressourcen befinden sich 
    in 1031.

res\JTLSteuer.ico
    Dies ist eine Symboldatei, die als Symbol der Anwendung verwendet wird. 
    Dieses Symbol ist in der Hauptressourcendatei "JTLSteuer.rc" 
    enthalten.

res\JTLSteuer.rc2
    Diese Datei enth�lt Ressourcen, die nicht von Microsoft Visual C++
    bearbeitet werden. Sie sollten alle Ressourcen, die nicht mit dem
    Ressourcen-Editor bearbeitet werden k�nnen, in dieser Datei platzieren.


/////////////////////////////////////////////////////////////////////////////

Der Anwendungs-Assistent erstellt eine Dialogfeldklasse:

JTLSteuerDlg.h, JTLSteuerDlg.cpp � das Dialogfeld
    Diese Dateien enthalten Ihre CJTLSteuerDlg�Klasse. Diese Klasse 
    bestimmt das Verhalten des Hauptdialogfelds Ihrer Anwendung. Die Vorlage des 
    Dialogfelds befindet sich in der Datei "JTLSteuer.rc", die in 
    Microsoft Visual C++ bearbeitet werden kann.


/////////////////////////////////////////////////////////////////////////////

Weitere Funktionen:

ActiveX-Steuerelemente
    Die Anwendung unterst�tzt die Verwendung von ActiveX-Steuerelementen.

/////////////////////////////////////////////////////////////////////////////

Weitere Standarddateien:

"StdAfx.h", "StdAfx.cpp"
    Diese Dateien werden verwendet, um eine vorkompilierte Headerdatei
    (PCH-Datei) mit dem Namen "JTLSteuer.pch.pch2 und eine 
    vorkompilierte Typendatei mit dem Namen "StdAfx.obj" zu erstellen.

"Resource.h"
    Dies ist die Standardheaderdatei, die neue Ressourcen-IDs definiert.
    Microsoft Visual C++ liest und aktualisiert diese Datei.

JTLSteuer.manifest
	Anwendungsmanifestdateien werden von Windows XP verwendet, um eine 
	Anwendungsabh�ngigkeit von verschiedenen Versionen paralleler Assemblys 
        zu beschreiben.
	Das Ladeprogramm verwendet diese Informationen, um die entsprechende 
	Assembly aus dem Assemblycache oder privat aus der Anwendung zu laden. Das
	Anwendungsmanifest kann zur Verteilung als externe Manifestdatei
	enthalten sein, die im gleichen Ordner installiert ist wie die ausf�hrbare 
	Datei der Anwendung, oder sie kann in Form einer Ressource in der 
	ausf�hrbaren Datei enthalten sein. 
/////////////////////////////////////////////////////////////////////////////

Weitere Hinweise:

Der Anwendungs-Assistent verwendet "TODO:", um auf Teile des Quellcodes
hinzuweisen, die Sie erg�nzen oder anpassen sollten.

Wenn Ihre Anwendung MFC in einer freigegebenen DLL verwendet, m�ssen Sie die 
MFC-DLLs verteilen. Wenn die Anwendung eine andere Sprache als die des 
Gebietsschemas des Betriebssystems verwendet, m�ssen Sie au�erdem die 
entsprechenden lokalisierten Ressourcen "MFC100XXX.DLL" verteilen. Weitere 
Informationen zu diesen beiden Themen finden Sie im Abschnitt zum Verteilen 
von Visual C++-Anwendungen in der MSDN-Dokumentation.

/////////////////////////////////////////////////////////////////////////////
