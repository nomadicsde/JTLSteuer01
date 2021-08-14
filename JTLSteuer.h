
// JTLSteuer.h: Hauptheaderdatei für die PROJECT_NAME-Anwendung
//

#pragma once

#ifndef __AFXWIN_H__
	#error "'stdafx.h' vor dieser Datei für PCH einschließen"
#endif

#include "resource.h"		// Hauptsymbole


// CJTLSteuerApp:
// Siehe JTLSteuer.cpp für die Implementierung dieser Klasse
//

class CJTLSteuerApp : public CWinApp
{
public:
	CJTLSteuerApp();

// Überschreibungen
public:
	virtual BOOL InitInstance();

// Implementierung

	DECLARE_MESSAGE_MAP()
};

extern CJTLSteuerApp theApp;