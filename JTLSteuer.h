
// JTLSteuer.h: Hauptheaderdatei f�r die PROJECT_NAME-Anwendung
//

#pragma once

#ifndef __AFXWIN_H__
	#error "'stdafx.h' vor dieser Datei f�r PCH einschlie�en"
#endif

#include "resource.h"		// Hauptsymbole


// CJTLSteuerApp:
// Siehe JTLSteuer.cpp f�r die Implementierung dieser Klasse
//

class CJTLSteuerApp : public CWinApp
{
public:
	CJTLSteuerApp();

// �berschreibungen
public:
	virtual BOOL InitInstance();

// Implementierung

	DECLARE_MESSAGE_MAP()
};

extern CJTLSteuerApp theApp;