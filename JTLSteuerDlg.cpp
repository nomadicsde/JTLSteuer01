
// JTLSteuerDlg.cpp: Implementierungsdatei
//

#include "stdafx.h"
#include "JTLSteuer.h"
#include "JTLSteuerDlg.h"
#include "afxdialogex.h"
#include "atlpath.h"
#include "CSVFile.h"


bool  g_fOSS = true;

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CJTLSteuerDlg-Dialogfeld

#define FAKTOR_KORREKT false
#define MIT_FAKTOR     false
#define DELTA_CT_MIN   500




CJTLSteuerDlg::CJTLSteuerDlg(CWnd* pParent /*=NULL*/)
              :CDialogEx(CJTLSteuerDlg::IDD, pParent)
              , m_dateVon(COleDateTime::GetCurrentTime())
              , m_dateBis(COleDateTime::GetCurrentTime())
              , m_nStartNummer(1)
              , m_fReadGutschriften(false)
              , m_fReadRechnungen(false)
              , m_fReadRechnungenHaendler(false)
              , m_fReadAmazon(false)
              , m_fReadPOS(false)
              , m_GutschriftNotFound(0)
              , m_GutschriftFound(0)
              , m_GutschriftExceedCount(0)
              , m_GutschriftExceedValue(0)
{

    char szFileName[_MAX_PATH];
    GetModuleFileName(AfxGetInstanceHandle(), szFileName, _MAX_PATH);
    char* pLast = strrchr(szFileName, '\\');
    *(pLast + 1) = '\0';
    strcat(szFileName, "JtlSteuer.ini");
    m_iniFileName = szFileName;

    m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
    m_monthMap["Jan"] = 1;
    m_monthMap["Feb"] = 2;
    m_monthMap["Mar"] = 3;
    m_monthMap["Apr"] = 4;
    m_monthMap["May"] = 5;
    m_monthMap["Jun"] = 6;
    m_monthMap["Jul"] = 7;
    m_monthMap["Aug"] = 8;
    m_monthMap["Sep"] = 9;
    m_monthMap["Oct"] = 10;
    m_monthMap["Nov"] = 11;
    m_monthMap["Dec"] = 12;

    #define ADD_EU(iso)  m_mapEU[iso] = 1;

    // EU-Mitgliedstaaten
    ADD_EU("BE")
    ADD_EU("BG")
    ADD_EU("DK")
    ADD_EU("DE")
    ADD_EU("EE")
    ADD_EU("FI")
    ADD_EU("FR")
    ADD_EU("GR")
    ADD_EU("HR")
    ADD_EU("IE")
    ADD_EU("IT")
    ADD_EU("LV")
    ADD_EU("LT")
    ADD_EU("LU")
    ADD_EU("MT")
    ADD_EU("NL")
    ADD_EU("AT")
    ADD_EU("PL")
    ADD_EU("PT")
    ADD_EU("RO")
    ADD_EU("SE")
    ADD_EU("SK")
    ADD_EU("SI")
    ADD_EU("ES")
    ADD_EU("CZ")
    ADD_EU("HU")
    ADD_EU("GB")
    ADD_EU("CY")
    
    #undef ADD_EU

    ReadSaveData();

    // m_dateVon.SetDate(2019, 01, 01);
    // m_dateBis.SetDate(2019, 12, 31);

}

#define SECTION_STEUER   _T("JTLSteuer")
#define ENTRY_DATUM_VON  _T("Datum Von")
#define ENTRY_DATUM_BIS  _T("Datum Bis")
#define ENTRY_STARTNR    _T("Startnummer")

void CJTLSteuerDlg::ReadSaveData(bool fSave)
{
    LPCSTR  lpszSection = SECTION_STEUER;
    if (fSave)
    {
        CString szData;
        szData.Format("%02d.%02d.%d", m_dateVon.GetDay(), m_dateVon.GetMonth(), m_dateVon.GetYear());
        WritePrivateProfileString(lpszSection, ENTRY_DATUM_VON, szData, m_iniFileName);
        szData.Format("%02d.%02d.%d", m_dateBis.GetDay(), m_dateBis.GetMonth(), m_dateBis.GetYear());
        WritePrivateProfileString(lpszSection, ENTRY_DATUM_BIS, szData, m_iniFileName);
        szData.Format("%d", m_nStartNummer);
        WritePrivateProfileString(lpszSection, ENTRY_STARTNR, szData, m_iniFileName);
    }
    else
    {
        char data[128];
        if (GetPrivateProfileString(lpszSection, ENTRY_DATUM_VON, "", data, sizeof(data), m_iniFileName))
        {
            data[2] = '\0';
            data[5] = '\0';
            m_dateVon.SetDate(atoi(data + 6), atoi(data + 3), atoi(data));
        }
        else
            m_dateVon.SetDate(2019, 01, 01);

        if (GetPrivateProfileString(lpszSection, ENTRY_DATUM_BIS, "", data, sizeof(data), m_iniFileName))
        {
            data[2] = '\0';
            data[5] = '\0';
            m_dateBis.SetDate(atoi(data + 6), atoi(data + 3), atoi(data));
        }
        else
            m_dateBis.SetDate(2019, 12, 31);

        if (GetPrivateProfileString(lpszSection, ENTRY_STARTNR, "", data, sizeof(data), m_iniFileName))
            m_nStartNummer = atoi(data);
        else
            m_nStartNummer = 1;
    }
}

void CJTLSteuerDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialogEx::DoDataExchange(pDX);
    DDX_DateTimeCtrl(pDX, IDC_DATUM_VON, m_dateVon);
    DDX_DateTimeCtrl(pDX, IDC_DATUM_BIS, m_dateBis);
    DDX_Text(pDX, IDC_EDIT_STARTNR, m_nStartNummer);
	DDV_MinMaxInt(pDX, m_nStartNummer, 1, 999999);
}

BEGIN_MESSAGE_MAP(CJTLSteuerDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
    ON_BN_CLICKED(IDC_RECHNUNGEN, &CJTLSteuerDlg::OnBnClickedRechnungen)
    ON_BN_CLICKED(IDC_GUTSCHRIFTEN, &CJTLSteuerDlg::OnBnClickedGutschriften)
    ON_BN_CLICKED(IDCANCEL, &CJTLSteuerDlg::OnBnClickedCancel)
    ON_BN_CLICKED(IDC_AMAZON_TAX, &CJTLSteuerDlg::OnBnAmazonTax)
    ON_BN_CLICKED(IDC_RECHNUNGEN_POS, &CJTLSteuerDlg::OnBnPOS)
    ON_BN_CLICKED(IDC_START_CALC, &CJTLSteuerDlg::OnBnStartCalc)
    ON_NOTIFY(DTN_DATETIMECHANGE, IDC_DATUM_VON, &CJTLSteuerDlg::OnDtnDatetimechangeDatumVon)
    ON_NOTIFY(DTN_DATETIMECHANGE, IDC_DATUM_BIS, &CJTLSteuerDlg::OnDtnDatetimechangeDatumBis)
    ON_EN_CHANGE(IDC_EDIT_STARTNR, &CJTLSteuerDlg::OnEnChangeEditStartnr)
    ON_BN_CLICKED(IDC_RECHNUNGEN_HAENDLER, &CJTLSteuerDlg::OnBnClickedRechnungenHaendler)
END_MESSAGE_MAP()


// CJTLSteuerDlg-Meldungshandler

BOOL CJTLSteuerDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Symbol für dieses Dialogfeld festlegen. Wird automatisch erledigt
	//  wenn das Hauptfenster der Anwendung kein Dialogfeld ist
	SetIcon(m_hIcon, TRUE);			// Großes Symbol verwenden
	SetIcon(m_hIcon, FALSE);		// Kleines Symbol verwenden

	// TODO: Hier zusätzliche Initialisierung einfügen
    SetButtons();
    GetDlgItem(IDC_RECHNUNGEN)->SetFocus();
    SetDefID(IDC_RECHNUNGEN);
    return FALSE;
}

// Wenn Sie dem Dialogfeld eine Schaltfläche "Minimieren" hinzufügen, benötigen Sie
//  den nachstehenden Code, um das Symbol zu zeichnen. Für MFC-Anwendungen, die das 
//  Dokument/Ansicht-Modell verwenden, wird dies automatisch ausgeführt.

void CJTLSteuerDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // Gerätekontext zum Zeichnen

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Symbol in Clientrechteck zentrieren
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Symbol zeichnen
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

// Die System ruft diese Funktion auf, um den Cursor abzufragen, der angezeigt wird, während der Benutzer
//  das minimierte Fenster mit der Maus zieht.
HCURSOR CJTLSteuerDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}



void CJTLSteuerDlg::OnBnClickedRechnungen()
{
  CFileDialog dlg(TRUE);
  if (dlg.DoModal() == IDOK)
  {
      m_fReadRechnungen = DoReadRechnung(dlg.GetPathName(), dlg.GetFolderPath(), dlg.GetFileTitle(), true, true);
      if (m_fReadRechnungen)
      { 
          SetButtons();
          GetDlgItem(IDC_RECHNUNGEN_HAENDLER)->SetFocus();
          SetDefID(IDC_RECHNUNGEN_HAENDLER);
          UpdateWindow();
      }
  }
}


void CJTLSteuerDlg::OnBnClickedRechnungenHaendler()
{
    CFileDialog dlg(TRUE);
    if (dlg.DoModal() == IDOK)
    {
        m_fReadRechnungenHaendler = DoReadRechnung(dlg.GetPathName(), dlg.GetFolderPath(), dlg.GetFileTitle(), false, false);
        if (m_fReadRechnungenHaendler)
        {
            SetButtons();
            GetDlgItem(IDC_RECHNUNGEN_POS)->SetFocus();
            SetDefID(IDC_RECHNUNGEN_POS);
            UpdateWindow();
        }
    }
}

void CJTLSteuerDlg::OnBnPOS()
{
  CFileDialog dlg(TRUE);
  if (dlg.DoModal() == IDOK)
    m_fReadPOS = DoReadPOS(dlg.GetPathName(), dlg.GetFolderPath(), dlg.GetFileTitle());
  else
    m_fReadPOS = true;

  if (m_fReadPOS)
  {
    SetButtons();
    GetDlgItem(IDC_GUTSCHRIFTEN)->SetFocus();
    SetDefID(IDC_GUTSCHRIFTEN);
    UpdateWindow();
  }
}

void CJTLSteuerDlg::OnBnAmazonTax()
{
    if (m_fReadAmazon = ReadAmazon())
    {
        SetButtons();
        GetDlgItem(IDC_START_CALC)->SetFocus();
        SetDefID(IDC_START_CALC);
        UpdateWindow();
    }
    

}

void CJTLSteuerDlg::OnBnStartCalc()
{
    StoreResults();
    GetDlgItem(IDOK)->SetFocus();
    SetDefID(IDOK);
    UpdateWindow();
}


void CJTLSteuerDlg::OnBnClickedGutschriften()
{
  CFileDialog dlg(TRUE);
  if (dlg.DoModal() == IDOK)
  {
    m_fReadGutschriften = DoReadGutschriften(dlg.GetPathName(), dlg.GetFolderPath(), dlg.GetFileTitle());
    if (m_fReadGutschriften)
    {
        SetButtons();
        GetDlgItem(IDC_AMAZON_TAX)->SetFocus();
        SetDefID(IDC_AMAZON_TAX);
        UpdateWindow();
    }
    // StoreResult(dlg.GetFolderPath(), dlg.GetFileTitle());

    LPCSTR lpszArt[] = PLATFORM_ARR_OUT;
    for (int i = 0; i < CRechnung::countPlatform; i++)
        m_arrSummery[i] = "";
    
    // ErstelleEasyCashRechnungen(dlg.GetFolderPath(), dlg.GetFileTitle());
    // ErstelleEasyCashGutschriften(dlg.GetFolderPath(), dlg.GetFileTitle());

    /*
    CString szFilePath, szFile, szFileName, szLine;
    CFile   fl;

    szFilePath = dlg.GetFolderPath();
    szFile = dlg.GetFileTitle();
    szFile.Replace(".csv", "");
    szFileName.Format("%s\\%s_Summery.csv", szFilePath, szFile);

    if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
    {
        
        szLine = ECT_HEADER_SUMMERY;
        fl.Write((LPCSTR)szLine, szLine.GetLength());

        for (int i = 0; i < CRechnung::countPlatform; i++)
        {
            szLine = m_arrSummery[i];
            fl.Write((LPCSTR)szLine, szLine.GetLength());
        }
        fl.Close();

        szLine.Format("Die Summary-Datei\n%s\nwurde erzeugt.", szFileName);
        MessageBox(szLine, "Info", MB_OK);
    }
    */

    // ReadAmazon();
    
    /*
    StoreResult(CRechnung::Amazon,   dlg.GetFolderPath(), dlg.GetFileTitle());
    StoreResult(CRechnung::Nomadics, dlg.GetFolderPath(), dlg.GetFileTitle());
    StoreResult(CRechnung::unknown,  dlg.GetFolderPath(), dlg.GetFileTitle());
    */

  }
}

bool CJTLSteuerDlg::ReadAmazon(void)
{
    CFileDialog dlg(TRUE);
    bool        fReturn = false;
    if (dlg.DoModal() == IDOK) 
        fReturn = DoAmazonSteuer(dlg.GetPathName(), dlg.GetFolderPath(), dlg.GetFileTitle());
    return fReturn;
}

int CJTLSteuerDlg::GetCentBetrag(CString szWert)
{
  szWert.Replace(',', '.');
  double df = atof(szWert);
  double dl = df < 0 ? -0.005 : +0.005;
  return (int)(100.0 * df + dl);
}

int CJTLSteuerDlg::GetIntWert(CString szWert)
{
  szWert.Replace(',', '.');
  double df = atof(szWert);
  double dl = df < 0 ? -0.005 : +0.005;
  return (int)(df + dl);
}


bool CJTLSteuerDlg::DoReadPOS(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName)
{
   
  CCSVFile     csv(lpszFilePath);
  CStringArray arrFields;
  POSDATA      posdata;
  
  CString      szDatum, szRechnungsNr;
  LPCSTR       headerNameArr[] = POS_HEADER_FLD_NAME;
  bool         fLineError      = false;
   
  int          i, lineCount, currentTag, minTag, maxTag, menge, sign;
  int          countHeaderFldName;

  countHeaderFldName = sizeof(headerNameArr) / sizeof(LPCSTR);

  minTag = GetSimpleDay(m_dateVon);
  maxTag = GetSimpleDay(m_dateBis);

  m_mapPOSData.RemoveAll();
  lineCount = 0;

  while (!fLineError && csv.ReadData(arrFields))
  {

    fLineError = countHeaderFldName != arrFields.GetCount();
    if (0 == lineCount++)
    {
      // Prüfe die erste Zeile
      for (i = 0; i < countHeaderFldName && !fLineError; i++)
        fLineError = arrFields[i] != headerNameArr[i];
    }
    else
    {
      if (!fLineError)
      {
        szDatum = arrFields[POS_INDEX_DATUM];
        currentTag = GetSimpleDay(szDatum);

        // Datums-Kontrolle
        if (currentTag >= minTag && currentTag <= maxTag)
        {
          // POS-Rechnung ? 
          szRechnungsNr = arrFields[POS_INDEX_EXTERN_NR];
          if (szRechnungsNr.Left(3) == "RB-")
          {
            

            if (!m_mapPOSData.Lookup(szRechnungsNr, posdata))
            {
              posdata.szAuftragsnummer = arrFields[POS_INDEX_AUFTRAGS_NR];
              posdata.szDatum          = arrFields[POS_INDEX_DATUM];
              posdata.szKundenNummer   = arrFields[POS_INDEX_KD_NR];
              posdata.szUst            = arrFields[POS_INDEX_UST_PROZ];
              posdata.rbNummer         = atoi(szRechnungsNr.Mid(3));
              posdata.nBruttonVK       = 0;
              posdata.nNettoVK         = 0;
            }

            menge = GetIntWert(arrFields[POS_INDEX_ARTIKEL_MENGE]);
            sign  = menge < 0 ? -1 : 1;
            
            posdata.nBruttonVK += (menge * GetCentBetrag(arrFields[POS_INDEX_BRUTTO]));
            posdata.nNettoVK   += (        GetCentBetrag(arrFields[POS_INDEX_NETTO]));

            m_mapPOSData[szRechnungsNr] = posdata;
          }
        }
      }
    }
  }

  return !fLineError;
}

bool CJTLSteuerDlg::DoReadRechnung(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName, bool fKunden, bool fReset)
{

  CStringArray arrFields;
  CCSVFile     csv(lpszFilePath);
  CRechnung    rechnung;
  LPCSTR       headerNameArr[] = RE_HEADER_FLD_NAME;
  LPCSTR       platformNames[] = PLATFORM_ARR;
  CString      szFaktor, szMessage, szRechnung, szDatum, szKundenGruppe;
  double       faktor, dfValue, dfSkonto;
  bool         fLineError = false;
  bool         fFound, fIsSkonto;
  bool         fReturn    = false;
  bool         fKundenGrp = false;

  int          i, lineCount = 0;
  int          currentTag, minTag, maxTag;
  int          countHeaderFldName, countPlatform;

  countPlatform      = sizeof(platformNames) / sizeof(LPCSTR);
  countHeaderFldName = sizeof(headerNameArr) / sizeof(LPCSTR);
  
  if (fReset)
  {
      m_mapRechnung.RemoveAll();
      m_mapAll.RemoveAll();
  }

  minTag = GetSimpleDay(m_dateVon);
  maxTag = GetSimpleDay(m_dateBis);

  while (!fLineError && csv.ReadData(arrFields))
  {
    
    fLineError = countHeaderFldName != arrFields.GetCount();
    if (0 == lineCount++)
    {
      // Prüfe die erste Zeile
      for (i=0; i<countHeaderFldName && !fLineError; i++)
        fLineError = arrFields[i] != headerNameArr[i];
    }
    else
    {
      if (!fLineError)
      {
        szKundenGruppe = arrFields[RE_INDEX_KUNDENGRUPPE];
        fKundenGrp     = szKundenGruppe == TEXT_KUNDENGRUPPE;

        if (fKunden != fKundenGrp)
          continue;
  
        rechnung.Reset();
        szFaktor = arrFields[RE_INDEX_WAEHRUNG_FAKT];
        szFaktor.Replace(',', '.');
        faktor = atof(szFaktor);
        if (0 == faktor)
            faktor = 1;

        // Achtung:
        // Rechnungsbetrag
        if (!fKunden)
        {
            rechnung.m_dfBetragRechnung = atof(arrFields[RE_INDEX_ZAHLUNGSBETRAG])/100.0;
            if (arrFields[RE_INDEX_ZAHLUNGSHINWEIS] == TEXT_SKONTOBUCHUNG)
            {
                dfSkonto  = atof(arrFields[RE_INDEX_BETRAG]) * FAKTOR_SKONTO;
                dfValue   = atof(arrFields[RE_INDEX_ZAHLUNGSBETRAG]) / 100.0;
                
                fIsSkonto = 0 == dfValue || (fabs(dfSkonto - dfValue) / fabs(dfValue) <= 0.01);
                if (fIsSkonto)
                    continue;
                else
                    rechnung.m_dfBetragRechnung = dfValue;
            }
            szDatum = arrFields[RE_INDEX_ZAHLUNGSDATUM];
        }
        else
        {
            rechnung.m_dfBetragRechnung = atof(arrFields[RE_INDEX_BETRAG]);
            szDatum = arrFields[RE_INDEX_DATUM];
        }

        currentTag = GetSimpleDay(szDatum);

        rechnung.m_faktor             = faktor;
        rechnung.m_bestellNummer      = arrFields[RE_INDEX_BESTELLNUMMER];
        rechnung.m_extBestellnummer   = arrFields[RE_INDEX_EXT_BESTELLNUMMER];
        rechnung.m_szRechnungsNr      = arrFields[RE_INDEX_RECHNUNGSNUMMER];
        rechnung.m_szDatum            = szDatum;
        rechnung.m_szUmst             = arrFields[RE_INDEX_UMST];
        rechnung.m_platform           = fKunden ? arrFields[RE_INDEX_PLATTFORM] : platformNames[CRechnung::Haendler];
        rechnung.m_szRechnungEmpf     = arrFields[RE_INDEX_VORNACHNAME];
        rechnung.m_ISO                = arrFields[RE_INDEX_ISO];

        if (rechnung.m_extBestellnummer.IsEmpty())
            rechnung.m_extBestellnummer = rechnung.m_szRechnungsNr;

        if (MIT_FAKTOR)
        {
          if (FAKTOR_KORREKT)
            rechnung.m_dfBetragRechnungKorr = rechnung.m_dfBetragRechnung / faktor;
          else
            rechnung.m_dfBetragRechnungKorr = rechnung.m_dfBetragRechnung * faktor;
        }
        else
          rechnung.m_dfBetragRechnungKorr = rechnung.m_dfBetragRechnung;


        rechnung.m_waehrungRechnung   = arrFields[RE_INDEX_WAEHRUNG];
        
        if (!fKunden)
        {
            rechnung.m_typ = CRechnung::Haendler;
        }
        else
        {
            for (i = 0, fFound = false; !fFound && i < countPlatform; i++)
            {
              if (fFound = rechnung.m_platform == platformNames[i])
                rechnung.m_typ = (CRechnung::typRechnung)i;
            }
        }

        // Als Referenz ist hier stets die externe Nummer - nur falls diese nicht vorhanden, dann die Rechnungsnummer
        szRechnung = rechnung.m_extBestellnummer;
        if (szRechnung.IsEmpty())
          szRechnung = rechnung.m_szRechnungsNr;

        // Speicher nun die Rechnung ...
        m_mapAll[szRechnung]      = rechnung;
        if (currentTag >= minTag && currentTag <= maxTag)
            m_mapRechnung[szRechnung] = rechnung;

        // rechnung.m_typ                = 'A' == arrFields[RE_INDEX_BESTELLNUMMER][0] ? CRechnung::Amazon : ('B' == arrFields[RE_INDEX_BESTELLNUMMER][0] ? CRechnung::Nomadics : CRechnung::unknown);
        // m_mapRechnung[arrFields[RE_INDEX_RECHNUNGSNUMMER]] = rechnung;

      }
    }
    if (fLineError)
    {
      szMessage.Format("In der Zeile %d stimmte die Feldanzahl nicht mit der Soll-Anzahl überein.\n\nDie Funktion wird abgebrochen.", lineCount);
      MessageBox(szMessage, "Achtung", MB_OK);
    }
    else
      arrFields.RemoveAll();
  }

  szMessage.Format("Eingelesen: %d Rechnungen\ndavon %d Rechnungen im Zeitraum.", lineCount, m_mapRechnung.GetCount());
  MessageBox(szMessage, "Info", MB_OK);

  fReturn = !fLineError;
  return fReturn;

}

void CJTLSteuerDlg::ResetMapGutschriftenWawi(void)
{
    for (CMapGutschriftArray::CPair* pCurVal = m_mapArrRechnungGutschriftWawi.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapArrRechnungGutschriftWawi.PGetNextAssoc(pCurVal))
        delete pCurVal->value;
    m_mapArrRechnungGutschriftWawi.RemoveAll();
}


void CJTLSteuerDlg::ResetMapGutschriftenAmazon(void)
{
    for (CMapGutschriftArray::CPair* pCurVal = m_mapArrRechnungGutschriftAmazon.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapArrRechnungGutschriftAmazon.PGetNextAssoc(pCurVal))
        delete pCurVal->value;
    m_mapArrRechnungGutschriftAmazon.RemoveAll();
}

bool CJTLSteuerDlg::DoReadGutschriften(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName)
{
  CArrGutschrift* parrRechnungGutschrift;
  CStringArray    arrFields;
  CCSVFile        csv(lpszFilePath);
  CRechnung       rechnung;
  CGutschrift     gutschrift;
  LPCSTR          headerNameArr[] = GU_HEADER_FLD_NAME;
  double          faktor, dfBetrag, summeTotal;
  bool            fRechnungInYear, fLineError;

  CString         szFaktor, rechnungsNummer, szError, szMessage, bestellNummer, bestellNummerExt;
  bool            fHasError;
  int             countHeaderFldName, lineCount;

  m_GutschriftNotFound    = 0;
  m_GutschriftFound       = 0;
  m_GutschriftExceedCount = 0;
  m_GutschriftExceedValue = 0;

  
  m_mapGutschrift.RemoveAll();
  ResetMapGutschriftenWawi();

  countHeaderFldName = sizeof(headerNameArr) / sizeof(LPCSTR);
  lineCount          = 0;
  fLineError         = false;
  summeTotal         = 0;

  while (!fLineError && csv.ReadData(arrFields))
  {
    
    fHasError  = false;
    fLineError = countHeaderFldName != arrFields.GetCount();
    
    if (0 == lineCount++)
    {
      // Prüfe die erste Zeile
      for (int i=0; i<countHeaderFldName && !fLineError; i++)
        fLineError = arrFields[i] != headerNameArr[i];
    }
    else
    {
      
      rechnungsNummer  = arrFields[GU_INDEX_RECHNUNGSNUMMER];
      bestellNummerExt = arrFields[GU_INDEX_BESTELLNUMMER];
      

      if (bestellNummerExt.IsEmpty())
        bestellNummerExt = rechnungsNummer;

      fRechnungInYear = m_mapRechnung.Lookup(bestellNummerExt, rechnung);
      if (!fRechnungInYear && !m_mapAll.Lookup(bestellNummerExt, rechnung))
      {
        rechnung.Reset();
        rechnung.m_szRechnungsNr = rechnungsNummer;
        rechnung.m_szDatum = arrFields[GU_INDEX_DATUM];
        rechnung.m_szUmst = arrFields[GU_INDEX_UMST];
        rechnung.m_extBestellnummer = bestellNummerExt;
        bestellNummer = bestellNummerExt;
        rechnung.m_ISO = arrFields[GU_INDEX_ISO]; 

        rechnung.m_typ = 'B' == bestellNummer[0] || 'S' == bestellNummer[0] || '#' == bestellNummer[0] || ' ' == bestellNummer[0] || "" == bestellNummer ? CRechnung::OnlineShop : CRechnung::Amazon_DE;
      }

      szFaktor                       = arrFields[GU_INDEX_WAHRUNG_FAKT];
      szFaktor.Replace(',', '.');
      faktor                         = atof(szFaktor);
      if (0 == faktor)
        faktor = 1;
      
      dfBetrag                          = atof(arrFields[GU_INDEX_BETRAG]);
      rechnung.m_dfBetragGutschriftOrg += dfBetrag;

      dfBetrag = atof(arrFields[GU_INDEX_BETRAG]);
      
      rechnung.m_gutschriftNummer         = arrFields[GU_INDEX_GUTSCHRIFTNUMMER]; // sollte nicht mehr genutzt werden
      
      rechnung.m_dfBetragGutschriftWawi  += dfBetrag;
      rechnung.m_countGutschriftWawi     += 1;

      rechnung.m_waehrungGutschrift       = arrFields[GU_INDEX_WAHRUNG];
      rechnung.m_fHasGutschriftWawi       = true;

      summeTotal                         += dfBetrag; 
      
      // Aktualisiere die Rechnung
      if (fRechnungInYear)
        m_mapRechnung[bestellNummerExt]     = rechnung;
      else
        m_mapAll[bestellNummerExt] = rechnung;


      gutschrift.Reset();
      
      gutschrift.m_rechnungsNummer        = rechnung.m_szRechnungsNr;
      gutschrift.m_szRechnungsDatum       = rechnung.m_szDatum;
      gutschrift.m_wawiBestellNummer      = rechnung.m_bestellNummer;
      gutschrift.m_extBestellNummer       = rechnung.m_extBestellnummer;
      gutschrift.m_gutschriftNummer       = rechnung.m_gutschriftNummer;

      gutschrift.m_platform               = rechnung.m_platform;
      gutschrift.m_typ                    = rechnung.m_typ;

      gutschrift.m_waehrungGutschrift     = rechnung.m_waehrungGutschrift;
      gutschrift.m_dfRechnungsBetrag      = rechnung.m_dfBetragRechnung;
      gutschrift.m_faktorRechnung         = rechnung.m_faktor;
      gutschrift.m_empfaenger             = rechnung.m_szRechnungEmpf;

      gutschrift.m_dfBetragGutschrift     = dfBetrag;
      gutschrift.m_dfBetragGutschriftOrg  = dfBetrag;                                          // rechnung.m_dfBetragGutschriftOrg;
      gutschrift.m_szDatum                = arrFields[GU_INDEX_DATUM];
      gutschrift.m_szUmst                 = arrFields[GU_INDEX_UMST];
      gutschrift.m_faktor                 = faktor;
      gutschrift.m_ISO                    = rechnung.m_ISO;

      // Hinzufügen der Gutschrift zur Liste für Rechnung ...
      if (!m_mapArrRechnungGutschriftWawi.Lookup(bestellNummerExt, parrRechnungGutschrift))
        m_mapArrRechnungGutschriftWawi[bestellNummerExt] = new CArrGutschrift();
      m_mapArrRechnungGutschriftWawi[bestellNummerExt]->Add(gutschrift);
      
      m_mapGutschrift[gutschrift.m_gutschriftNummer] = gutschrift;

      /*
        if (!m_mapGutschrift.Lookup(rechnung.m_szRechnungsNr, gutschrift))
        {
              gutschrift.m_rechnungsNummer = rechnung.m_szRechnungsNr;
              gutschrift.m_wawiBestellNummer = rechnung.m_bestellNummer;
              gutschrift.m_extBestellNummer = rechnung.m_extBestellnummer;
              gutschrift.m_gutschriftNummer = rechnung.m_gutschriftNummer;

              gutschrift.m_platform = rechnung.m_platform;
              gutschrift.m_typ = rechnung.m_typ;

              gutschrift.m_waehrungGutschrift = rechnung.m_waehrungGutschrift;
              gutschrift.m_dfRechnungsBetrag  = rechnung.m_dfBetragRechnung;
              gutschrift.m_faktorRechnung     = rechnung.m_faktor;
              gutschrift.m_empfaenger         = rechnung.m_szRechnungEmpf;
        }


        gutschrift.m_dfBetragGutschrift           = dfBetrag;
        gutschrift.m_dfBetragGutschriftOrg        = dfBetrag; // rechnung.m_dfBetragGutschriftOrg;
        gutschrift.m_szDatum                      = arrFields[GU_INDEX_DATUM];
        gutschrift.m_szUmst                       = arrFields[GU_INDEX_UMST];
        gutschrift.m_faktor                       = faktor;

        if (0 == gutschrift.m_faktorRechnung)
           gutschrift.m_faktorRechnung = gutschrift.m_faktor;

        m_mapGutschrift[rechnung.m_szRechnungsNr] = gutschrift;
        */
    }
    
    if (fHasError)
    {
      szMessage.Format("Es trat folgender Fehler auf:\n%s\n", szError);
      MessageBox(szMessage, "Achtung", MB_OK);
    }
    else
      arrFields.RemoveAll();

  }

  szMessage.Format("Eingelesen: %d Gutschriften.\nGesamtsumme: %lf", lineCount, summeTotal);
  MessageBox(szMessage, "Info", MB_OK);

  return !fLineError;

}

void CJTLSteuerDlg::StoreResult(LPCSTR lpsPath, LPCSTR lpszName)
{
  
  CString szNewFilePath, szLineCsv, szCsv, szType, szInfo;
  LPCSTR  szArrPlatform[] = PLATFORM_ARR_OUT;
  int     i;
  long    index, length;

  double  dfTotalRechnung[COUNT_PLATFORM], dfTotalGutschrift[COUNT_PLATFORM];
  int     countRechnung[COUNT_PLATFORM], countGutschrift[COUNT_PLATFORM];

  // Initialisierung
  for (i = 0; i < COUNT_PLATFORM; i++)
  {
    dfTotalRechnung[i]   = 0;
    dfTotalGutschrift[i] = 0;
    countRechnung[i]     = 0; 
    countGutschrift[i]   = 0;
  }
  
  // Gehe nun alle durch ...
  for (CMapRechnung::CPair* pCurVal = m_mapRechnung.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapRechnung.PGetNextAssoc(pCurVal))
  {
    index = (int)pCurVal->value.m_typ;
    dfTotalRechnung[index]   += pCurVal->value.m_dfBetragRechnungKorr;
    dfTotalGutschrift[index] += pCurVal->value.m_dfBetragGutschriftWawi;
    countRechnung[index]     += 1;
    countGutschrift[index]   += pCurVal->value.m_fHasGutschriftWawi ? 1 : 0;
  }

  for (CMapRechnung::CPair* pCurVal = m_mapAll.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapAll.PGetNextAssoc(pCurVal))
  {
      index = (int)pCurVal->value.m_typ;
      if (pCurVal->value.m_fHasGutschriftWawi)
      {
          dfTotalGutschrift[index] += pCurVal->value.m_dfBetragGutschriftWawi;
          countGutschrift[index]   += 1;
      }
  }

  szCsv = "\"Platform\";\"Rechnung Betrag\";\"Anzahl Rechnung\";\"Gutschrift Betrag\";\"Anzahl Gutschrift\";\"Gesamt-Betrag\"\r\n";
  for (i = 0; i < COUNT_PLATFORM; i++)
  {
    szLineCsv.Format("%s;%6.2f;%d;%6.2f;%d;%6.2f\r\n", szArrPlatform[i], dfTotalRechnung[i]/1.19, countRechnung[i], dfTotalGutschrift[i] / 1.19, countGutschrift[i], (dfTotalRechnung[i] + dfTotalGutschrift[i])/1.19);
    szLineCsv.Replace('.', ',');
    szCsv += szLineCsv;
  }

  length = szCsv.GetLength();

  szNewFilePath.Format("%s\\%s_Platform.csv", lpsPath, lpszName);
  CFile fl;
  if (fl.Open(szNewFilePath, CFile::modeCreate | CFile::modeWrite))
  {
    fl.Write((LPCSTR)szCsv, length);
    fl.Close();
  }

  CString szMessage;
  szMessage.Format("Die berechneten Daten sind in der Datei\n%s\ngespeichert.\n\n\n%s\n", szNewFilePath, szCsv);
  MessageBox(szMessage, "Info", MB_OK);

}

// --------------------------------------------------------------------------------------------------------------------------------------------

double GetPreis(LPCSTR szPreis)
{
    double pr = 0;
    if (szPreis)
    {
        CString preis(szPreis);
        if (preis.GetLength() > 0)
        {
            preis.Replace(',', '.');
            pr = atof(preis);
        }
    }
    return pr;
}


#define SET_AS_INDEX(id) int _##id = -1;
#define GET_AS_INDEX(id) _##id = mapHeaderIndex[headerNameArr[id]];



bool CJTLSteuerDlg::DoAmazonSteuer(LPCSTR lpszFilePath, LPCSTR lpszPath, LPCSTR szFileTitle)
{

  CRechnung::typRechnung typ;
  CMapVerkaufInfo       *pMap;
  CVerkaufInfo           verkaufInfo;
  CStringArray           arrFields;
  CCSVFile               csv(lpszFilePath, ',');
  CRechnung              rechnung;
  CGutschrift            gutschrift;
  CString                szDatumAmazon, szFaktor, rechnungsNummer, szError, szMessage, bestellNummer, bestellNummerExt, szRechnungsnummer, sandale;
  CString                szBestellnummer, szPlatform, platform, szTyp, szUmst, sku, szOSS;
  LPCSTR                 lpszPlatform[] = PLATFORM_ARR;

  double                 summeTotal, dfPreis, dfUmst, dfGutDiffTotal, dfRechDiffTotal, dfWaehrungFaktor, betrag, betragEur;
  double                 sumReDe, sumReFr, sumReGb, sumGuDe, sumGuFr, sumGuGb, sumErDe, sumErFr, sumErGb;
  bool                   fRechnungInDatum, fHasError, fLineError, fRechnung, fErstattung, fRefund, fOSS;
  long                   ctPreis, ctPreisCmp;
  int                    i, countErstattung, countRechnung, countGutschrift, countUnknown, countHeaderFldName, lineCount, countNeuRechnung,
                         nDatumTag, minTag, maxTag, anzahl, multiplikator, countHeaderFields, index;

  LPCSTR                 headerNameArr[] = AS_HEADER_FLD_NAME;

    SET_AS_INDEX(AS_INDEX_MARKETPLACE)
    SET_AS_INDEX(AS_INDEX_DATUM)
    SET_AS_INDEX(AS_INDEX_TYP)
    SET_AS_INDEX(AS_INDEX_BESTELLNR_EXT)
    SET_AS_INDEX(AS_INDEX_SKU)
    SET_AS_INDEX(AS_INDEX_ANZAHL)
    SET_AS_INDEX(AS_INDEX_TAX_REP_SCHEME)
    SET_AS_INDEX(AS_INDEX_UMST)
    SET_AS_INDEX(AS_INDEX_WAEHRUNG)
    SET_AS_INDEX(AS_INDEX_PREIS_BRUTTO)
    SET_AS_INDEX(AS_INDEX_PREIS_UMST)
    SET_AS_INDEX(AS_INDEX_PREIS_NETTO)
    SET_AS_INDEX(AS_INDEX_VERSAND_BRUTTO)
    SET_AS_INDEX(AS_INDEX_VERSAND_UMST)
    SET_AS_INDEX(AS_INDEX_VERSAND_NETTO)
    SET_AS_INDEX(AS_INDEX_VERSAND_AKT_BRUTTO)
    SET_AS_INDEX(AS_INDEX_VERSAND_AKT_UMST)
    SET_AS_INDEX(AS_INDEX_VERSAND_AKT_NETTO)
    SET_AS_INDEX(AS_INDEX_GESCHENK_BRUTTO)
    SET_AS_INDEX(AS_INDEX_GESCHENK_UMST)
    SET_AS_INDEX(AS_INDEX_GESCHENK_NETTO)
    SET_AS_INDEX(AS_INDEX_UMST_ID_KAEUFER)
    SET_AS_INDEX(AS_INDEX_UMST_BERECHNET_CUR)
    SET_AS_INDEX(AS_INDEX_FAKTOR_WAEHRUNG)
    SET_AS_INDEX(AS_INDEX_UMST_BERECHNET)
    SET_AS_INDEX(AS_INDEX_EXPORT_NONEU)
    SET_AS_INDEX(AS_INDEX_EMPF_PLZ)
    SET_AS_INDEX(AS_INDEX_EMPF_STADT)
    SET_AS_INDEX(AS_INDEX_EMPF_ISO)

    CMap<CString, LPCSTR, int, int> mapHeaderIndex;

  m_arrGuKorr.RemoveAll();
  ResetMapGutschriftenAmazon();

  countHeaderFldName = sizeof(headerNameArr) / sizeof(LPCSTR);
  lineCount = 0;
  fLineError = false;
  summeTotal = 0;

  countErstattung  = 0;
  countRechnung    = 0;
  countGutschrift  = 0;
  countUnknown     = 0;
  countNeuRechnung = 0;
  
  dfGutDiffTotal   = 0;
  dfRechDiffTotal  = 0;

  sumReDe          = 0;
  sumReFr          = 0;
  sumReGb          = 0;
  sumGuDe          = 0;
  sumGuFr          = 0;
  sumGuGb          = 0;
  sumErDe          = 0;
  sumErFr          = 0;
  sumErGb          = 0;

  minTag           = GetSimpleDay(m_dateVon);
  maxTag           = GetSimpleDay(m_dateBis);

  while (!fLineError && csv.ReadData(arrFields))
  {

      fHasError = false;
      // fLineError = countHeaderFldName != arrFields.GetCount();
      szError = "Anzahl der Feldnamen stimmt nicht mit der Vorgabe überein";

      if (0 == lineCount++) 
      {
          // Prüfe die erste Zeile
          countHeaderFields = arrFields.GetCount();
          for (i = 0; i < arrFields.GetCount(); i++)
              mapHeaderIndex[arrFields[i]] = i;
          
          for (i = 0; i < countHeaderFldName && !fHasError; i++)
              fHasError = !mapHeaderIndex.Lookup(headerNameArr[i], index);

          if (!fHasError)
          {
                GET_AS_INDEX(AS_INDEX_MARKETPLACE)
                GET_AS_INDEX(AS_INDEX_DATUM)
                GET_AS_INDEX(AS_INDEX_TYP)
                GET_AS_INDEX(AS_INDEX_BESTELLNR_EXT)
                GET_AS_INDEX(AS_INDEX_SKU)
                GET_AS_INDEX(AS_INDEX_ANZAHL)
                GET_AS_INDEX(AS_INDEX_TAX_REP_SCHEME)
                GET_AS_INDEX(AS_INDEX_UMST)
                GET_AS_INDEX(AS_INDEX_WAEHRUNG)
                GET_AS_INDEX(AS_INDEX_PREIS_BRUTTO)
                GET_AS_INDEX(AS_INDEX_PREIS_UMST)
                GET_AS_INDEX(AS_INDEX_PREIS_NETTO)
                GET_AS_INDEX(AS_INDEX_VERSAND_BRUTTO)
                GET_AS_INDEX(AS_INDEX_VERSAND_UMST)
                GET_AS_INDEX(AS_INDEX_VERSAND_NETTO)
                GET_AS_INDEX(AS_INDEX_VERSAND_AKT_BRUTTO)
                GET_AS_INDEX(AS_INDEX_VERSAND_AKT_UMST)
                GET_AS_INDEX(AS_INDEX_VERSAND_AKT_NETTO)
                GET_AS_INDEX(AS_INDEX_GESCHENK_BRUTTO)
                GET_AS_INDEX(AS_INDEX_GESCHENK_UMST)
                GET_AS_INDEX(AS_INDEX_GESCHENK_NETTO)
                GET_AS_INDEX(AS_INDEX_UMST_ID_KAEUFER)
                GET_AS_INDEX(AS_INDEX_UMST_BERECHNET_CUR)
                GET_AS_INDEX(AS_INDEX_FAKTOR_WAEHRUNG)
                GET_AS_INDEX(AS_INDEX_UMST_BERECHNET)
                GET_AS_INDEX(AS_INDEX_EXPORT_NONEU)
                GET_AS_INDEX(AS_INDEX_EMPF_PLZ)
                GET_AS_INDEX(AS_INDEX_EMPF_STADT)
                GET_AS_INDEX(AS_INDEX_EMPF_ISO)
          }

          /*
          for (int i = 0; i < countHeaderFldName && !fLineError; i++)
          {
              fLineError = arrFields[i] != headerNameArr[i];
              if (fLineError)
              {
                  szError.Format("Fehler bei Feldnummer %d. Soll: \"%s\" - Ist: \"%s\"", i, headerNameArr[i], arrFields[i]);
                  fHasError = true;
              }
          }
          */
          fLineError = fHasError;
              
      }
      else
      {
          szOSS = arrFields[_AS_INDEX_TAX_REP_SCHEME];
          szTyp = arrFields[_AS_INDEX_TYP];
          fErstattung = false; // szTyp == "REFUND";
          fRefund = szTyp == "REFUND";

          szPlatform = arrFields[_AS_INDEX_MARKETPLACE];
          typ        = szPlatform == "DE" ? CRechnung::Amazon_DE : (szPlatform == "FR" ? CRechnung::Amazon_Fr : (szPlatform == "GB" ? CRechnung::Amazon_GB : CRechnung::unknown));
          szPlatform = lpszPlatform[typ];

          szUmst = arrFields[_AS_INDEX_UMST];
          dfUmst = 100.0 * atof(szUmst);
          szUmst.Format("%2.1f", dfUmst);

          dfWaehrungFaktor = 1;
          szFaktor = arrFields[_AS_INDEX_FAKTOR_WAEHRUNG];
          if (szFaktor.GetLength() > 0)
          {
              double fak = atof(szFaktor);
              if (0 < fak)
                  dfWaehrungFaktor = (double)1.0 / fak;
          }

          fRechnung = szTyp == "SHIPMENT";
          fOSS      = szOSS == "VCS_EU_OSS";

          // Merke die Sandalen für die Statistik ...
          pMap          = NULL;
          sku           = arrFields[_AS_INDEX_SKU];
          anzahl        = atoi(arrFields[_AS_INDEX_ANZAHL]);
          betrag        = fabs(atof(arrFields[_AS_INDEX_PREIS_BRUTTO]));
          betragEur     = betrag / dfWaehrungFaktor;
          multiplikator = fRechnung ? 1 : -1;

          switch (typ)
          {
              case CRechnung::Amazon_DE:
                  pMap = &m_mapVerkaufDE;
              break;
              case CRechnung::Amazon_Fr:
                  pMap = &m_mapVerkaufFR;
              break;
              case CRechnung::Amazon_GB:
                  pMap = &m_mapVerkaufGB;
              break;
          }
          if (pMap)
          {
              
              if (!pMap->Lookup(sku, verkaufInfo))
                  verkaufInfo.Reset(sku);
              verkaufInfo.anzahlGesamt         += ( multiplikator * anzahl );
              verkaufInfo.summmeEinnahmeOrg    += ( multiplikator * betrag );
              verkaufInfo.summmeEinnahmeFaktor += ( multiplikator * betragEur );
              verkaufInfo.anzahlRetoure        += fRechnung ? 0 : 1;
              verkaufInfo.anzahlBestellung     += fRechnung ? 1 : 0;
              pMap->SetAt(sku, verkaufInfo);
          }

          if (!fErstattung)
          {
              // fRechnung = szTyp == "SHIPMENT";
              
              if (fRechnung || szTyp == "RETURN" || fRefund)
              {

                  /*
                  dfWaehrungFaktor = 1;
                  szFaktor         = arrFields[AS_INDEX_FAKTOR_WAEHRUNG];
                  if (szFaktor.GetLength() > 0)
                  {
                      double fak = atof(szFaktor);
                      if (0 < fak)
                          dfWaehrungFaktor = (double)1.0 / fak;
                  }
                  */

                  if (fRechnung)
                    countRechnung += 1;
                  else
                    countErstattung += 1;

                  // Berechne den Gesamtbetrag ...
                  
                  dfPreis       = GetPreis(arrFields[_AS_INDEX_PREIS_BRUTTO]) + GetPreis(arrFields[_AS_INDEX_VERSAND_BRUTTO]) + GetPreis(arrFields[_AS_INDEX_VERSAND_AKT_BRUTTO]) + GetPreis(arrFields[AS_INDEX_GESCHENK_BRUTTO]);
                  dfUmst        = GetPreis(arrFields[_AS_INDEX_PREIS_UMST])   + GetPreis(arrFields[_AS_INDEX_GESCHENK_UMST]) + GetPreis(arrFields[_AS_INDEX_VERSAND_UMST]) + GetPreis(arrFields[AS_INDEX_VERSAND_AKT_UMST]);
                  szDatumAmazon = GetDate(arrFields[_AS_INDEX_DATUM]);

                  nDatumTag = GetSimpleDay(szDatumAmazon);
                  

                  szBestellnummer = arrFields[_AS_INDEX_BESTELLNR_EXT];

                  fRechnungInDatum = m_mapRechnung.Lookup(szBestellnummer, rechnung);
                  if (!fRechnungInDatum && !m_mapAll.Lookup(szBestellnummer, rechnung))
                  {
                      
                      //SKU. Ship From Postal Code Ship To City 
                      // 

                      rechnung.Reset();

                      rechnung.m_bestellNummer        = szBestellnummer;
                      rechnung.m_szRechnungsNr        = szBestellnummer;

                      rechnung.m_typ                  = typ;
                      rechnung.m_platform             = szPlatform;
                      rechnung.m_fAmazonNeuRechnung   = true;
                      rechnung.m_fAmazonNeuGutschrift = !fRechnung;
                      rechnung.m_waehrungRechnung     = arrFields[_AS_INDEX_WAEHRUNG];
                      rechnung.m_szUmst               = szUmst;
                      rechnung.m_faktor               = dfWaehrungFaktor;

                      rechnung.m_dfBetragRechnung     = fRechnung ? dfPreis : 0;
                      rechnung.m_dfBetragRechnungKorr = rechnung.m_dfBetragRechnung;
                      rechnung.m_szDatum             = szDatumAmazon;
                      rechnung.m_ISO                 = arrFields[_AS_INDEX_EMPF_ISO];

                      rechnung.m_szRechnungEmpf.Format("%s, %s %s", arrFields[_AS_INDEX_SKU], arrFields[_AS_INDEX_EMPF_PLZ], arrFields[_AS_INDEX_EMPF_STADT]);

                      // Kontrolle, ob Rechnung im Bereich ist ...
                      fRechnungInDatum                = nDatumTag >= minTag && nDatumTag <= maxTag;

                      countNeuRechnung               += 1;
                  }
                  else
                  {
                      // Kontrolliere die Umsatzsteuer ..
                      if (atof(rechnung.m_szUmst) != atof(szUmst))
                          rechnung.m_szUmst = szUmst;
                  }
                  
                  rechnung.m_fAmazonTaxReport = true;
                  rechnung.m_fAmazonOSS       = fOSS;

                  /*
                  if (szBestellnummer == "304-5723581-4861919")
                  { 
                      CString msg;
                      msg.Format("Rechnungsnummer: %s - Preis: %lf", rechnung.m_szRechnungsNr, dfPreis);
                      MessageBox(msg, szTyp, MB_OK);
                  }
                  */

                  if (fRechnung)
                  {
                      rechnung.m_dfRechnungSumme            += dfPreis;
                      /*
                      ctPreis                                = (int)(100.0 * dfPreis + 0.005);
                      ctPreisCmp                             = (int)(100.0 * rechnung.m_dfBetragRechnung + 0.005);
                      rechnung.m_fAmazonDiffBetragRechnung   = ctPreis != ctPreisCmp;
                      if (abs(ctPreis) != abs(ctPreisCmp))
                          dfRechDiffTotal += (fabs(dfPreis) - fabs(rechnung.m_dfBetragRechnung));
                       */

                      switch (typ)
                      {
                          case CRechnung::Amazon_DE:
                              sumReDe += dfPreis;
                          break;
                          case CRechnung::Amazon_Fr:
                              sumReFr += dfPreis;
                          break;
                          case CRechnung::Amazon_GB:
                              sumReGb += dfPreis;
                          break;
                      }

                  }
                  else
                  {
                      
                        rechnung.m_fHasGutschriftAmazon      = true;
                        rechnung.m_dfBetragGutschriftAmazon += dfPreis;
                        rechnung.m_countGutschriftAmazon    += 1;
                        if (!fRechnung && rechnung.m_szDatumGutschriftAmazon.IsEmpty())
                            rechnung.m_szDatumGutschriftAmazon = szDatumAmazon;
                      
                      // Vergleich mit Gutschrift-Mapping über die Rechnungsnummer

                      szRechnungsnummer = rechnung.m_szRechnungsNr;
                      if (szRechnungsnummer.IsEmpty())
                          szRechnungsnummer = rechnung.m_extBestellnummer;

                      /*
                      if (!m_mapGutschrift.Lookup(szRechnungsnummer, gutschrift))
                      {
                      */
                      gutschrift.Reset();
                          
                          
                      gutschrift.m_rechnungsNummer       = szRechnungsnummer;
                      gutschrift.m_szRechnungsDatum      = rechnung.m_szDatum;
                      gutschrift.m_wawiBestellNummer     = rechnung.m_bestellNummer;
                      gutschrift.m_extBestellNummer      = rechnung.m_extBestellnummer;
                      gutschrift.m_gutschriftNummer      = rechnung.m_gutschriftNummer;
                      gutschrift.m_szUmst                = rechnung.m_szUmst;

                      gutschrift.m_platform              = rechnung.m_platform;
                      gutschrift.m_typ                   = rechnung.m_typ;

                      gutschrift.m_dfRechnungsBetrag     = rechnung.m_dfBetragRechnung;
                      gutschrift.m_empfaenger            = rechnung.m_szRechnungEmpf;
                      gutschrift.m_ISO                   = rechnung.m_ISO;
                      gutschrift.m_faktorRechnung        = dfWaehrungFaktor;
                      gutschrift.m_faktor                = dfWaehrungFaktor;

                      gutschrift.m_dfBetragGutschrift    = dfPreis;
                      gutschrift.m_dfBetragGutschriftOrg = dfPreis;
                      gutschrift.m_waehrungGutschrift    = rechnung.m_waehrungGutschrift;

                      gutschrift.m_fAmazonNeuGutschrift  = true;
                      gutschrift.m_szDatum               = szDatumAmazon;

                      // }
                      gutschrift.m_fErstattung           = fRefund;

                      gutschrift.m_fAmazonTaxReport      = true;
                      gutschrift.m_fAmazonOSS            = fOSS;


                      /*
                        pDatum  = NULL;
                        pBetrag = NULL;
                        switch (gutschrift.m_countItem)
                        {
                        case 0: 
                            pDatum  = &gutschrift.m_szDatum1;
                            pBetrag = &gutschrift.m_dfBetrag1;
                        break;
                        case 1:
                            pDatum  = &gutschrift.m_szDatum2;
                            pBetrag = &gutschrift.m_dfBetrag2;
                        break;
                        case 2:
                            pDatum  = &gutschrift.m_szDatum3;
                            pBetrag = &gutschrift.m_dfBetrag3;
                        break;

                        }
                        if (NULL != pDatum)
                            *pDatum = szDatumAmazon;
                        if (NULL != pBetrag)
                            *pBetrag = dfPreis;

                        gutschrift.m_countItem += 1;
                        gutschrift.m_dfGutschriftSumme     += dfPreis;
                      */
                      
                      /*
                      if (-124 == gutschrift.m_dfGutschriftSumme)
                          break;
                      */
                      
                      CArrGutschrift* parrRechnungGutschrift;
                      if (!m_mapArrRechnungGutschriftAmazon.Lookup(bestellNummerExt, parrRechnungGutschrift))
                          m_mapArrRechnungGutschriftAmazon[bestellNummerExt] = new CArrGutschrift();
                      m_mapArrRechnungGutschriftAmazon[bestellNummerExt]->Add(gutschrift);

                      /*
                      m_mapGutschrift[szRechnungsnummer]  = gutschrift;
                      
                      
                      rechnung.m_dfGutschriftSumme += dfPreis;
                      rechnung.m_fHasGutschrift     = true;
                      */

                      /*
                      ctPreis = (int)(100.0 * dfPreis + 0.005);
                      ctPreisCmp = (int)(100.0 * rechnung.m_dfBetragGutschriftOrg + 0.005);
                      rechnung.m_fAmazonDiffBetragGutschrift = ctPreis != ctPreisCmp;
                      if (abs(ctPreis) != abs(ctPreisCmp))
                          dfGutDiffTotal += (fabs(dfPreis) - fabs(rechnung.m_dfBetragGutschriftOrg));
                      */

                      switch (typ)
                      {
                      case CRechnung::Amazon_DE:
                          sumGuDe += dfPreis;
                      break;
                      case CRechnung::Amazon_Fr:
                          sumGuFr += dfPreis;
                      break;
                      case CRechnung::Amazon_GB:
                          sumGuGb += dfPreis;
                      break;
                      }
                  }

              }
              else
                  countUnknown += 1;

              if (fRechnungInDatum)
                  m_mapRechnung[szBestellnummer] = rechnung;
              else
                  m_mapAll[szBestellnummer] = rechnung;

          }
          
          if (fRefund)
          {
              switch (typ)
              {
                case CRechnung::Amazon_DE:
                  sumErDe += dfPreis;
                break;
                case CRechnung::Amazon_Fr:
                  sumErFr += dfPreis;
                break;
                case CRechnung::Amazon_GB:
                  sumErGb += dfPreis;
                break;
              }
              countErstattung += 1;
          }


      }

      if (fHasError)
      {
          szMessage.Format("Es trat folgender Fehler auf:\n%s\n", szError);
          MessageBox(szMessage, "Achtung", MB_OK);
      }
      else
          arrFields.RemoveAll();
  }


  CRechnung* pRechnung;
  CString    szGuLine;
  double     dfRechDiff, dfGutDiff;
  double     dfDiffReDe, dfDiffReFr, dfDiffReGb, dfDiffGuDe, dfDiffGuFr, dfDiffGuGb;

  dfDiffReDe = 0;
  dfDiffReFr = 0; 
  dfDiffReGb = 0;
  dfDiffGuDe = 0;
  dfDiffGuFr = 0;
  dfDiffGuGb = 0;

  // Hier nun alle durchgehen ...
  for (CMapRechnung::CPair* pCurVal = m_mapRechnung.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapRechnung.PGetNextAssoc(pCurVal))
  {
      
      
      pRechnung = &pCurVal->value;
      
      dfRechDiff = 0;
      dfGutDiff  = 0;
      
      if (pRechnung->m_dfRechnungSumme > 0)
      {
          dfPreis    = pRechnung->m_dfRechnungSumme;
          ctPreis    = (int)(100.0 * pRechnung->m_dfRechnungSumme + 0.005);
          ctPreisCmp = (int)(100.0 * pRechnung->m_dfBetragRechnung + 0.005);
          
          pRechnung->m_fAmazonDiffBetragRechnung = ctPreis != ctPreisCmp;
          if (abs(ctPreis) != abs(ctPreisCmp))
              dfRechDiff = (fabs(dfPreis) - fabs(rechnung.m_dfBetragRechnung));
      }
      
      if (pRechnung->m_fHasGutschriftWawi || pRechnung->m_fHasGutschriftAmazon)
      {
          dfPreis    = pRechnung->m_dfBetragGutschriftWawi;
          ctPreis    = (int)(100.0 * dfPreis + 0.005);
          ctPreisCmp = (int)(100.0 * pRechnung->m_dfBetragGutschriftAmazon + 0.005);
          pRechnung->m_fAmazonDiffBetragGutschrift = ctPreis != ctPreisCmp;
          if (abs(ctPreis) != abs(ctPreisCmp))
          {
              dfGutDiff = (fabs(dfPreis) - fabs(pRechnung->m_dfBetragGutschriftAmazon));
              // szGuLine.Format("\"%s\";\"%s\";%5.2lf;%5.2lf\r\n", pRechnung->m_szRechnungsNr, pRechnung->m_bestellNummer, dfPreis, pRechnung->m_dfBetragGutschriftOrg);
          }
      }


      switch (pRechnung->m_typ)
      {
        case CRechnung::Amazon_DE:
          dfDiffReDe += dfRechDiff;
          dfDiffGuDe += dfGutDiff;
          // m_arrGuKorr.Add(szGuLine);
        break;
        case CRechnung::Amazon_Fr:
          dfDiffReFr += dfRechDiff;
          dfDiffGuFr += dfGutDiff;
        break;
        case CRechnung::Amazon_GB:
          dfDiffReGb += dfRechDiff;
          dfDiffGuGb += dfGutDiff;
        break;
      }


      // m_mapRechnung[szBestellnummer] = rechnung;

  }

  /*
  m_arrGuKorr.RemoveAll();
  CGutschrift* pGutschrift;
  
  szGuLine = "\"Rechnungsnummer\";\"Bestellnummer\";\"Gutschrift Amazon\";\"Gutschrift JTL\";\"Rechnungsbetrag JTL\"\r\n";
  m_arrGuKorr.Add(szGuLine);

  for (CMapGutschrift::CPair* pCurVal = m_mapGutschrift.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapGutschrift.PGetNextAssoc(pCurVal))
  {
      pGutschrift       = &pCurVal->value;
      szRechnungsnummer = pCurVal->key;

      if (pGutschrift->m_typ == CRechnung::Amazon_DE)
      {
          dfPreis = pGutschrift->m_dfGutschriftSumme;
          ctPreis = (int)(100.0 * dfPreis + 0.005);
          ctPreisCmp = (int)(100.0 * pGutschrift->m_dfBetragGutschriftOrg + 0.005);
          pGutschrift->m_fAmazonDiffBetragGutschrift = abs(ctPreis - ctPreisCmp) >= DELTA_CT_MIN;
          if (pGutschrift->m_fAmazonDiffBetragGutschrift)
          {
              dfGutDiff = (fabs(dfPreis) - fabs(pGutschrift->m_dfBetragGutschriftOrg));
              szGuLine.Format("\"%s\";\"%s\";%5.2lf;%5.2lf;%5.2lf\r\n", pGutschrift->m_rechnungsNummer, pGutschrift->m_extBestellNummer, dfPreis, pGutschrift->m_dfBetragGutschriftOrg, pGutschrift->m_dfRechnungsBetrag);
              szGuLine.Replace('.', ',');
              m_arrGuKorr.Add(szGuLine);
          }
      }



  }
  */

  // Ergebnis
  if (!fLineError)
  {
      szMessage.Format("Anzahl Erstattungen: %d\nAnzahl Nicht zugeordnet: %d\n\nAnzahl neuer Rechnungen: %d\nAnzahl Datensätze insgesamt: %d\n\nDifferenz Rechnungen : %5.2lf\nDifferenz Gutschrift : %5.2lf\n", countErstattung, countUnknown, countNeuRechnung, lineCount - 1, dfRechDiffTotal, dfGutDiffTotal);
      MessageBox(szMessage, "Info", MB_OK);

      szMessage.Format("Rechnungen:\nDe: %5.2lf\nFr: %5.2lf\nGb: %5.2lf\n", sumReDe/1.19, sumReFr / 1.19, sumReGb / 1.19);
      MessageBox(szMessage, "Info", MB_OK);

      szMessage.Format("Gutschriften:\nDe: %5.2lf\nFr: %5.2lf\nGb: %5.2lf\n", sumGuDe / 1.19, sumGuFr / 1.19, sumGuGb / 1.19);
      MessageBox(szMessage, "Info", MB_OK);

      szMessage.Format("Erstattungen:\nDe: %5.2lf\nFr: %5.2lf\nGb: %5.2lf\n", sumErDe / 1.19, sumErFr / 1.19, sumErGb / 1.19);
      MessageBox(szMessage, "Info", MB_OK);

      szMessage.Format("Differenz Rechnungen:\nDe: %5.2lf\nFr: %5.2lf\nGb: %5.2lf\n", dfDiffReDe / 1.19, dfDiffReFr / 1.19, dfDiffReGb / 1.19);
      MessageBox(szMessage, "Info", MB_OK);

      szMessage.Format("Differenz Gutschriften:\nDe: %5.2lf\nFr: %5.2lf\nGb: %5.2lf\n", dfDiffGuDe / 1.19, dfDiffGuFr / 1.19, dfDiffGuGb / 1.19);
      MessageBox(szMessage, "Info", MB_OK);

      /*
      CString szNewFilePath, szCsvLines;
      int     length;
      
      szNewFilePath = lpszFilePath;
      szCsvLines    = "";
      szNewFilePath.Replace(".csv", "_korr.csv");
      
      for (int i = 0; i < m_arrGuKorr.GetCount(); i++)
          szCsvLines += m_arrGuKorr[i];


      length = szCsvLines.GetLength();
      CFile fl;
      if (fl.Open(szNewFilePath, CFile::modeCreate | CFile::modeWrite))
      {
          fl.Write((LPCSTR)szCsvLines, length);
          fl.Close();
      }
      */

      m_storePath = lpszPath;
      m_storeTitle = szFileTitle;

      /*
      StoreResult(lpszPath, szFileTitle);
      ErstelleTaxStatistik(lpszPath, szFileTitle);
      ErstelleEasyCashRechnungen(lpszPath, szFileTitle);
      ErstelleEasyCashGutschriften(lpszPath, szFileTitle);

      CString szFilePath, szFile, szFileName, szLine;
      CFile   fl;

      szFilePath = lpszPath;
      szFile     = szFileTitle;
      szFile.Replace(".csv", "");
      szFileName.Format("%s\\%s_Summery.csv", szFilePath, szFile);

      if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
      {

          szLine = ECT_HEADER_SUMMERY;
          fl.Write((LPCSTR)szLine, szLine.GetLength());

          for (int i = 0; i < CRechnung::countPlatform; i++)
          {
              szLine = m_arrSummery[i];
              fl.Write((LPCSTR)szLine, szLine.GetLength());
          }
          fl.Close();

          szLine.Format("Die Summary-Datei\n%s\nwurde erzeugt.", szFileName);
          MessageBox(szLine, "Info", MB_OK);
      }


      ErstelleEasyCashKorrekturImport(lpszPath, szFileTitle);
      */

  }

  return !fLineError;

}

void CJTLSteuerDlg::StoreResults(void)
{
    CString szFilePath, szFile, szFileName, szLine;
    CFile   fl;

    StoreResult(m_storePath, m_storeTitle);
    ErstelleTaxStatistik(m_storePath, m_storeTitle);
    ErstelleEasyCashRechnungen(m_storePath, m_storeTitle);
    ErstelleEasyCashPOS(m_storePath, m_storeTitle);
    ErstelleEasyCashGutschriften(m_storePath, m_storeTitle);

    szFilePath = m_storePath;
    szFile = m_storeTitle;
    szFile.Replace(".csv", "");
    szFileName.Format("%s\\%s_Summery.csv", szFilePath, szFile);

    if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
    {

        szLine = ECT_HEADER_SUMMERY;
        fl.Write((LPCSTR)szLine, szLine.GetLength());

        for (int i = 0; i < CRechnung::countPlatform; i++)
        {
            szLine = m_arrSummery[i];
            fl.Write((LPCSTR)szLine, szLine.GetLength());
        }
        fl.Close();

        szLine.Format("Die Summary-Datei\n%s\nwurde erzeugt.", szFileName);
        MessageBox(szLine, "Info", MB_OK);
    }


    ErstelleEasyCashKorrekturImport(m_storePath, m_storeTitle);

    // OSS-Daten ...
    CString szNewFilePath;
    szNewFilePath.Format("%s\\%s_OSS_Final.csv", m_storePath, m_storeTitle);
    m_ossData.WriteData(this, szNewFilePath);

}


bool CJTLSteuerDlg::IsOSSDate(LPCSTR lpszDatum)
{
  // 01234567890
  // tt.mm.jjjj
  bool fReturn       = true;
  bool fIsDatumValid = strlen(lpszDatum) == 10;
  if (fIsDatumValid)    
  {
    CString datum(lpszDatum);
    fReturn = atoi(datum.Mid(6)) >= 2022 || (atoi(datum.Mid(6)) >= 2021 && atoi(datum.Mid(3, 2)) >= 7);
  }
  return fReturn;
}

bool CJTLSteuerDlg::CheckOSS(LPCSTR lpszISO, LPCSTR lpszUmsatz, LPCSTR lpszDatum)
{
  CString szIso   = lpszISO;
  bool    fReturn = false;
  if (szIso != "DE" && IsEU(szIso) && IsOSSDate(lpszDatum))
  {
    double fUmst = atof(lpszUmsatz);
    int    umst;
    if (fUmst < 1.0)
        umst = (int)(100.0 * fUmst + 0.005);
    else
        umst = (int)(fUmst + 0.005);
    fReturn = 0 != umst && 19 != umst;
    if (!fReturn)
      fReturn = szIso == "RO" || szIso == "CY";
  }
  return fReturn;
}

bool CJTLSteuerDlg::CheckOSS(CGutschrift* pGutschrift) 
{ 
    bool fReturn = pGutschrift->m_szUmst.GetLength() > 0 && pGutschrift->m_ISO.GetLength() > 0 && CheckOSS(pGutschrift->m_ISO, pGutschrift->m_szUmst, pGutschrift->m_szDatum); 
    if (pGutschrift->m_fAmazonTaxReport && fReturn != pGutschrift->m_fAmazonOSS)
    {
        CString szMessage;
        CString szAmazonOSS = (LPCSTR)(pGutschrift->m_fAmazonOSS ? "Ja" : "Nein");
        CString szInternOSS = (LPCSTR)(fReturn                   ? "Ja" : "Nein");
        szMessage.Format("Widerspruch bei Rechnung bezüglich OSS-Zuordnung.\n\nRechnungsnummer: %s\nAmazon-Zuordnung OSS: %s\nInterne Zuordnung OSS: %s", pGutschrift->m_gutschriftNummer, szAmazonOSS, szInternOSS);
        MessageBox(szMessage, "Gutschriften", MB_ICONEXCLAMATION | MB_OK);
    }
    return fReturn;
}

bool CJTLSteuerDlg::CheckOSS(CRechnung* pRechnung) 
{ 
    bool fReturn = pRechnung->m_szUmst.GetLength() > 0 && pRechnung->m_ISO.GetLength() > 0 && CheckOSS(pRechnung->m_ISO, pRechnung->m_szUmst, pRechnung->m_szDatum);
    if (pRechnung->m_fAmazonTaxReport && fReturn != pRechnung->m_fAmazonOSS)
    {
        CString szMessage;
        CString szAmazonOSS = (LPCSTR)(pRechnung->m_fAmazonOSS ? "Ja" : "Nein");
        CString szInternOSS = (LPCSTR)(fReturn                 ? "Ja" : "Nein");
        szMessage.Format("Widerspruch bei Rechnung bezüglich OSS-Zuordnung.\n\nRechnungsnummer: %s\nAmazon-Zuordnung OSS: %s\nInterne Zuordnung OSS: %s", pRechnung->m_szRechnungsNr, szAmazonOSS, szInternOSS);
        MessageBox(szMessage, "Rechnungen", MB_ICONEXCLAMATION | MB_OK);
    }
    return fReturn;
}

void CJTLSteuerDlg::ErstelleEasyCashPOS(LPCSTR lpszPath, LPCSTR szFileTitle)
{
  CString szTitle, szFileName, arrRechnungPOS, szLine;
  CString szRBNummer;
  POSDATA posSave;
  CArray <POSDATA, POSDATA> arrPosData;
  CFile   fl;
  long    length, lBrutto;
  bool    fSort;
  int     i, countPos, lineCount;

  szTitle = szFileTitle;
  szTitle.Replace(".csv", "");
  
  szFileName.Format("%s\\%s_Rechnung_POS.csv", lpszPath, szTitle);
  
  arrRechnungPOS = ECT_RECHNUNG_TITEL;
  lineCount      = 0;
  lBrutto        = 0;
  
  // "Buchungsart";"Erstelldatum Rechnung";"Gesamtbetrag Brutto (alle Ust.)";"RA Vorname RA Nachname";"Rechnungsnummer";"Konto";"USt.";"Währung";"Währungsfaktor"
  for (CMapPOSData::CPair* pCurVal = m_mapPOSData.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapPOSData.PGetNextAssoc(pCurVal))
    arrPosData.Add(pCurVal->value);

  // Sortiere ...
  countPos = arrPosData.GetCount();
  fSort = true;
  while (fSort)
  {
    fSort = false;
    for (i = 0; i < countPos - 1; i++)
    {
      fSort = arrPosData[i + 1].rbNummer < arrPosData[i].rbNummer;
      if (fSort)
      {
        posSave         = arrPosData[i];
        arrPosData[i]   = arrPosData[i+1];
        arrPosData[i+1] = posSave;
      }
    }
  }

  // Ausgabe ...
  for (i=0; i< countPos; i++)
  {
    
    szRBNummer.Format("RB-%04d", arrPosData[i].rbNummer);
    lBrutto += arrPosData[i].nBruttonVK;
    szLine.Format("\r\n\"Einnahmen\";\"%s\";\"%3.2f\";\"%s\";\"%s\";\"Kasse\";\"%s\";\"EUR\";\"1\"",
      arrPosData[i].szDatum, (double)arrPosData[i].nBruttonVK / 100.0, arrPosData[i].szKundenNummer, szRBNummer, arrPosData[i].szUst);

    /*
    lBrutto += pCurVal->value.nBruttonVK;
    szLine.Format("\r\n\"Einnahmen\";\"%s\";\"%3.2f\";\"%s\";\"%s\";\"Kasse\";\"%s\";\"EUR\";\"1\"",
                   pCurVal->value.szDatum, (double)pCurVal->value.nBruttonVK / 100.0, pCurVal->value.szKundenNummer, pCurVal->key, pCurVal->value.szUst);
    arrRechnungPOS += szLine;
    */
    lineCount      += 1;
  }

  
  length = arrRechnungPOS.GetLength();

  if (lineCount && fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
  {
    fl.Write((LPCSTR)arrRechnungPOS, length);
    fl.Close();

    CString szMessage;
    szMessage.Format("Gesamtwert der Kassen-Einnahmen: %3.2lf EUR", (double)lBrutto / 100.0);
    MessageBox(szMessage, "Kasse", MB_OK);
  }

}

void CJTLSteuerDlg::ErstelleEasyCashRechnungen(LPCSTR lpszPath, LPCSTR szFileTitle)
{

    CMap<CString, LPCSTR, int, int>  mapPlatform;
    CKeyArray                        arrRechnungen[CRechnung::countPlatform];

    LPCSTR     pszRechnungstitel = ECT_RECHNUNG_TITEL;
    LPCSTR     arrBuchungsart[CRechnung::countPlatform] = ECT_RECHNUNG_KONTO;
    LPCSTR     arrFileTyp[CRechnung::countPlatform] = PLATFORM_ARR_OUT;
    CString    arrechnungPlatform[CRechnung::countPlatform];
    CString    szLine, szKonto;
    int        countPlatForm[CRechnung::countPlatform];
    int        countOSSNetto[COSSIndex::pl_count],
               countOSSUmst[COSSIndex::pl_count];
    double     arrRechnungTotal[CRechnung::countPlatform], fUmst;

    CString    ossEinnahmenNetto[COSSIndex::pl_count],
               ossEinnahmenUmst[COSSIndex::pl_count];

    CRechnung* pRechnung;
    int        i, j, indexPlatform, countPl, tag;

    CFile      fl;
    CString    szMessage, szFileName, SortKey;
    long       length;

    mapPlatform["Amazon.de"] = CRechnung::Amazon_DE;
    mapPlatform["Amazon.co.uk"] = CRechnung::Amazon_GB;
    mapPlatform["Amazon.fr"] = CRechnung::Amazon_Fr;
    mapPlatform["Onlineshop"] = CRechnung::OnlineShop;
    mapPlatform["JTL-Wawi"] = CRechnung::JTL_WaWi;
    mapPlatform["Händler"] = CRechnung::Haendler;
    mapPlatform["UnKnown"] = CRechnung::unknown;

    szLine = ECT_RECHNUNG_TITEL;
    for (i = 0; i < CRechnung::countPlatform; i++)
    {
        countPlatForm[i] = 0;
        arrechnungPlatform[i] = szLine;
        arrRechnungTotal[i] = 0;
    }
    
    for (i = 0; i < COSSIndex::pl_count; i++)
    {
      ossEinnahmenNetto[i] = szLine;
      ossEinnahmenUmst[i]  = szLine;
      countOSSNetto[i]     = 0;
      countOSSUmst[i]      = 0;
    }

    for (CMapRechnung::CPair* pCurVal = m_mapRechnung.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapRechnung.PGetNextAssoc(pCurVal))
    {
        pRechnung = &pCurVal->value;
        if (!mapPlatform.Lookup(pRechnung->m_platform, indexPlatform))
            indexPlatform = CRechnung::unknown;
        tag = GetSimpleDay(pRechnung->m_szDatum);
        SortKey.Format("%d_%s", tag, pRechnung->m_szRechnungsNr);
        //arrRechnungen[indexPlatform].Add({ pCurVal->key, pRechnung->m_szRechnungsNr });
        arrRechnungen[indexPlatform].Add({ pCurVal->key, SortKey });
    }

    COSSIndex ossIndex;
    CRechnung rech;
    double    dfNetto, dfUmst;
    CString   szText, zahlUmst, zahlFaktor;
    int       umSt, indexOSS;

    for (i = 0; i < CRechnung::countPlatform; i++)
    {
        DoSort(arrRechnungen[i], OrderString);
        countPl = arrRechnungen[i].GetCount();
        for (j = 0; j < countPl; j++)
        {
            if (m_mapRechnung.Lookup(arrRechnungen[i][j].szKey, rech))
            {

                pRechnung = &rech;
                fUmst = atof(pRechnung->m_szUmst);
                // if (pRechnung->m_ISO.GetLength() > 0 && pRechnung->m_ISO != "DE")
                if (CheckOSS(pRechnung))
                {
                    
                    m_ossData.AddLine(pRechnung);
                    
                    umSt = (int)((100.0 * fUmst) + 0.05);
                    szText.Format("%s (%s %s %%)", pRechnung->m_szRechnungEmpf, pRechnung->m_ISO, pRechnung->m_szUmst);
                    if (0 != umSt)
                    {
                        dfNetto = pRechnung->m_dfBetragRechnung / ( 1.0 + fUmst/100.0 );
                        dfUmst  = pRechnung->m_dfBetragRechnung - dfNetto;

                    }
                    else
                    {
                        dfNetto = pRechnung->m_dfBetragRechnung;
                        dfUmst  = 0;
                        
                    }
                    
                    // szKonto = KONTO_OSS_EINNAHMEN_NETTO;

                    switch (i)
                    {
                    case CRechnung::Amazon_DE:
                      indexOSS = COSSIndex::pl_amazonDE;
                    break;
                    case CRechnung::Amazon_Fr:
                      indexOSS = COSSIndex::pl_amazonFR;
                    break;
                    case CRechnung::OnlineShop:
                    case CRechnung::JTL_WaWi:
                    case CRechnung::Haendler:
                      indexOSS = COSSIndex::pl_webshop;
                    break;
                    case CRechnung::Amazon_GB:
                    case CRechnung::unknown:
                      indexOSS = COSSIndex::pl_amazonRest;
                    break;
                    }

                    szKonto = ossIndex.GetTitle((COSSIndex::tagPlatform)indexOSS, COSSIndex::art_Rechnung, COSSIndex::betrag_Netto);
                    szLine.Format(ECT_RE_GU_FORM, pRechnung->m_szDatum, dfNetto, szText, pRechnung->m_szRechnungsNr, szKonto, "0", pRechnung->m_waehrungRechnung, pRechnung->m_faktor);
                    
                    ossEinnahmenNetto[indexOSS] += szLine;
                    countOSSNetto[indexOSS]     += 1;

                    if (0 != umSt)
                    {
                        zahlUmst.Format("%3.2f", dfUmst);
                        zahlUmst.Replace(".", ",");
                        
                        zahlFaktor.Format("%1.6f", pRechnung->m_faktor);
                        zahlFaktor.Replace(".", ",");
                        
                        // szKonto = KONTO_OSS_EINNAHMEN_UMST;
                        szKonto = ossIndex.GetTitle((COSSIndex::tagPlatform)indexOSS, COSSIndex::art_Rechnung, COSSIndex::betrag_Umst);
                        szLine.Format(ECT_RE_GU_FORM_EXCEL, pRechnung->m_szDatum, zahlUmst, szText, pRechnung->m_szRechnungsNr, "0", pRechnung->m_szUmst, pRechnung->m_waehrungRechnung, zahlFaktor);
                        ossEinnahmenUmst[indexOSS] += szLine;
                        countOSSUmst[indexOSS]     += 1;
                    }

                }
                else
                {
                    szKonto = arrBuchungsart[i];
                    if (0 == (int)((100.0 * fUmst) + 0.05))
                        szKonto += " (0%)";
                    szLine.Format(ECT_RE_GU_FORM, pRechnung->m_szDatum, pRechnung->m_dfBetragRechnung, pRechnung->m_szRechnungEmpf,
                                                  pRechnung->m_szRechnungsNr, szKonto, pRechnung->m_szUmst,
                                                  pRechnung->m_waehrungRechnung, pRechnung->m_faktor);
                    countPlatForm[i] += 1;
                    arrechnungPlatform[i] += szLine;
                    
                }
                arrRechnungTotal[i] += pRechnung->m_dfBetragRechnung;

            }
        }
    }

    /*
    for (CMapRechnung::CPair* pCurVal = m_mapRechnung.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapRechnung.PGetNextAssoc(pCurVal))
    {
        pRechnung = &pCurVal->value;
        if (!mapPlatform.Lookup(pRechnung->m_platform, indexPlatform))
            indexPlatform = CRechnung::unknown;
        szLine.Format(ECT_RE_GU_FORM, pRechnung->m_szDatum, pRechnung->m_dfBetragRechnung, pRechnung->m_szRechnungEmpf, 
                                         pRechnung->m_szRechnungsNr, arrBuchungsart[indexPlatform], pRechnung->m_szUmst, 
                                         pRechnung->m_waehrungRechnung, pRechnung->m_faktor);
        countPlatForm[indexPlatform]      += 1;
        arrechnungPlatform[indexPlatform] += szLine;
        arrRechnungTotal[indexPlatform]   += pRechnung->m_dfBetragRechnung;

    }
    */
    
    CString szTitle = szFileTitle;
    szTitle.Replace(".csv", "");

    for (i = 0; i < CRechnung::countPlatform; i++)
    {
        if (countPlatForm[i])
        {
            szFileName.Format("%s\\%s_Rechnung_%s.csv", lpszPath, szTitle, arrFileTyp[i]);
            length = arrechnungPlatform[i].GetLength();
            if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
            {
                fl.Write((LPCSTR)arrechnungPlatform[i], length);
                fl.Close();
            }

            
            szLine.Format("\r\n\"%s\";\"%d\";\"%5.2lf\";", arrFileTyp[i], countPlatForm[i], arrRechnungTotal[i]);
            szLine.Replace(".", ",");
            m_arrSummery[i] += szLine;

            szMessage.Format("Es wurde die Rechnungsdatei\n%s\nmit %d Datensätzen erzeugt.\n", szFileName, countPlatForm[i]);
            MessageBox(szMessage, "Info", MB_OK);
        }
        else
        {
            szLine.Format("\r\n\"%s\";\"%d\";\"%5.2lf\";", arrFileTyp[i], 0, (double)0.0);
            szLine.Replace(".", ",");
            m_arrSummery[i] += szLine;
        }
    }

    
    // Nun noch die OSS-Rechnungen ..
    CString szFilePath;
    for (i = 0; i < COSSIndex::pl_count; i++)
    {
      if (countOSSNetto[i])
      {
        szFileName = ossIndex.GetFileTitle((COSSIndex::tagPlatform)i, COSSIndex::art_Rechnung, COSSIndex::betrag_Netto);
        szFilePath.Format("%s\\%s", lpszPath, szFileName);
        length = ossEinnahmenNetto[i].GetLength();
        if (fl.Open(szFilePath, CFile::modeCreate | CFile::modeWrite))
        {
          fl.Write((LPCSTR)ossEinnahmenNetto[i], length);
          fl.Close();
        }
      }

      if (countOSSUmst[i])
      {
        szFileName = ossIndex.GetFileTitle((COSSIndex::tagPlatform)i, COSSIndex::art_Rechnung, COSSIndex::betrag_Umst);
        szFilePath.Format("%s\\%s", lpszPath, szFileName);
        length = ossEinnahmenUmst[i].GetLength();
        if (fl.Open(szFilePath, CFile::modeCreate | CFile::modeWrite))
        {
          fl.Write((LPCSTR)ossEinnahmenUmst[i], length);
          fl.Close();
        }
      }
    }

    /*
    szFileName.Format("%s\\%s_OSS_Rechnung_Netto.csv", lpszPath, szTitle);
    length = ossEinnahmenNetto.GetLength();
    if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
    {
        fl.Write((LPCSTR)ossEinnahmenNetto, length);
        fl.Close();
    }

    szFileName.Format("%s\\%s_OSS_Rechnung_Umst.csv", lpszPath, szTitle);
    length = ossEinnahmenUmSt.GetLength();
    if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
    {
        fl.Write((LPCSTR)ossEinnahmenUmSt, length);
        fl.Close();
    }
    */

}

void CJTLSteuerDlg::PruefeGutschriftenSumme(void)
{
    CRechnung* pRechnung;
    CString    szMessage;
    double     dfTotal, dfRechnung, dfKontroll;
    int        countTotal, countBad;
    
    dfTotal    = 0;
    countTotal = 0;
    countBad   = 0;
    for (CMapRechnung::CPair* pCurVal = m_mapRechnung.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapRechnung.PGetNextAssoc(pCurVal))
    {
        pRechnung   = &pCurVal->value;
        countTotal += 1;
        dfRechnung = fabs(pRechnung->m_dfBetragRechnung);
        dfKontroll = fabs(pRechnung->m_dfKontrollSumme);
        if (dfKontroll - dfRechnung > 0.5)
        {
            dfTotal  += dfRechnung - dfKontroll;
            countBad += 1;
        }
    }

    szMessage.Format("Auswertung:\n\nNicht gefundene Gutschriften: %d von %d\n\nAnzahl fehlerhafter Gutschriften (bezogen auf Rechnungen): %d von %d\n\nGesamtbetrag: %3.2f\n",
                      m_GutschriftNotFound, m_GutschriftNotFound + m_GutschriftFound, countBad, countBad + countTotal, dfTotal);
    MessageBox(szMessage, "Auswertung der Gutschriften", MB_OK);
}

bool CJTLSteuerDlg::AddGutschriftZuRechnung(CGutschrift* pGutschrift, double *pBetrag)
{
    // 1. Finde die Rechnung ...
    CRechnung rechnung;
    double    delta, betragGutschrift, betragRechnung, betragKontrollSumme;
    int       sign;
    bool      fReturn = true;
    if (m_mapRechnung.Lookup(pGutschrift->m_extBestellNummer, rechnung))
    {
        m_GutschriftFound  += 1;
        betragGutschrift    = fabs(pGutschrift->m_dfBetragGutschriftOrg);
        betragRechnung      = fabs(rechnung.m_dfBetragRechnung);
        betragKontrollSumme = fabs(rechnung.m_dfKontrollSumme);
        sign                = pGutschrift->m_dfBetragGutschriftOrg > 0 ? 1 : -1;
        delta               = betragRechnung - betragKontrollSumme;
        if (delta >= betragGutschrift)
        {
            *pBetrag = betragGutschrift * sign;
        }
        else if (delta > 0)
        {
            *pBetrag = delta * sign;
        }
        else
        { 
            *pBetrag = 0;
            fReturn = false;
        }

        rechnung.m_dfKontrollSumme += pGutschrift->m_dfBetragGutschriftOrg;
        m_mapRechnung[pGutschrift->m_extBestellNummer] = rechnung;
    }
    else
    {
        m_GutschriftNotFound += 1;
        *pBetrag = pGutschrift->m_dfBetragGutschriftOrg;
    }
    return fReturn;
}

void CJTLSteuerDlg::ErstelleEasyCashGutschriften(LPCSTR lpszPath, LPCSTR szFileTitle)
{

    CMap<CString, LPCSTR, int, int>  mapPlatform;
    CKeyArray                        arrGutschriften[CRechnung::countPlatform];

    LPCSTR       pszGutschrifttitel = ECT_GUTSCHRIFT_TITEL;
    LPCSTR       arrBuchungsart[CRechnung::countPlatform] = ECT_GUTSCHRIFT_KONTO;
    LPCSTR       arrFileTyp[CRechnung::countPlatform]     = PLATFORM_ARR_OUT;

    CString      arrechnungPlatform[CRechnung::countPlatform];
    CString      szLine;
    double       arrGutschriftTotal[CRechnung::countPlatform];
    double       betragGutschrift;
    int          countPlatForm[CRechnung::countPlatform];

    CGutschrift* pGutschrift;
    int          i, j, countPl, indexPlatform;

    CFile        fl;
    CString      szMessage, szFileName, szKonto;
    double       fUmst;
    long         length;

    // CString      ossGutschriftNetto, ossGutschriftUmSt;
    CString      ossGutschriftNetto[COSSIndex::pl_count],
                 ossGutschriftUmst[COSSIndex::pl_count];

    int          countGutschriftNetto[COSSIndex::pl_count],
                 countGutschriftUmst[COSSIndex::pl_count];


    mapPlatform["Amazon.de"]    = CRechnung::Amazon_DE;
    mapPlatform["Amazon.co.uk"] = CRechnung::Amazon_GB;
    mapPlatform["Amazon.fr"]    = CRechnung::Amazon_Fr;
    mapPlatform["Onlineshop"]   = CRechnung::OnlineShop;
    mapPlatform["JTL-Wawi"]     = CRechnung::JTL_WaWi;
    mapPlatform["Händler"]      = CRechnung::Haendler;
    mapPlatform["UnKnown"]      = CRechnung::unknown;

    szLine = ECT_GUTSCHRIFT_TITEL;
    for (i = 0; i < CRechnung::countPlatform; i++)
    {
        countPlatForm[i]      = 0;
        arrechnungPlatform[i] = szLine;
        arrGutschriftTotal[i] = 0;
    }

    for (i = 0; i < COSSIndex::pl_count; i++)
    {
      ossGutschriftNetto[i]   = szLine;
      ossGutschriftUmst[i]    = szLine;
      countGutschriftNetto[i] = 0;
      countGutschriftUmst[i]  = 0;
    }

    for (CMapGutschrift::CPair* pCurVal = m_mapGutschrift.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapGutschrift.PGetNextAssoc(pCurVal))
    {
        pGutschrift = &pCurVal->value;
        if (!mapPlatform.Lookup(pGutschrift->m_platform, indexPlatform))
            indexPlatform = CRechnung::unknown;
        arrGutschriften[indexPlatform].Add({ pCurVal->key, pGutschrift->m_gutschriftNummer});
    }

    CGutschrift gut;
    COSSIndex   ossIndex;
    double      dfNetto, dfUmst;
    CString     szText, zahlUmst, zahlFaktor;;
    int         umSt, indexOSS;

    for (i = 0; i < CRechnung::countPlatform; i++)
    {
        DoSort(arrGutschriften[i], OrderRechnung);
        countPl = arrGutschriften[i].GetCount();
        for (j = 0; j < countPl; j++)
        {
            if (m_mapGutschrift.Lookup(arrGutschriften[i][j].szKey, gut))
            {
                pGutschrift = &gut;
                if (AddGutschriftZuRechnung(pGutschrift, &betragGutschrift))
                {
                    
                    if (pGutschrift->m_ISO.GetLength() > 0 && pGutschrift->m_szUmst.GetLength() > 0)
                      szText.Format("%s (%s %s %%)", pGutschrift->m_empfaenger, pGutschrift->m_ISO, pGutschrift->m_szUmst);
                    else
                      szText = pGutschrift->m_empfaenger;

                    fUmst = atof(pGutschrift->m_szUmst);
                    // if (pGutschrift->m_ISO.GetLength() > 0 && pGutschrift->m_ISO != "DE")
                    if (CheckOSS(pGutschrift))
                    {
                      
                        m_ossData.AddLine(pGutschrift);
                        
                        umSt = (int)((100.0 * fUmst) + 0.05);
                        if (0 != umSt)
                        {
                            dfNetto = betragGutschrift / (1.0 + fUmst/100.0);
                            dfUmst  = betragGutschrift - dfNetto;

                        }
                        else
                        {
                            dfNetto = betragGutschrift;
                            dfUmst = 0;

                        }
                        
                        // szKonto = KONTO_OSS_GUTSCHRIFTEN_NETTO;

                        switch (i)
                        {
                        case CRechnung::Amazon_DE:
                          indexOSS = COSSIndex::pl_amazonDE;
                        break;
                        case CRechnung::Amazon_Fr:
                          indexOSS = COSSIndex::pl_amazonFR;
                        break;
                        case CRechnung::OnlineShop:
                        case CRechnung::JTL_WaWi:
                        case CRechnung::Haendler:
                          indexOSS = COSSIndex::pl_webshop;
                        break;
                        case CRechnung::Amazon_GB:
                        case CRechnung::unknown:
                          indexOSS = COSSIndex::pl_amazonRest;
                        break;
                        }

                        szKonto = ossIndex.GetTitle((COSSIndex::tagPlatform)indexOSS, COSSIndex::art_Gutschrift, COSSIndex::betrag_Netto);

                        szLine.Format(ECT_RE_GU_FORM, pGutschrift->m_szDatum, dfNetto, szText, pGutschrift->m_gutschriftNummer, szKonto, "0.00", pGutschrift->m_waehrungGutschrift, pGutschrift->m_faktorRechnung);
                        ossGutschriftNetto[indexOSS]   += szLine;
                        countGutschriftNetto[indexOSS] += 1;

                        if (0 != umSt)
                        {
                            zahlUmst.Format("%3.2f", dfUmst);
                            zahlUmst.Replace(".", ",");

                            zahlFaktor.Format("%1.6f", pGutschrift->m_faktorRechnung);
                            zahlFaktor.Replace(".", ",");

                            // szKonto = KONTO_OSS_GUTSCHRIFTEN_UMST;
                            szKonto = ossIndex.GetTitle((COSSIndex::tagPlatform)indexOSS, COSSIndex::art_Gutschrift, COSSIndex::betrag_Umst);
                            szLine.Format(ECT_RE_GU_FORM_EXCEL, pGutschrift->m_szDatum, zahlUmst, szText, pGutschrift->m_gutschriftNummer, szKonto, "0,00", pGutschrift->m_waehrungGutschrift, zahlFaktor);
                            ossGutschriftUmst[indexOSS]   += szLine;
                            countGutschriftUmst[indexOSS] += 1;
                        }

                    }
                    else
                    { 
                        szKonto = arrBuchungsart[i];
                        if (0 == (int)((100.0 * fUmst) + 0.05) )
                            szKonto += " (0%)";
                        // szLine.Format(ECT_RE_GU_FORM, pGutschrift->m_szDatum, betragGutschrift, pGutschrift->m_empfaenger, pGutschrift->m_gutschriftNummer, szKonto, pGutschrift->m_szUmst, pGutschrift->m_waehrungGutschrift, pGutschrift->m_faktorRechnung);
                        szLine.Format(ECT_RE_GU_FORM, pGutschrift->m_szDatum, betragGutschrift, szText, pGutschrift->m_gutschriftNummer, szKonto, pGutschrift->m_szUmst, pGutschrift->m_waehrungGutschrift, pGutschrift->m_faktorRechnung);
                        countPlatForm[i] += 1;
                        arrechnungPlatform[i] += szLine;
                    }
                    arrGutschriftTotal[i] += betragGutschrift; // pGutschrift->m_dfBetragGutschriftOrg;
                }
            }
        }
    }


    PruefeGutschriftenSumme();
    /*
    for (CMapGutschrift::CPair* pCurVal = m_mapGutschrift.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapGutschrift.PGetNextAssoc(pCurVal))
    {
        pGutschrift = &pCurVal->value;
        if (!mapPlatform.Lookup(pGutschrift->m_platform, indexPlatform))
            indexPlatform = CRechnung::unknown;
        szLine.Format(ECT_RE_GU_FORM, pGutschrift->m_szDatum, pGutschrift->m_dfBetragGutschriftOrg, pGutschrift->m_empfaenger,
            pGutschrift->m_gutschriftNummer, arrBuchungsart[indexPlatform], pGutschrift->m_szUmst,
            pGutschrift->m_waehrungGutschrift, pGutschrift->m_faktor);
        countPlatForm[indexPlatform] += 1;
        arrechnungPlatform[indexPlatform] += szLine;
        arrGutschriftTotal[indexPlatform] += pGutschrift->m_dfBetragGutschriftOrg;
    }
    */

    CString szTitle = szFileTitle;
    szTitle.Replace(".csv", "");

    for (i = 0; i < CRechnung::countPlatform; i++)
    {
        if (countPlatForm[i])
        {
            szFileName.Format("%s\\%s_Gutschrift_%s.csv", lpszPath, szTitle, arrFileTyp[i]);
            length = arrechnungPlatform[i].GetLength();
            if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
            {
                fl.Write((LPCSTR)arrechnungPlatform[i], length);
                fl.Close();
            }

            szLine.Format("\"%d\";\"%5.2lf\"", countPlatForm[i], arrGutschriftTotal[i]);
            szLine.Replace(".", ",");
            m_arrSummery[i] += szLine;

            szMessage.Format("Es wurde die Gutschriftdatei\n%s\nmit %d Datensätzen erzeugt.\n", szFileName, countPlatForm[i]);
            MessageBox(szMessage, "Info", MB_OK);
        }
        else
        {
            szLine = "\"0\";\"0,00\"";
            m_arrSummery[i] += szLine;
        }
    }

    // Nun noch die OSS-Gutschriften ..
    CString szFilePath;
    for (i = 0; i < COSSIndex::pl_count; i++)
    {
      if (countGutschriftNetto[i])
      {
        szFileName = ossIndex.GetFileTitle((COSSIndex::tagPlatform)i, COSSIndex::art_Gutschrift, COSSIndex::betrag_Netto);
        szFilePath.Format("%s\\%s", lpszPath, szFileName);
        length = ossGutschriftNetto[i].GetLength();
        if (fl.Open(szFilePath, CFile::modeCreate | CFile::modeWrite))
        {
          fl.Write((LPCSTR)ossGutschriftNetto[i], length);
          fl.Close();
        }
      }

      if (countGutschriftUmst[i])
      {
        szFileName = ossIndex.GetFileTitle((COSSIndex::tagPlatform)i, COSSIndex::art_Gutschrift, COSSIndex::betrag_Umst);
        szFilePath.Format("%s\\%s", lpszPath, szFileName);
        length = ossGutschriftUmst[i].GetLength();
        if (fl.Open(szFilePath, CFile::modeCreate | CFile::modeWrite))
        {
          fl.Write((LPCSTR)ossGutschriftUmst[i], length);
          fl.Close();
        }
      }
    }

    /*
    szFileName.Format("%s\\%s_OSS_Gutschriften_Netto.csv", lpszPath, szTitle);
    length = ossGutschriftNetto.GetLength();
    if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
    {
        fl.Write((LPCSTR)ossGutschriftNetto, length);
        fl.Close();
    }

    szFileName.Format("%s\\%s_OSS_Gutschriften_Umst.csv", lpszPath, szTitle);
    length = ossGutschriftUmSt.GetLength();
    if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
    {
        fl.Write((LPCSTR)ossGutschriftUmSt, length);
        fl.Close();
    }
    */

}


#define DELTA_F  0.0000005
#define NINT(x)  (x < 0 ? (int)(x-DELTA_F) : (x > 0 ? (int)(x+DELTA_F) : 0))
#define DEZIMALTRENNER        '.'
#define DEZIMALTRENNER_EXCEL  ','

void CJTLSteuerDlg::ErstelleEasyCashKorrekturImport(LPCSTR lpszPath, LPCSTR szFlTitle)
{
    CRechnung      rech;
    CString        szFormat, szDate, szLine, szRechnungsnummer, szFileName, szFileTitle, szUmst, szWaehrung;
    long           ctAmazonGutschrift, ctPreisRechnung;
    int            i, countExport, index, faktorVK, faktorNK;

    LPCSTR         pszGutschriftPre[ECTK_COUNT_BUCHUNG] = ECTK_GUTSCHRIFT_PRE;
    LPCSTR         pszBuchungsArt[ECTK_COUNT_BUCHUNG]   = ECTK_BUCHUNGSART;
    LPCSTR         pszFileExt[ECTK_COUNT_BUCHUNG]       = ECTK_FILE_EXT;
    CString        szExportLines[ECTK_COUNT_BUCHUNG];
    int            countExportLines[ECTK_COUNT_BUCHUNG];

    LPCSTR         pszErstattungPre[ECTK_COUNT_BUCHUNG]         = ECTK_ERSTATTUNG_PRE;
    LPCSTR         pszBuchungsArtErstattung[ECTK_COUNT_BUCHUNG] = ECTK_BUCHUNG_ERST;
    LPCSTR         pszFileExtErstattung[ECTK_COUNT_BUCHUNG]     = ECTK_FILE_ERST_EXT;
    CString        szExportLinesErstattung[ECTK_COUNT_BUCHUNG];
    int            countErstattungLines[ECTK_COUNT_BUCHUNG];

    CString        szKonto, ossRueckNetto[COSSIndex::pl_count], ossRueckUmst[COSSIndex::pl_count];
    int            indexOSS, countRueckNetto[COSSIndex::pl_count], countRueckUmst[COSSIndex::pl_count];

    CFile          fl;
    long           length;
    
    CMap<CString, LPCSTR, int, int>  mapPlatform;
    CRechnung*                       pRechnung;
    CKeyArray                        arrRechnungen[CRechnung::countPlatform];
    COSSIndex                        ossIndex;
    int                              indexPlatform, countPl, j;

    // Initialisierung
    szFormat      = "\r\n\"Einnahmen\";\"%s\";\"%0d%c%02d\";\"%s\";\"%s%04d\";\"%s\";\"%s\";\"%s\";\"%0d%c%05d\";\"%s\";\"%s\";\"0\";\"0\"";
    szLine        = "\"Buchungsart\";\"Erstelldatum\";\"Gesamtbetrag Brutto (alle Ust.)\";\"RA Vorname RA Nachname\";\"Gutschriftsnummer\";\"Konto\";\"USt.\";\"Währung\";\"Währungsfaktor\";\"Bezug Rechnungsnummer\";\"Externe Bestellnummer\";\"Debitorennummer\";\"Buchungskonto\"";

    for (index = 0; index < ECTK_COUNT_BUCHUNG; index++)
    {
        szExportLines[index]           = szLine;
        szExportLinesErstattung[index] = szLine;
        countExportLines[index]        = 0;
        countErstattungLines[index]    = 0;
    }

    mapPlatform["Amazon.de"] = CRechnung::Amazon_DE;
    mapPlatform["Amazon.co.uk"] = CRechnung::Amazon_GB;
    mapPlatform["Amazon.fr"] = CRechnung::Amazon_Fr;
    mapPlatform["Onlineshop"] = CRechnung::OnlineShop;
    mapPlatform["JTL-Wawi"] = CRechnung::JTL_WaWi;
    mapPlatform["UnKnown"] = CRechnung::unknown;

    for (i = 0; i < COSSIndex::pl_count; i++)
    {
      ossRueckNetto[i]   = szLine;
      ossRueckUmst[i]    = szLine;
      countRueckNetto[i] = 0;
      countRueckUmst[i]  = 0;
    }

    for (CMapRechnung::CPair* pCurVal = m_mapRechnung.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapRechnung.PGetNextAssoc(pCurVal))
    {
        pRechnung = &pCurVal->value;
        if (!mapPlatform.Lookup(pRechnung->m_platform, indexPlatform))
            indexPlatform = CRechnung::unknown;
        arrRechnungen[indexPlatform].Add({ pCurVal->key, pRechnung->m_szRechnungsNr });
    }

    long ctPreisGutWawi, ctPreisGutAmazon, ctPreisDiff;

    countExport = m_nStartNummer;
    double    fUmst;
    long      dfNetto, dfBetrag, dfUmst;
    CString   szText;
    int       umSt, nummer;

    for (index = 0; index < CRechnung::OnlineShop; index++)
    {
        DoSort(arrRechnungen[index], OrderRechnung);
        countPl = arrRechnungen[index].GetCount();
        for (j = 0; j < countPl; j++)
        {
            if (m_mapRechnung.Lookup(arrRechnungen[index][j].szKey, rech))
            {
                pRechnung = &rech;

                if (rech.m_countGutschriftAmazon > 0)
                {
                    ctPreisRechnung = abs(NINT(100.0 * rech.m_dfBetragRechnung));
                    ctPreisGutWawi = abs(NINT(100.0 * rech.m_dfBetragGutschriftWawi));
                    ctPreisGutAmazon = abs(NINT(100.0 * rech.m_dfBetragGutschriftAmazon));
                    ctPreisGutAmazon = min(ctPreisGutAmazon, ctPreisRechnung);

                    if (ctPreisGutAmazon > ctPreisGutWawi)
                    {
                        ctPreisDiff = ctPreisGutAmazon - ctPreisGutWawi;
                        if (ctPreisDiff >= DELTA_CT_MIN && ctPreisDiff > 100)
                        {

                            if (index >= 0 && index < ECTK_COUNT_BUCHUNG)
                            {
                                ctAmazonGutschrift = ctPreisDiff;

                                faktorVK = NINT(100000.0 * rech.m_faktor);
                                faktorNK = faktorVK % 100000;
                                faktorVK /= 100000;
                                 
                                szUmst   = rech.m_szUmst;
                                fUmst    = atof(szUmst);
                                dfBetrag = ctAmazonGutschrift;

                                szWaehrung = rech.m_waehrungRechnung;
                                if (szWaehrung.GetLength() <= 0)
                                    szWaehrung = index == 1 ? "GBP" : "EUR";

                                // if (rech.m_ISO.GetLength() > 0 && rech.m_ISO != "DE")
                                if (CheckOSS(&rech))
                                {
                                    
                                    rech.m_dfAmazonKorrekturBetrag = dfBetrag / 100.0;
                                    m_ossData.AddLine(&rech, true);

                                    umSt = (int)((100.0 * fUmst) + 0.05);
                                    szText.Format("%s (%s %s %%)", rech.m_szRechnungEmpf, rech.m_ISO, szUmst);
                                    if (0 != umSt)
                                    {
                                        dfNetto = (long)(( dfBetrag / (1.0 + fUmst/100.0) + 0.5));
                                        dfUmst  = dfBetrag - dfNetto;

                                    }
                                    else
                                    {
                                        dfNetto = (long)(dfBetrag + 0.5);
                                        dfUmst  = 0;

                                    }
                                    
                                    nummer = ++countExport;
                                    
                                    
                                    // szKonto = KONTO_OSS_RUECK_NETTO;
                                    
                                    switch (index)
                                    {
                                      case CRechnung::Amazon_DE:
                                        indexOSS = COSSIndex::pl_amazonDE;
                                      break;
                                      case CRechnung::Amazon_Fr:
                                        indexOSS = COSSIndex::pl_amazonFR;
                                      break;
                                      case CRechnung::OnlineShop:
                                      case CRechnung::JTL_WaWi:
                                      case CRechnung::Haendler:
                                        indexOSS = COSSIndex::pl_webshop;
                                      break;
                                      case CRechnung::Amazon_GB:
                                      case CRechnung::unknown:
                                        indexOSS = COSSIndex::pl_amazonRest;
                                      break;
                                    }

                                    szKonto = ossIndex.GetTitle((COSSIndex::tagPlatform)indexOSS, COSSIndex::art_Ruecklaeufer, COSSIndex::betrag_Netto);
                                    szLine.Format(szFormat, rech.m_szDatum, -dfNetto / 100, DEZIMALTRENNER, dfNetto % 100, szText, pszGutschriftPre[index], nummer, szKonto, "0.00",szWaehrung,faktorVK, DEZIMALTRENNER, faktorNK,rech.m_szRechnungsNr, rech.m_extBestellnummer);
                                    ossRueckNetto[indexOSS] += szLine;
                                    countRueckNetto[indexOSS]   += 1;

                                    if (0 != umSt)
                                    {
                                        // szKonto = KONTO_OSS_RUECK_UMST;
                                        szKonto = ossIndex.GetTitle((COSSIndex::tagPlatform)indexOSS, COSSIndex::art_Ruecklaeufer, COSSIndex::betrag_Netto);
                                        szLine.Format(szFormat, rech.m_szDatum, -dfUmst / 100, DEZIMALTRENNER_EXCEL, dfUmst % 100, szText, pszGutschriftPre[index], nummer, szKonto, "0,00", szWaehrung, faktorVK, DEZIMALTRENNER_EXCEL, faktorNK, rech.m_szRechnungsNr, rech.m_extBestellnummer);
                                        ossRueckUmst[indexOSS]   += szLine;
                                        countRueckUmst[indexOSS] += 1;
                                    }

                                }
                                else
                                { 
                                    countExportLines[index] += 1;
                                    szLine.Format(szFormat, rech.m_szDatum, -ctAmazonGutschrift / 100, DEZIMALTRENNER, ctAmazonGutschrift % 100, rech.m_szRechnungEmpf,
                                        pszGutschriftPre[index], ++countExport, pszBuchungsArt[index], szUmst,
                                        szWaehrung,
                                        faktorVK, DEZIMALTRENNER, faktorNK,
                                        rech.m_szRechnungsNr, rech.m_extBestellnummer);
                                    szExportLines[index] += szLine;
                                }
                            }

                        }
                    }

                }
            }
        }
    }

    /*


    CKeyArray arrGutschriften;
    for (CMapGutschrift::CPair* pCurVal = m_mapGutschrift.PGetFirstAssoc(); pCurVal != NULL; pCurVal = m_mapGutschrift.PGetNextAssoc(pCurVal))
    {
        pGutschrift = &pCurVal->value;
        arrGutschriften.Add({ pCurVal->key, pGutschrift->m_szDatum});
    }

    DoSort(arrGutschriften, OrderDatum);

    totalCount = arrGutschriften.GetCount();

    CGutschrift gut;
    for (i = 0; i < totalCount; i++)
    {
        if (m_mapGutschrift.Lookup(arrGutschriften[i].szKey, gut))
        {
            pGutschrift = &gut;
            szRechnungsnummer = arrGutschriften[i].szKey;

            index = pGutschrift->m_typ == CRechnung::Amazon_DE ?  0 : (\
                    pGutschrift->m_typ == CRechnung::Amazon_GB ?  1 :  (\
                    pGutschrift->m_typ == CRechnung::Amazon_Fr ?  2 :   (\
                    pGutschrift->m_typ == CRechnung::OnlineShop ? 3 : -1)));

            if (index >= 0 && index < ECTK_COUNT_BUCHUNG)
            {
                dfPreis            = pGutschrift->m_dfGutschriftSumme;
                ctPreis            = abs(NINT(100.0 * dfPreis));
                ctPreisCmp         = abs(NINT(100.0 * pGutschrift->m_dfBetragGutschriftOrg));
                ctPreisRechnung    = abs(NINT(100.0 * pGutschrift->m_dfRechnungsBetrag));
                ctAmazonGutschrift = min(ctPreis, ctPreisRechnung);

                faktorVK           = NINT(100000.0 * pGutschrift->m_faktorRechnung);
                faktorNK           = faktorVK % 100000;
                faktorVK          /= 100000;

                szUmst             = pGutschrift->m_szUmst;
                szWaehrung         = pGutschrift->m_waehrungGutschrift;
                if (szWaehrung.GetLength() <= 0)
                    szWaehrung = index == 1 ? "GBP" : "EUR";

                if (pGutschrift->m_fErstattung)
                { 
                    ctGutschrift = ctAmazonGutschrift;
                    if (ctGutschrift > 0)
                    {
                        szLine.Format(szFormat, pGutschrift->m_szDatum, ctGutschrift / 100, DEZIMALTRENNER, ctGutschrift % 100, pGutschrift->m_empfaenger,
                            pszErstattungPre[index], ++countErstattungLines[index], pszBuchungsArtErstattung[index], szUmst,
                            szWaehrung,
                            faktorVK, DEZIMALTRENNER, faktorNK,
                            pGutschrift->m_rechnungsNummer, pGutschrift->m_extBestellNummer);
                        szExportLinesErstattung[index] += szLine;
                    }

                }
                if (abs(ctAmazonGutschrift - ctPreisCmp) >= DELTA_CT_MIN && ctPreis > 100)  // nur, wenn es auch eine Amazon-Gutschrift gibt ...
                {
                    ctGutschrift = ctAmazonGutschrift - ctPreisCmp;

                    if (ctGutschrift > 0)
                    {

                        {
                            szLine.Format(szFormat, pGutschrift->m_szDatum, -ctGutschrift / 100, DEZIMALTRENNER, ctGutschrift % 100, pGutschrift->m_empfaenger,
                                pszGutschriftPre[index], ++countExportLines[index], pszBuchungsArt[index], szUmst,
                                szWaehrung,
                                faktorVK, DEZIMALTRENNER, faktorNK,
                                pGutschrift->m_rechnungsNummer, pGutschrift->m_extBestellNummer);
                            szExportLines[index] += szLine;
                        }
                    }
                }
            }
        }
    }

*/    
    
    szFileTitle = szFlTitle;
    szFileTitle.Replace(".csv", "");

    for (index = 0; index < ECTK_COUNT_BUCHUNG; index++)
    {
        if (countExportLines[index] > 0)
        {
            szFileName.Format("%s\\%s%s.csv", lpszPath, szFileTitle, pszFileExt[index]);
            length = szExportLines[index].GetLength();
            
            if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
            {
                fl.Write((LPCSTR)szExportLines[index], length);
                fl.Close();
            }

            CString szMessage;
            szMessage.Format("Es wurde eine Amazon-Korrekturdatei\n%s\nmit %d Datensätzen erzeugt.\n", szFileName, countExportLines[index]);
            MessageBox(szMessage, "Info", MB_OK);
        }

        if (countErstattungLines[index] > 0)
        {
            szFileName.Format("%s\\%s%s.csv", lpszPath, szFileTitle, pszFileExtErstattung[index]);
            length = szExportLinesErstattung[index].GetLength();

            if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
            {
                fl.Write((LPCSTR)szExportLinesErstattung[index], length);
                fl.Close();
            }

            CString szMessage;
            szMessage.Format("Es wurde eine Amazon-Korrekturdatei\n%s\nmit %d Datensätzen erzeugt.\n", szFileName, countErstattungLines[index]);
            MessageBox(szMessage, "Info", MB_OK);
        }


    }

    // Nun noch die OSS-Rückbuchungen ..
    CString szFilePath;
    for (i = 0; i < COSSIndex::pl_count; i++)
    {
      if (countRueckNetto[i])
      {
        szFileName = ossIndex.GetFileTitle((COSSIndex::tagPlatform)i, COSSIndex::art_Ruecklaeufer, COSSIndex::betrag_Netto);
        szFilePath.Format("%s\\%s", lpszPath, szFileName);
        length = ossRueckNetto[i].GetLength();
        if (fl.Open(szFilePath, CFile::modeCreate | CFile::modeWrite))
        {
          fl.Write((LPCSTR)ossRueckNetto[i], length);
          fl.Close();
        }
      }

      if (countRueckUmst[i])
      {
        szFileName = ossIndex.GetFileTitle((COSSIndex::tagPlatform)i, COSSIndex::art_Ruecklaeufer, COSSIndex::betrag_Umst);
        szFilePath.Format("%s\\%s", lpszPath, szFileName);
        length = ossRueckUmst[i].GetLength();
        if (fl.Open(szFilePath, CFile::modeCreate | CFile::modeWrite))
        {
          fl.Write((LPCSTR)ossRueckUmst[i], length);
          fl.Close();
        }
      }
    }

    /*
    szFileName.Format("%s\\%s_OSS_Rückläufer_Netto.csv", lpszPath, szFileTitle);
    length = ossRueckNetto.GetLength();
    if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
    {
        fl.Write((LPCSTR)ossRueckNetto, length);
        fl.Close();
    }

    szFileName.Format("%s\\%s_OSS_Rückläufer_Umst.csv", lpszPath, szFileTitle);
    length = ossRueckUmSt.GetLength();
    if (fl.Open(szFileName, CFile::modeCreate | CFile::modeWrite))
    {
        fl.Write((LPCSTR)ossRueckUmSt, length);
        fl.Close();
    }
    */


}

void CJTLSteuerDlg::OnBnClickedCancel()
{
  // TODO: Fügen Sie hier Ihren Handlercode für Benachrichtigungen des Steuerelements ein.
  CDialogEx::OnCancel();
}

// 01234567
// 01-Jun-2019 UTC
CString CJTLSteuerDlg::GetDate(LPCSTR lpszData)
{
    CString szDate = lpszData;
    int len = strlen(lpszData);
    if (len > 11 && lpszData[2] == '-' && lpszData[6] == '-')
    {
        int mon, t;

        char tag[3], monat[4], jahr[5];
        memcpy(tag, lpszData, 2);
        tag[2] = '\0';
        t = atoi(tag);
        memcpy(monat, lpszData + 3, 3);
        monat[3] = '\0';
        memcpy(jahr, lpszData + 7, 4);
        jahr [4] = '\0';

        if (m_monthMap.Lookup(monat, mon))
            szDate.Format("%02d.%02d.%s", t, mon, jahr);
    }
    return szDate;

}

// --------------------------------------------------------------

void CJTLSteuerDlg::DoSort(CKeyArray& pArray, COMPARE comp)
{
    CKeys   save;
    bool    fChange;
    int     i, nCount;

    nCount = pArray.GetCount();
    do
    {
        fChange = false;
        for (i = 0; i < nCount - 1; i++)
        {
            /*
            nFirst = atoi(pArray[i].szValue.Mid(1, 4));
            nNext  = atoi(pArray[i+1].szValue.Mid(1, 4));
            if (nNext < nFirst)
            */
            if (comp(pArray[i].szValue, pArray[i + 1].szValue))
            {
                save = pArray[i];
                pArray[i] = pArray[i + 1];
                pArray[i + 1] = save;
                fChange = true;
            }
        }
    } while (fChange);
}

bool OrderString(CString& szFirst, CString& szNext)
{
    return strcmp(szFirst, szNext) > 0;
}


bool OrderRechnung(CString& szFirst, CString& szNext)
{
    int nFirst = atoi(szFirst.Mid(1, 4));
    int nNext = atoi(szNext.Mid(1, 4));
    return (nNext < nFirst);
}

bool OrderDatum(CString& szFirst, CString& szNext)
{
    int nFirst = GetSimpleDay(szFirst);
    int nNext  = GetSimpleDay(szNext);
    return (nNext < nFirst);
}

int GetSimpleDay(CString& szDatum)
{
    return atoi(szDatum.Mid(0, 2)) + 31 * (atoi(szDatum.Mid(3, 2)) - 1) + 400 * (atoi(szDatum.Mid(6, 4)) - 2000);
}

int GetSimpleDay(COleDateTime& datum)
{
    return datum.GetDay() + 31 * (datum.GetMonth() - 1) + 400 * (datum.GetYear() - 2000);
}

void CJTLSteuerDlg::OnDtnDatetimechangeDatumVon(NMHDR* pNMHDR, LRESULT* pResult)
{
    LPNMDATETIMECHANGE pDTChange = reinterpret_cast<LPNMDATETIMECHANGE>(pNMHDR);
    UpdateData(TRUE);
    *pResult = 0;
}


void CJTLSteuerDlg::OnDtnDatetimechangeDatumBis(NMHDR* pNMHDR, LRESULT* pResult)
{
    LPNMDATETIMECHANGE pDTChange = reinterpret_cast<LPNMDATETIMECHANGE>(pNMHDR);
    UpdateData(TRUE);
    *pResult = 0;
}



void CJTLSteuerDlg::OnEnChangeEditStartnr()
{
    UpdateData(TRUE);
}

void CJTLSteuerDlg::ErstelleTaxStatistik(LPCSTR lpszPath, LPCSTR szFileTitle)
{
    CMapVerkaufInfo* pInfo;
    CString          szTitle, szFileName;
    LPCSTR           arrEndungen[] = { "_Auswertung_DE.csv", "_Auswertung_GB.csv", "_Auswertung_FR.csv" };

    szTitle = szFileTitle;
    szTitle.Replace(".csv", "_statistik_");

    for (int i = 0; i < 3; i++)
    {
        szFileName.Format("%s\\%s%s", lpszPath, szTitle, arrEndungen[i]);
        pInfo = i == 0 ? &m_mapVerkaufDE : (i == 1 ? &m_mapVerkaufGB : &m_mapVerkaufFR);
        StoreAmazonAuswertung(pInfo, szFileName);
    }
}




// sj
// kids
// xx
// platform

/*
SandalenArt;Bestellungen;Retouren;Gesamt;Betrag;Betrag Euro;

Standard
San Juan  "sj-"
Kids "kids"
With Sole "xx"
Platform "platform"

*/

#define SA_INDEX_STANDARD   0
#define SA_INDEX_SAN_JUAN   1
#define SA_INDEX_KIDS       2
#define SA_INDEX_SOLE       3
#define SA_INDEX_PLATFORM   4 
#define SA_INDEX_COUNT      5


#define SA_TITLE "\"SandalenArt\";\"Anazhl Bestellungen\";\"Anazhl Retouren\";\"Anazhl Gesamt\";\"Betrag\";\"Betrag Euro;\""
#define SA_FORMAT "\r\n\"%s\";\"%d\";\"%d\";\"%d\";\"%3.02f\";\"%3.02f\""


void CJTLSteuerDlg::StoreAmazonAuswertung(CMapVerkaufInfo *pMapVerkauf, LPCSTR lpszFileName)
{
    CVerkaufInfo verkaufInfo;
    CVerkaufInfo arrVerkauf[SA_INDEX_COUNT];
    LPCSTR       arrSandalenArt[SA_INDEX_COUNT] = { "Standard", "San Juan", "Kids", "With Sole", "Platform" };
    CFile        fl;
    CString      szLine;
    int          index, len;


    for (index = 0; index < SA_INDEX_COUNT; index++)
        arrVerkauf[index].sandaleArt = arrSandalenArt[index];

    for (CMapVerkaufInfo::CPair* pCurVal = pMapVerkauf->PGetFirstAssoc(); pCurVal != NULL; pCurVal = pMapVerkauf->PGetNextAssoc(pCurVal))
    {
        verkaufInfo = pCurVal->value;
        index       = verkaufInfo.sandaleArt.Find("sj3")      >= 0 ? SA_INDEX_SAN_JUAN :                 (\
                      verkaufInfo.sandaleArt.Find("sj4")      >= 0 ? SA_INDEX_SAN_JUAN :                  (\
                      verkaufInfo.sandaleArt.Find("kids")     >= 0 ? SA_INDEX_KIDS     :                   (\
                      verkaufInfo.sandaleArt.Find("xx")       >= 0 ? SA_INDEX_SOLE     :                    (\
                      verkaufInfo.sandaleArt.Find("platform") >= 0 ? SA_INDEX_PLATFORM : SA_INDEX_STANDARD ))));

        arrVerkauf[index].anzahlBestellung     += verkaufInfo.anzahlBestellung;
        arrVerkauf[index].anzahlGesamt         += verkaufInfo.anzahlGesamt;
        arrVerkauf[index].anzahlRetoure        += verkaufInfo.anzahlRetoure;
        arrVerkauf[index].summmeEinnahmeOrg    += verkaufInfo.summmeEinnahmeOrg;
        arrVerkauf[index].summmeEinnahmeFaktor += verkaufInfo.summmeEinnahmeFaktor;

    }



    if (fl.Open(lpszFileName, CFile::modeCreate | CFile::modeWrite))
    {
        szLine = SA_TITLE;
        len    = szLine.GetLength();
        fl.Write((LPCSTR)szLine, len);

        for (index = 0; index < SA_INDEX_COUNT; index++)
        {
            szLine.Format(SA_FORMAT, arrVerkauf[index].sandaleArt, arrVerkauf[index].anzahlBestellung, 
                                     arrVerkauf[index].anzahlRetoure, arrVerkauf[index].anzahlGesamt, 
                                     arrVerkauf[index].summmeEinnahmeOrg, arrVerkauf[index].summmeEinnahmeFaktor);
            szLine.Replace(".", ",");
            len = szLine.GetLength();
            fl.Write((LPCSTR)szLine, len);
        }
        fl.Close();
    }

}

void CJTLSteuerDlg::SetButtons()
{

    GetDlgItem(IDC_RECHNUNGEN)->EnableWindow(TRUE);
    GetDlgItem(IDC_RECHNUNGEN_HAENDLER)->EnableWindow(m_fReadRechnungen);
    GetDlgItem(IDC_RECHNUNGEN_POS)->EnableWindow(m_fReadRechnungen && m_fReadRechnungenHaendler);

    GetDlgItem(IDC_GUTSCHRIFTEN)->EnableWindow(m_fReadPOS && m_fReadRechnungen && m_fReadRechnungenHaendler);
    GetDlgItem(IDC_AMAZON_TAX)->EnableWindow(m_fReadPOS && m_fReadRechnungen && m_fReadRechnungenHaendler && m_fReadGutschriften);

    GetDlgItem(IDC_AMAZON_TAX)->EnableWindow(m_fReadPOS && m_fReadRechnungen && m_fReadRechnungenHaendler && m_fReadGutschriften);
    GetDlgItem(IDC_START_CALC)->EnableWindow(m_fReadPOS && m_fReadRechnungen && m_fReadRechnungenHaendler && m_fReadGutschriften && m_fReadAmazon);
}


COSSData::COSSData()
{
    Reset();
}

void COSSData::Reset(void)
{
    m_szTotalLines = OSS_DATA_HEADER;
    m_totalLines = 0;
}

void COSSData::AddLine(CRechnung* pRechnung, bool fAmazon)
{
    CString szBrutto, szNetto, szUmst, szLine, szKonto;
    double  dfNetto, dfBrutto, dfUmst, fUst;

    fUst    = atof(pRechnung->m_szUmst);
    szKonto = (LPCSTR)(fAmazon ? "Amazon-Korrektur" : "Rechnung");
    dfBrutto = fAmazon ? -pRechnung->m_dfAmazonKorrekturBetrag : pRechnung->m_dfBetragRechnung;
    dfNetto  = dfBrutto / (1.0 + fUst / 100.0);
    dfUmst   = dfBrutto - dfNetto;

    szBrutto.Format("%02.2f", dfBrutto);
    szNetto.Format("%02.2f", dfNetto);
    szUmst.Format("%02.2f", dfUmst);

    szBrutto.Replace(".", ",");
    szNetto.Replace(".", ",");
    szUmst.Replace(".", ",");
    
 
    szLine.Format(OSS_DATA_FORMAT, pRechnung->m_szDatum, pRechnung->m_szRechnungEmpf, pRechnung->m_szRechnungsNr, pRechnung->m_extBestellnummer, szBrutto, szNetto, szUmst, pRechnung->m_ISO, pRechnung->m_waehrungRechnung, pRechnung->m_szUmst, szKonto, pRechnung->m_bestellNummer);
    m_szTotalLines += szLine;

    m_totalLines += 1;
}

void COSSData::AddLine(CGutschrift* pGutschrift)
{
    CString szBrutto, szNetto, szUmst, szLine, szKonto;
    double  dfNetto, dfUmst, fUst;

    fUst = atof(pGutschrift->m_szUmst);
    dfNetto = pGutschrift->m_dfBetragGutschrift / (1.0 + fUst / 100.0);
    dfUmst = pGutschrift->m_dfBetragGutschrift - dfNetto;

    szBrutto.Format("%02.2f", pGutschrift->m_dfBetragGutschrift);
    szNetto.Format("%02.2f", dfNetto);
    szUmst.Format("%02.2f", dfUmst);

    szBrutto.Replace(".", ",");
    szNetto.Replace(".", ",");
    szUmst.Replace(".", ",");

    szKonto = "Gutschrift";

    szLine.Format(OSS_DATA_FORMAT, pGutschrift->m_szDatum, pGutschrift->m_empfaenger, pGutschrift->m_gutschriftNummer, pGutschrift->m_extBestellNummer, szBrutto, szNetto, szUmst, pGutschrift->m_ISO, pGutschrift->m_waehrungGutschrift, pGutschrift->m_szUmst, szKonto, pGutschrift->m_wawiBestellNummer);
    m_szTotalLines += szLine;

    m_totalLines += 1;
}


bool COSSData::WriteData(CDialog *pParent, LPCSTR lpszFilenname)
{
    CFile fl;
    bool  fReturn = false;
    
    CString szMessage;
    szMessage.Format("Oss-Daten für Excel\nAnzahl der Zeilen: %d\nLänge: %d Bytes", m_totalLines, m_szTotalLines.GetLength());
    pParent->MessageBox(szMessage, "OSS-Daten", MB_OK);
    
    if (fl.Open(lpszFilenname, CFile::modeCreate | CFile::modeWrite))
    {
        int len = m_szTotalLines.GetLength();
        TRY
        {
            fl.Write((LPCSTR)m_szTotalLines, len);
            fReturn = true;
        }
        CATCH(CException, pE)
        {
            ;
        }
        END_CATCH
        fl.Close();
    }
    
    return fReturn;
}

// --------------------------------------------

COSSIndex::COSSIndex()
{

  CString platform[pl_count], art[art_count], frm[2], frmFileName[2], strBezeichnung;
  int     indexPlatform, indexArt, indexBetrag;

  platform[pl_amazonDE]     = "Amazon DE";
  platform[pl_amazonFR]     = "Amazon FR";
  platform[pl_amazonRest]   = "Amazon Rest";
  platform[pl_webshop]      = "Webshop";
  platform[pl_unknown]      = "Unknown";

  art[art_Rechnung]         = "Rechnungen";
  art[art_Gutschrift]       = "Gutschriften";
  art[art_Ruecklaeufer]     = "Rückäufer";

  frm[betrag_Netto]         = "%s %s EU ohne EU VAT Nummer";
  frm[betrag_Umst]          = "UmSt %s %s EU ohne EU VAT Nummer";

  frmFileName[betrag_Netto] = "OSS_%s_%s_Netto.csv";
  frmFileName[betrag_Umst]  = "OSS_%s_%s_Umst.csv";

  for (indexPlatform = 0; indexPlatform < pl_count; indexPlatform++)
  {
    for (indexArt = 0; indexArt < art_count; indexArt++)
    {
      for (indexBetrag = 0; indexBetrag < betrag_count; indexBetrag++)
      {
        // nummer = indexPlatform * pl_faktor + indexArt * art_faktor + indexBetrag * betrag_faktor;
        strBezeichnung.Format(frm[indexBetrag], art[indexArt], platform[indexPlatform]);
        m_mapTitle[GetInternNummer(indexPlatform, indexArt, indexBetrag)] = strBezeichnung;

        strBezeichnung.Format(frmFileName[indexBetrag], art[indexArt], platform[indexPlatform]);
        m_mapFileTitle[GetInternNummer(indexPlatform, indexArt, indexBetrag)] = strBezeichnung;
      }
    }
  }

}


int COSSIndex::GetInternNummer(int indexPlatform, int indexArt, int indexBetrag)
{
  if (indexBetrag >= 0)
    return indexPlatform * pl_faktor + indexArt * art_faktor + indexBetrag * betrag_faktor;
  else
    return indexPlatform * pl_faktor + indexArt * art_faktor;
}


CString COSSIndex::GetTitle(tagPlatform platform, tagArt art, tagBetrag betrag)
{
  CString szTitle = "";
  int nummer = GetInternNummer(platform, art, betrag);
  m_mapTitle.Lookup(nummer, szTitle);
  return szTitle;
}

CString COSSIndex::GetFileTitle(tagPlatform platform, tagArt art, tagBetrag betrag)
{
  CString szTitle = "";
  int nummer = GetInternNummer(platform, art, betrag);
  m_mapFileTitle.Lookup(nummer, szTitle);
  return szTitle;
}


