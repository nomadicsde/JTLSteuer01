
// JTLSteuerDlg.h: Headerdatei
//

#pragma once


#define RE_HEADER_FLD_NAME                \
{                                         \
  "Buchungsart",                          \
  "Erstelldatum Rechnung",                \
  "Gesamtbetrag Brutto (alle Ust.)",      \
  "RA Vorname RA Nachname",               \
  "Rechnungsnummer",                      \
  "Konto",                                \
  "USt.",                                 \
  "Währung",                              \
  "Währungsfaktor",                       \
  "Bestellnummer",                        \
  "Externe Bestellnummer",                \
  "Plattform",                            \
  "Debitorennummer",                      \
  "Kunden Gruppe",                        \
  "Zahlungsbetrag",                       \
  "Zahlungshinweis",                      \
  "Zahlungsdatum",                        \
  "ISO"                                   \
}


#define GU_HEADER_FLD_NAME                \
{                                         \
  "Buchungsart",                          \
  "Erstelldatum",                         \
  "Gesamtbetrag Brutto (alle Ust.)",      \
  "RA Vorname RA Nachname",               \
  "Gutschriftsnummer",                    \
  "Konto",                                \
  "USt.",                                 \
  "Währung",                              \
  "Währungsfaktor",                       \
  "Bezug Rechnungsnummer",                \
  "Externe Bestellnummer",                \
  "Debitorennummer",                      \
  "Buchungskonto",                        \
  "ISO"                                   \
};

#define AS_HEADER_FLD_NAME               { \
  "Marketplace ID",                        \
  "Merchant ID",                           \
  "Order Date",                            \
  "Transaction Type",                      \
  "Is Invoice Corrected",                  \
  "Order ID",                              \
  "Shipment Date",                         \
  "Shipment ID",                           \
  "Transaction ID",                        \
  "ASIN",                                  \
  "SKU",                                   \
  "Quantity",                              \
  "Tax Calculation Date",                  \
  "Tax Rate",                              \
  "Product Tax Code",                      \
  "Currency",                              \
  "Tax Type",                              \
  "Tax Calculation Reason Code",           \
  "Tax Reporting Scheme",                  \
  "Tax Collection Responsibility",         \
  "Tax Address Role",                      \
  "Jurisdiction Level",                    \
  "Jurisdiction Name",                     \
  "OUR_PRICE Tax Inclusive Selling Price", \
  "OUR_PRICE Tax Amount",                  \
  "OUR_PRICE Tax Exclusive Selling Price", \
  "OUR_PRICE Tax Inclusive Promo Amount",  \
  "OUR_PRICE Tax Amount Promo",            \
  "OUR_PRICE Tax Exclusive Promo Amount",  \
  "SHIPPING Tax Inclusive Selling Price",  \
  "SHIPPING Tax Amount",                   \
  "SHIPPING Tax Exclusive Selling Price",  \
  "SHIPPING Tax Inclusive Promo Amount",   \
  "SHIPPING Tax Amount Promo",             \
  "SHIPPING Tax Exclusive Promo Amount",   \
  "GIFTWRAP Tax Inclusive Selling Price",  \
  "GIFTWRAP Tax Amount",                   \
  "GIFTWRAP Tax Exclusive Selling Price",  \
  "GIFTWRAP Tax Inclusive Promo Amount",   \
  "GIFTWRAP Tax Amount Promo",             \
  "GIFTWRAP Tax Exclusive Promo Amount",   \
  "Seller Tax Registration",               \
  "Seller Tax Registration Jurisdiction",  \
  "Buyer Tax Registration",                \
  "Buyer Tax Registration Jurisdiction",   \
  "Buyer Tax Registration Type",           \
  "Buyer E Invoice Account Id",            \
  "Invoice Level Currency Code",           \
  "Invoice Level Exchange Rate",           \
  "Invoice Level Exchange Rate Date",      \
  "Converted Tax Amount",                  \
  "VAT Invoice Number",                    \
  "Invoice Url",                           \
  "Export Outside EU",                     \
  "Ship From City",                        \
  "Ship From State",                       \
  "Ship From Country",                     \
  "Ship From Postal Code",                 \
  "Ship From Tax Location Code",           \
  "Ship To City",                          \
  "Ship To State",                         \
  "Ship To Country",                       \
  "Ship To Postal Code",                   \
  "Ship To Location Code",                 \
  "Return Fc Country",                     \
  "Is Amazon Invoiced",                    \
  "Original VAT Invoice Number",           \
  "Invoice Correction Details",            \
  "SDI Invoice Delivery Status",           \
  "SDI Invoice Error Code",                \
  "SDI Invoice Error Description",         \
  "SDI Invoice Status Last Updated Date",  \
  "EInvoice URL"                           \
}

/*
#define AS_HEADER_FLD_NAME                   { \
   "Marketplace ID",                           \
   "Merchant ID",                              \
   "Order Date",                               \
   "Transaction Type",                         \
   "Order ID",                                 \
   "Shipment Date",                            \
   "Shipment ID",                              \
   "Transaction ID",                           \
   "ASIN",                                     \
   "SKU",                                      \
   "Quantity",                                 \
   "Tax Calculation Date",                     \
   "Tax Rate",                                 \
   "Product Tax Code",                         \
   "Currency",                                 \
   "Tax Type",                                 \
   "Tax Calculation Reason Code",              \
   "Tax Address Role",                         \
   "Jurisdiction Level",                       \
   "Jurisdiction Name",                        \
   "OUR_PRICE Tax Inclusive Selling Price",    \
   "OUR_PRICE Tax Amount",                     \
   "OUR_PRICE Tax Exclusive Selling Price",    \
   "OUR_PRICE Tax Inclusive Promo Amount",     \
   "OUR_PRICE Tax Amount Promo",               \
   "OUR_PRICE Tax Exclusive Promo Amount",     \
   "SHIPPING Tax Inclusive Selling Price",     \
   "SHIPPING Tax Amount",                      \
   "SHIPPING Tax Exclusive Selling Price",     \
   "SHIPPING Tax Inclusive Promo Amount",      \
   "SHIPPING Tax Amount Promo",                \
   "SHIPPING Tax Exclusive Promo Amount",      \
   "GIFTWRAP Tax Inclusive Selling Price",     \
   "GIFTWRAP Tax Amount",                      \
   "GIFTWRAP Tax Exclusive Selling Price",     \
   "GIFTWRAP Tax Inclusive Promo Amount",      \
   "GIFTWRAP Tax Amount Promo",                \
   "GIFTWRAP Tax Exclusive Promo Amount",      \
   "Seller Tax Registration",                  \
   "Seller Tax Registration Jurisdiction",     \
   "Buyer Tax Registration",                   \
   "Buyer Tax Registration Jurisdiction",      \
   "Buyer Tax Registration Type",              \
   "Buyer E Invoice Account Id",               \
   "Invoice Level Currency Code",              \
   "Invoice Level Exchange Rate",              \
   "Invoice Level Exchange Rate Date",         \
   "Converted Tax Amount",                     \
   "VAT Invoice Number",                       \
   "Invoice Url",                              \
   "Export Outside EU",                        \
   "Ship From City",                           \
   "Ship From State",                          \
   "Ship From Country",                        \
   "Ship From Postal Code",                    \
   "Ship From Tax Location Code",              \
   "Ship To City",                             \
   "Ship To State",                            \
   "Ship To Country",                          \
   "Ship To Postal Code",                      \
   "Ship To Location Code",                   \
   "Return Fc Country",                       \
   "Is Amazon Invoiced"                       \
};
*/

#define AS_INDEX_MARKETPLACE          0    // FR, DE, GB                   
#define AS_INDEX_DATUM                2    // Rechnungsdatum
#define AS_INDEX_TYP                  3    // SHIPMENT, RETURN, REFUND
#define AS_INDEX_BESTELLNR_EXT        5
#define AS_INDEX_SKU                 10    // Typ
#define AS_INDEX_ANZAHL              11    // Anzahl - Betrag bezieht sich auf die Gesamtsumme
#define AS_INDEX_UMST                13
#define AS_INDEX_WAEHRUNG            15    // EUR, GBP
#define AS_INDEX_TAX_REP_SCHEME      18    // VCS_EU_OSS
#define AS_INDEX_PREIS_BRUTTO        23
#define AS_INDEX_PREIS_UMST          24
#define AS_INDEX_PREIS_NETTO         25
#define AS_INDEX_VERSAND_BRUTTO      29    // Versandkosten
#define AS_INDEX_VERSAND_UMST        30
#define AS_INDEX_VERSAND_NETTO       31
#define AS_INDEX_VERSAND_AKT_BRUTTO  32    // Aktion Versand
#define AS_INDEX_VERSAND_AKT_UMST    33
#define AS_INDEX_VERSAND_AKT_NETTO   34
#define AS_INDEX_GESCHENK_BRUTTO     35    // Als Geschenk
#define AS_INDEX_GESCHENK_UMST       36
#define AS_INDEX_GESCHENK_NETTO      37
#define AS_INDEX_UMST_ID_KAEUFER     43
#define AS_INDEX_UMST_BERECHNET_CUR  47    // Währung für berechnete Umsatzsteuer (solte EUR sein)
#define AS_INDEX_FAKTOR_WAEHRUNG     48    // Wenn hier ein Wert enthalten ist, ist dies der Faktor * 100 für die Umrechnung [Währung] -> [Umst Eur] = Faktor * [Betrag UmSt]
#define AS_INDEX_UMST_BERECHNET      50    // Berechneter Wert der Umsatzsteuer bei Ausland-Versand
#define AS_INDEX_EXPORT_NONEU        53    // true oder false
#define AS_INDEX_EMPF_STADT          59
#define AS_INDEX_EMPF_ISO            61
#define AS_INDEX_EMPF_PLZ            62



#define GU_INDEX_DATUM                    1
#define GU_INDEX_BETRAG                   2
#define GU_INDEX_GUTSCHRIFTNUMMER         4
#define GU_INDEX_UMST                     6
#define GU_INDEX_WAHRUNG                  7
#define GU_INDEX_WAHRUNG_FAKT             8
#define GU_INDEX_RECHNUNGSNUMMER          9
#define GU_INDEX_BESTELLNUMMER            10
#define GU_INDEX_ISO                      13

#define RE_INDEX_DATUM                    1
#define RE_INDEX_BETRAG                   2
#define RE_INDEX_VORNACHNAME              3
#define RE_INDEX_RECHNUNGSNUMMER          4
#define RE_INDEX_UMST                     6
#define RE_INDEX_BESTELLNUMMER            9
#define RE_INDEX_WAEHRUNG                 7
#define RE_INDEX_WAEHRUNG_FAKT            8
#define RE_INDEX_EXT_BESTELLNUMMER        10
#define RE_INDEX_PLATTFORM                11
#define RE_INDEX_DEBITORENNUMMER          12
#define RE_INDEX_KUNDENGRUPPE             13
#define RE_INDEX_ZAHLUNGSBETRAG           14
#define RE_INDEX_ZAHLUNGSHINWEIS          15
#define RE_INDEX_ZAHLUNGSDATUM            16
#define RE_INDEX_ISO                      17

#define TEXT_KUNDENGRUPPE                 "Endkunden"
#define TEXT_SKONTOBUCHUNG                "Skontobuchung"
#define FAKTOR_SKONTO                     0.02

#define PLATFORM_ARR     {"Amazon.de",    "Amazon.co.uk",  "Amazon.fr",   "Onlineshop",   "JTL-Wawi", "Händler", "UnKnown" }
#define PLATFORM_ARR_OUT {"Amazon_DE",    "Amazon_UK",     "Amazon_FR",   "Onlineshop",   "JTL_Wawi", "Händler", "UnKnown" }
#define COUNT_PLATFORM   7

#define ECTK_COUNT_BUCHUNG      4
//#define ECTK_GUTSCHRIFT_PRE    {"K_DE_",    "K_UK_",  "K_FR_",   "K_NO_"}
#define ECTK_GUTSCHRIFT_PRE    {"K",    "K",  "K",   "K"}
#define ECTK_BUCHUNGSART       {"Gutschriften Rückläufer Amazon DE", "Gutschriften Rückläufer Amazon UK", "Gutschriften Rückläufer Amazon FR", "Gutschriften Rückläufer Webshop Nomadics"}
#define ECTK_FILE_EXT          {"_Amazon_Korrektur_DE", "_Amazon_Korrektur_UK", "_Amazon_Korrektur_FR", "_Amazon_Korrektur_Shop"}

#define ECTK_ERSTATTUNG_PRE    {"E_DE_",    "E_UK_",  "E_FR_",   "E_NO_"}
#define ECTK_BUCHUNG_ERST      {"Erstattung Amazon DE", "Erstattung Amazon UK", "Erstattung Amazon FR", "Erstattung Webshop Nomadics"}
#define ECTK_FILE_ERST_EXT     {"_Amazon_Erstattung_DE", "_Amazon_Erstattung_UK", "_Amazon_Erstattung_FR", "_Amazon_Erstattung_Shop"}

#define ECT_GUTSCHRIFT_TITEL    "\"Buchungsart\";\"Erstelldatum\";\"Gesamtbetrag Brutto (alle Ust.)\";\"RA Vorname RA Nachname\";\"Gutschriftsnummer\";\"Konto\";\"USt.\";\"Währung\";\"Währungsfaktor\""
#define ECT_RECHNUNG_TITEL      "\"Buchungsart\";\"Erstelldatum Rechnung\";\"Gesamtbetrag Brutto (alle Ust.)\";\"RA Vorname RA Nachname\";\"Rechnungsnummer\";\"Konto\";\"USt.\";\"Währung\";\"Währungsfaktor\""
#define ECT_RECHNUNG_KONTO    { "Rechnung Amazon DE", "Rechnung Amazon UK", "Rechnung Amazon FR", "Rechnung Webshop Nomadics", "Rechnungen WaWi", "Rechnungen Händler", "Rechnungen Sonstige"}
#define ECT_GUTSCHRIFT_KONTO  { "Gutschriften Amazon DE", "Gutschriften Amazon UK", "Gutschriften Amazon FR", "Gutschriften Webshop Nomadics", "Gutschriften WaWi", "Gutschriften Händler", "Gutschriften Sonstige"}

#define ECT_RE_GU_FORM          "\r\n\"Einnahmen\";\"%s\";\"%3.2f\";\"%s\";\"%s\";\"%s\";\"%s\";\"%s\";\"%1.6f\""
#define ECT_RE_GU_FORM_EXCEL    "\r\n\"Einnahmen\";\"%s\";\"%s\";\"%s\";\"%s\";\"%s\";\"%s\";\"%s\";\"%s\""

// #define ECT_GUTSCHRIFT_FORM   "\r\n\"Einnahmen";"04.03.2020";"-666.40";"Mareen Burk";"G2662";"Gutschriften (19%)";"19.00";"EUR";"1,00000000000000";"R7504";;"0";"0";
#define ECT_HEADER_SUMMERY       "\"Buchungsart\";\"Anz. Rechn.\";\"Betrag Rechn\";\"Anz. Gutschr.\";\"Betrag Gutschr.\""

/*
#define KONTO_OSS_EINNAHMEN_NETTO    "Rechnungen EU ohne EU VAT Nummer"
#define KONTO_OSS_EINNAHMEN_UMST     "UmSt Rechnungen EU ohne EU VAT Nummer"

#define KONTO_OSS_GUTSCHRIFTEN_NETTO "Gutschriften EU ohne EU VAT Nummer"
#define KONTO_OSS_GUTSCHRIFTEN_UMST  "UmSt Gutschriften EU ohne EU VAT Nummer"

#define KONTO_OSS_RUECK_NETTO        "Rückläufer EU ohne EU VAT Nummer"
#define KONTO_OSS_RUECK_UMST         "UmSt Rückläufer EU ohne EU VAT Nummer"

// OSS - Amazon DE
#define KONTO_OSS_EINNAHMEN_NETTO    "Rechnungen Amazon DE EU ohne EU VAT Nummer"
#define KONTO_OSS_EINNAHMEN_UMST     "UmSt Rechnungen Amazon DE EU ohne EU VAT Nummer"

#define KONTO_OSS_GUTSCHRIFTEN_NETTO "Gutschriften Amazon DE EU ohne EU VAT Nummer"
#define KONTO_OSS_GUTSCHRIFTEN_UMST  "UmSt Gutschriften Amazon DE EU ohne EU VAT Nummer"

#define KONTO_OSS_RUECK_NETTO        "Rückläufer Amazon DE EU ohne EU VAT Nummer"
#define KONTO_OSS_RUECK_UMST         "UmSt Rückläufer Amazon DE EU ohne EU VAT Nummer"

// OSS - Amazon FR

#define KONTO_OSS_EINNAHMEN_NETTO    "Rechnungen Amazon FR EU ohne EU VAT Nummer"
#define KONTO_OSS_EINNAHMEN_UMST     "UmSt Rechnungen Amazon FR EU ohne EU VAT Nummer"

#define KONTO_OSS_GUTSCHRIFTEN_NETTO "Gutschriften Amazon FR EU ohne EU VAT Nummer"
#define KONTO_OSS_GUTSCHRIFTEN_UMST  "UmSt Gutschriften Amazon FR EU ohne EU VAT Nummer"

#define KONTO_OSS_RUECK_NETTO        "Rückläufer Amazon FR EU ohne EU VAT Nummer"
#define KONTO_OSS_RUECK_UMST         "UmSt Rückläufer Amazon FR EU ohne EU VAT Nummer"

// OSS Amazon Rest
#define KONTO_OSS_EINNAHMEN_NETTO    "Rechnungen Amazon Rest EU ohne EU VAT Nummer"
#define KONTO_OSS_EINNAHMEN_UMST     "UmSt Rechnungen Amazon Rest EU ohne EU VAT Nummer"

#define KONTO_OSS_GUTSCHRIFTEN_NETTO "Gutschriften Amazon Rest EU ohne EU VAT Nummer"
#define KONTO_OSS_GUTSCHRIFTEN_UMST  "UmSt Gutschriften Amazon Rest EU ohne EU VAT Nummer"

#define KONTO_OSS_RUECK_NETTO        "Rückläufer Amazon Rest EU ohne EU VAT Nummer"
#define KONTO_OSS_RUECK_UMST         "UmSt Rückläufer Amazon Rest EU ohne EU VAT Nummer"

// OSS Webshop
#define KONTO_OSS_EINNAHMEN_NETTO    "Rechnungen Webshop EU ohne EU VAT Nummer"
#define KONTO_OSS_EINNAHMEN_UMST     "UmSt Rechnungen Webshop EU ohne EU VAT Nummer"

#define KONTO_OSS_GUTSCHRIFTEN_NETTO "Gutschriften Webshop EU ohne EU VAT Nummer"
#define KONTO_OSS_GUTSCHRIFTEN_UMST  "UmSt Gutschriften Webshop EU ohne EU VAT Nummer"
*/

class COSSIndex
{
  public:
    enum tagPlatform
    {
      pl_amazonDE = 0,
      pl_amazonFR = 1,
      pl_amazonRest = 2,
      pl_webshop = 3,
      pl_unknown = 4,
      pl_count = 5,
      pl_faktor = 1
    };

    enum tagArt
    {
      art_Rechnung = 0,
      art_Gutschrift = 1,
      art_Ruecklaeufer = 2,
      art_count = 3,
      art_faktor = 10
    };

    enum tagBetrag
    {
      betrag_Netto = 0,
      betrag_Umst = 1,
      betrag_count = 2,
      betrag_faktor = 100
    };

  public:
    COSSIndex();

  public:
    int GetCount_Platform(void) { return pl_count; }
    int GetCount_Art(void) { return art_count; }
    int GetCount_Betrag(void) { return betrag_count; }

  public:
    CString GetTitle(tagPlatform platform, tagArt art, tagBetrag betrag);
    CString GetFileTitle(tagPlatform platform, tagArt art, tagBetrag betrag);

  protected:
    int  GetInternNummer(int indexPlatform, int indexArt, int indexBetrag);

  protected:
    CMap<int, int, CString, LPCSTR> m_mapTitle;
    CMap<int, int, CString, LPCSTR> m_mapFileTitle;

};





bool OrderRechnung(CString& szFirst, CString& szNext);
bool OrderDatum(CString& szFirst, CString& szNext);
bool OrderString(CString& szFirst, CString& szNext);
int  GetSimpleDay(COleDateTime& datum);
int  GetSimpleDay(CString& szDatum);


class CRechnung
{
  public:
    enum typRechnung     { Amazon_DE = 0, Amazon_GB = 1,   Amazon_Fr = 2, OnlineShop = 3, JTL_WaWi = 4, Haendler = 5, unknown = 6, countPlatform = COUNT_PLATFORM  };

  public:
    CRechnung()          { Reset();                                                                                                                  }
    void Reset(void) 
    {
        m_faktor = 1;              m_bestellNummer = ""; m_extBestellnummer = ""; m_dfBetragRechnung   = 0;     m_dfBetragRechnungKorr = 0;     m_waehrungRechnung     = "";    m_typ = unknown;                     m_gutschriftNummer            = "";    m_dfBetragGutschriftWawi = 0; 
        m_waehrungGutschrift = ""; m_platform = "";      m_szRechnungsNr = "";    m_fHasGutschriftWawi = false; m_fAmazonNeuRechnung   = false; m_fAmazonNeuGutschrift = false; m_fAmazonDiffBetragRechnung = false; m_fAmazonDiffBetragGutschrift = false; m_dfRechnungSumme    = 0;  m_dfGutschriftSumme = 0;
        m_dfBetragGutschriftOrg = 0; m_szRechnungEmpf = "", m_szDatum = ""; m_szUmst = "19.00";                 m_countGutschriftWawi  = 0; m_dfBetragGutschriftAmazon = 0;
        m_countGutschriftAmazon = 0; m_fHasGutschriftAmazon = false; m_szDatumGutschriftAmazon = ""; m_dfKontrollSumme = 0; m_ISO = ""; m_dfAmazonKorrekturBetrag = 0; m_fAmazonOSS = false; m_fAmazonTaxReport = false;
    }

  public:
    CString              m_bestellNummer;
    CString              m_szRechnungEmpf;
    double               m_dfBetragRechnung;
    double               m_dfBetragRechnungKorr;
    CString              m_waehrungRechnung;
    double               m_faktor;
    typRechnung          m_typ;
    
    bool                 m_fAmazonTaxReport;
    bool                 m_fAmazonOSS;
    
    // Wawi
    CString              m_gutschriftNummer;            // Sollte so nicht genutzt werden
    
    
    bool                 m_fHasGutschriftWawi;
    double               m_dfBetragGutschriftWawi;      // Summer aller Gutschriften
    double               m_dfBetragGutschriftOrg;       // entspricht m_dfBetragGutschriftWawi
    double               m_dfAmazonKorrekturBetrag;     // Hilfsvariable bei OSS
    int                  m_countGutschriftWawi;


    CString              m_waehrungGutschrift;
    CString              m_extBestellnummer;
    CString              m_platform;
    CString              m_szRechnungsNr;
    CString              m_szDatum;
    CString              m_szUmst;
    CString              m_ISO;


    // Amazon
    bool                 m_fHasGutschriftAmazon;
    double               m_dfBetragGutschriftAmazon;     // Summer aller Gutschriften
    int                  m_countGutschriftAmazon;
    CString              m_szDatumGutschriftAmazon;


    
    bool                 m_fAmazonNeuRechnung;           // Wenn die Rechnung zu einer Bestellnummer nicht gefunden werden kann
    bool                 m_fAmazonNeuGutschrift;
    bool                 m_fAmazonDiffBetragRechnung;
    bool                 m_fAmazonDiffBetragGutschrift;
    double               m_dfRechnungSumme;              // Amazon ..
    double               m_dfGutschriftSumme;

    double               m_dfKontrollSumme;
};

// Gutschriftnummer
// Rechnungsnummer, Bestellnummer, Ext-Bestellnummer, Datum, Platform, Faktor, Betrag, Währung, 

class CGutschrift
{

public:
    CGutschrift()     { Reset(); }
    void Reset(void) {m_wawiBestellNummer = "";      m_extBestellNummer   = "";    m_gutschriftNummer      = ""; m_szDatum            = "";     m_platform  = "";               m_typ               = CRechnung::unknown; 
                      m_dfRechnungsBetrag = 0;       m_dfBetragGutschrift =  0;    m_dfBetragGutschriftOrg =  0; m_waehrungGutschrift = "";     m_fAmazonNeuGutschrift = false; m_dfGutschriftSumme = 0; m_fAmazonDiffBetragGutschrift = false;
                      m_empfaenger        = "";      m_szDatum1           = "";    m_szDatum2              = ""; m_szDatum3           = "";     m_dfBetrag1 = 0;                m_dfBetrag2 = 0;         m_dfBetrag3 = 0; m_countItem = 0;   
                      m_szUmst            = "19.00"; m_faktor             =  1;    m_faktorRechnung        =  0; m_fErstattung        = false;  m_ISO = "";                     m_szRechnungsDatum = "";
                      m_fAmazonOSS        = false;   m_fAmazonTaxReport   = false;
                     }
    

public:
    CString                 m_rechnungsNummer;
    CString                 m_szRechnungsDatum;
    CString                 m_empfaenger;
    CString                 m_wawiBestellNummer;
    CString                 m_extBestellNummer;
    CString                 m_gutschriftNummer;
    CString                 m_szDatum;
    CString                 m_szUmst;
    CString                 m_ISO;
    double                  m_faktor;
    double                  m_faktorRechnung;

    CString                 m_platform;
    CRechnung::typRechnung  m_typ;

    
    // Amazon
    CString                 m_szDatum1,  m_szDatum2,  m_szDatum3;
    double                  m_dfBetrag1, m_dfBetrag2, m_dfBetrag3;
    int                     m_countItem;
    
    
    double                  m_dfRechnungsBetrag;
    
    double                  m_dfBetragGutschrift;
    double                  m_dfBetragGutschriftOrg;  // ohne Faktor
    CString                 m_waehrungGutschrift;

    bool                    m_fAmazonNeuGutschrift;
    bool                    m_fErstattung;
    double                  m_dfGutschriftSumme;

    bool                    m_fAmazonDiffBetragGutschrift;

    bool                    m_fAmazonOSS;
    bool                    m_fAmazonTaxReport;
};

#define OSS_DATA_HEADER  "\"Erstelldatum Rechnung\";\"Rechnungsempfänger\";\"Rechnungsnummer\";\"Externe Bestellnummer\";\"Betrag Brutto\";\"Betrag Netto\";\"Betrag UmSt.\";\"Land\";\"Währung\";\"Umsatsteuer-Satz\";\"Konto\";\"Bestellnummer\""
#define OSS_DATA_FORMAT  "\r\n\"%s\";\"%s\";\"%s\";\"%s\";\"%s\";\"%s\";\"%s\";\"%s\";\"%s\";\"%s\";\"%s\";\"%s\""

class COSSData
{
    public:
        COSSData();

    public:
        void Reset(void);

        void AddLine(CRechnung* pRechnung, bool fAmazon = false);
        void AddLine(CGutschrift* pGutschrift);

        bool WriteData(CDialog* pParent, LPCSTR lpszFilenname);

    protected:
        CString m_szTotalLines;
        int     m_totalLines;
};


typedef CMap<CString, LPCSTR, CRechnung, CRechnung> CMapRechnung;
typedef CMap < CString, LPCSTR, CGutschrift, CGutschrift> CMapGutschrift;
typedef CMap < CString, LPCSTR, int, int> CMapStringToInt;

typedef CArray <CGutschrift, CGutschrift> CArrGutschrift;
typedef CMap   <CString, LPCSTR, CArrGutschrift*, CArrGutschrift*> CMapGutschriftArray;

typedef bool (COMPARE)(CString& szFirst, CString& szNext);

class CKeys
{
public:
    CString szKey;
    CString szValue;
};
typedef CArray <CKeys, CKeys> CKeyArray;

class CVerkaufInfo
{
    public:
        CVerkaufInfo()                        { Reset(); }
        void Reset(LPCSTR lpszSandale = NULL) { sandaleArt = lpszSandale ? lpszSandale : ""; anzahlBestellung = 0; anzahlRetoure = 0; anzahlGesamt = 0; summmeEinnahmeOrg = 0; summmeEinnahmeFaktor = 0; }
          
    public:
        CString sandaleArt;
        int     anzahlBestellung;
        int     anzahlRetoure;
        int     anzahlGesamt;
        double  summmeEinnahmeOrg;
        double  summmeEinnahmeFaktor;
};

typedef CMap<CString, LPCSTR, CVerkaufInfo, CVerkaufInfo>  CMapVerkaufInfo;



// CJTLSteuerDlg-Dialogfeld
class CJTLSteuerDlg : public CDialogEx
{
// Konstruktion
public:
	CJTLSteuerDlg(CWnd* pParent = NULL);	// Standardkonstruktor
    ~CJTLSteuerDlg()     { ResetMapGutschriftenWawi(); ResetMapGutschriftenAmazon();  ReadSaveData(true);  }

// Dialogfelddaten
	enum { IDD = IDD_JTLSTEUER_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV-Unterstützung

  bool ReadAmazon(void);

  bool DoReadRechnung(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName, bool fKunden, bool fReset);
  bool DoReadGutschriften(LPCSTR lpszFilePath, LPCSTR lpsPath, LPCSTR lpszName);
  bool DoAmazonSteuer(LPCSTR lpszFilePath, LPCSTR lpszPath, LPCSTR szFileTitle);
  void StoreResult(LPCSTR lpsPath, LPCSTR lpszName);
  void ErstelleEasyCashKorrekturImport(LPCSTR lpszPath, LPCSTR szFileTitle);
  void ErstelleTaxStatistik(LPCSTR lpszPath, LPCSTR szFileTitle);
  void ErstelleEasyCashRechnungen(LPCSTR lpszPath, LPCSTR szFileTitle);
  void ErstelleEasyCashGutschriften(LPCSTR lpszPath, LPCSTR szFileTitle);

  void SetButtons(void);
  void StoreResults(void);
  bool AddGutschriftZuRechnung(CGutschrift* pGutschrift, double* pBetrag);
  void PruefeGutschriftenSumme(void);
  
  void ResetMapGutschriftenWawi(void);
  void ResetMapGutschriftenAmazon(void);
  void ReadSaveData(bool fSave = false);
  
  void StoreAmazonAuswertung(CMapVerkaufInfo* pMapVerkauf, LPCSTR lpszPath);

  bool CheckOSS(LPCSTR lpszISO, LPCSTR lpszUmsatz, LPCSTR lpszDatum);
  bool CheckOSS(CRechnung* pRechnung);                           // { return pRechnung->m_szUmst.GetLength() > 0 && pRechnung->m_ISO.GetLength() > 0 && CheckOSS(pRechnung->m_ISO, pRechnung->m_szUmst);          }
  bool CheckOSS(CGutschrift* pGutschrift);                       // { return pGutschrift->m_szUmst.GetLength() > 0 && pGutschrift->m_ISO.GetLength() > 0 && CheckOSS(pGutschrift->m_ISO, pGutschrift->m_szUmst);  }
  bool IsEU(LPCSTR lpszISO)                                         { int val = 0; return m_mapEU.Lookup(lpszISO, val) ? true : false;                                                                            }
  bool IsOSSDate(LPCSTR lpszDatum);

  CString GetDate(LPCSTR lpszData);
  void    DoSort(CKeyArray& pArray, COMPARE comp);


// Implementierung
protected:  
	HICON m_hIcon;

	// Generierte Funktionen für die Meldungstabellen
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
  afx_msg void OnBnClickedRechnungen();
  afx_msg void OnBnClickedGutschriften();
  afx_msg void OnBnAmazonTax();
  afx_msg void OnBnStartCalc();


protected:
  CMapGutschriftArray    m_mapArrRechnungGutschriftWawi;        // ordnet Rechnungsnummer eine Gutschrift zu 
  CMapGutschriftArray    m_mapArrRechnungGutschriftAmazon;  // ordnet Rechnungsnummer eine Gutschrift zu 
  CMapRechnung           m_mapRechnung;
  CMapRechnung           m_mapAll;
  CMapGutschrift         m_mapGutschrift;
  CStringArray           m_arrGuKorr;
  CMapStringToInt        m_monthMap;
  CString                m_arrSummery[CRechnung::countPlatform];

  COSSData               m_ossData;

public:
  afx_msg void OnBnClickedCancel();
  COleDateTime m_dateVon;
  COleDateTime m_dateBis;


  CMapVerkaufInfo m_mapVerkaufDE;
  CMapVerkaufInfo m_mapVerkaufGB;
  CMapVerkaufInfo m_mapVerkaufFR;

  CMapStringToInt m_mapEU;

  CString         m_iniFileName;

  bool m_fReadRechnungen;
  bool m_fReadRechnungenHaendler;
  bool m_fReadGutschriften;
  bool m_fReadAmazon;

  CString m_storePath;
  CString m_storeTitle;

  int     m_GutschriftNotFound;
  int     m_GutschriftFound;
  int     m_GutschriftExceedCount;
  double  m_GutschriftExceedValue;

  int m_nStartNummer;
  afx_msg void OnDtnDatetimechangeDatumVon(NMHDR* pNMHDR, LRESULT* pResult);
  afx_msg void OnDtnDatetimechangeDatumBis(NMHDR* pNMHDR, LRESULT* pResult);

  afx_msg void OnEnChangeEditStartnr();
  afx_msg void OnBnClickedRechnungenHaendler();
};
