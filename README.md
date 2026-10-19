# CiscoStackAnalyzer
**Dokumentation för SwitchInventory v2.2**  
PowerShell 5.1 • Plink (PuTTY) • ImportExcel • Airgappad Windowsmiljö

SwitchInventory inventerar Cisco-/industriswitchar via SSH och exporterar en Excelrapport med **en rad per fysisk switch/stackmedlem**. Det kan skanna alla enheter i `switches.csv` eller en enda switch. Upstream/core-anslutningar identifieras via CDP mot listan i `coreswitches.txt`.

> **Status:** Stödet är utvecklat med verkliga CLI-exempel för Catalyst 2960X, 9300, 9500, IE-3000, IE-3300 och Rockwell/Stratix 1783. Det finns också preliminärt stöd för Catalyst 3560 och 3850. Resultat bör kontrolleras mot ett urval enheter innan bred användning.

## Switchar som stöds

| Serie | Modeller (exempel) | Status |
|---|---|---|
| Catalyst 2960X | `WS-C2960X-8/12/24/48...` | Verifierad mot verkliga CLI-exempel |
| Catalyst 9200 | `C9200-8/12/24/48...` | Stöd för portantal, preliminärt för 9200 |
| Catalyst 9300 | `C9300-24/48...` (stackar) | Verifierad mot verkliga CLI-exempel |
| Catalyst 9500 | `C9500-8/12/24/32/48...` | Verifierad mot verkliga CLI-exempel |
| Catalyst 3560 | `WS-C3560-8/12/24/48...` | Preliminärt stöd |
| Catalyst 3850 | `WS-C3850-12/24/48...` | Preliminärt stöd |
| Cisco IE-3000 | `IE-3000-...` | Verifierad mot verkliga CLI-exempel |
| Cisco IE-3300 | `IE-3300-8T2S...` | Verifierad mot verkliga CLI-exempel |
| Rockwell/Stratix | `1783-...` (inkl. expansionsmoduler) | Verifierad mot verkliga CLI-exempel |

Övriga Cisco IOS/IOS-XE-switchar kan fungera, men då bestäms portantalet utifrån observerade fysiska portar i `show interfaces status` och resultatet bör kontrolleras manuellt. Kravet är att enheten svarar på `show version`, `show inventory`, `show interfaces status` och `show cdp neighbors detail` över SSH.

## 1. Filstruktur

```text
Cisco-StackAnalyzer/
├── SwitchInventory_v2.2.ps1
├── switches.csv
├── coreswitches.txt
├── putty/
│   └── plink.exe
├── Modules/
│   └── ImportExcel/
│       └── 7.x.x/
│           ├── ImportExcel.psd1
│           ├── ImportExcel.psm1
│           └── ... (resten av modulen inklusive DLL:er)
├── Reports/                 # Skapas vid export
└── hostkeys.json            # Skapas automatiskt vid första SSH-anslutningar
```

`ImportExcel.psd1` får även ligga direkt i `Modules/ImportExcel/` utan versionsmapp. Hela modulen måste kopieras. Ingen lokal installation av Microsoft Excel behövs.

## 2. Förutsättningar

- Windows PowerShell **5.1**.
- Nätverksåtkomst via SSH (vanligtvis TCP 22) och ett Cisco-/AAA-konto som får köra relevanta `show`-kommandon.
- `putty/plink.exe` som stöder `-pwfile` och `-hostkey`.
- PowerShell-modulen **ImportExcel** i den lokala katalogen ovan eller installerad för aktuell användare.
- Godkänt sätt att överföra Plink och ImportExcel till airgappad miljö. Verifiera filernas ursprung enligt organisationens rutiner.
- Skrivbehörighet i projektkatalogen för `hostkeys.json` och `Reports/`, och till användarens temporära katalog.

## 3. Konfiguration

### switches.csv

Lägg CSV-filen **i samma mapp som scriptet**. Rubrikerna måste vara exakt:

```csv
Name,IPAddress,Cabinet
SEUPPSW121,172.27.20.121,CA03
SEUPPSW195,172.27.20.195,CA04
SEUPPSW155,172.27.20.155,PROD01
```

- `Name`: switchens hostname (eller ert eget inventeringsnamn).
- `IPAddress`: adressen som Plink ansluter till.
- `Cabinet`: skåp/placering; används i rapportens detaljflikar.
- Använd **komma** som avgränsare och spara helst som UTF-8.
- Lägg inte in tabellmarkeringar (`|`) eller extra rubrikrader.

### coreswitches.txt

Lägg filen **i samma mapp som scriptet**. Skriv ett kort hostname per rad:

```text
# Godkända upstream-/core-switchar
SEUPPSW001
SEUPPSW010
SEUPPSW12
SEUPPSW193
```

- Rader som börjar med `#` och tomma rader ignoreras.
- Det går att lägga till och ta bort switchar utan att ändra PowerShell-scriptet.
- CDP-namn matchas mot kortnamn; `SEUPPSW010.example.local` matchar `SEUPPSW010`.
- **Exakt hostname**, inte delsträng: `SEUPPSW12` ska inte antas vara samma som `SEUPPSW120`.
- Filen måste innehålla minst en switch, annars avbryts scanning.
- Om en uppströms switch inte finns i filen visas den inte som core-uplink i kolumnerna H–K.

## 4. Starta verktyget

Öppna PowerShell i projektmappen:

```powershell
# Om filen är nedladdad och RemoteSigned tillåter det:
Unblock-File -LiteralPath '.\SwitchInventory_v2.2.ps1'

.\SwitchInventory_v2.2.ps1
```

`Unblock-File` hjälper vid Windows nedladdningsmarkering, men upphäver **inte** centralt styrd `AllSigned`-policy. Följ företagets säkerhetspolicy.

Menyn i v2.2:

```text
1. Registrera SSH-hostnycklar (valfri manuell verifiering)
2. Scanna alla switchar och exportera Excel (auto-TOFU)
3. Scanna en specifik switch (CSV eller IP, auto-TOFU)
4. Visa core-switchar
5. Redigera coreswitches.txt (sparas permanent)
6. Om autentisering och säkerhet
7. Avsluta
```

Vid menyval **2** eller **3** visas `Get-Credential`: ange ditt tillåtna administratörskonto och lösenord. Vid val **3** kan du välja switch med namnet i CSV-filen eller skriva in IP-adressen. En adress utanför CSV kan anges tillsammans med valfritt namn och skåp.

## 5. Excelrapport

Rapporten sparas i `Reports/` och innehåller dessa flikar:

| Flik | Innehåll |
|---|---|
| **Inventory** | Huvudrapport, en rad per fysisk switch/stackmedlem, kolumner A–K |
| **Overview** | Översikt per administrerad switch/stacksystem |
| **StackMembers** | Enheter, modell, serienummer, stackroll och portantal |
| **CoreUplinks** | Identifierade fysiska uplinks mot konfigurerade core-switchar |
| **ScanLog** | Resultat och fel per scannad enhet |

**Kolumner A–K i `Inventory`:**

| Kolumn | Fält |
|---|---|
| A | Serial number |
| B | Device Name |
| C | IP Address |
| D | Model |
| E | Stack Switch |
| F | Port Count (accessportar, t.ex. `48 Port`) |
| G | Switch Role (`Core`, `Distribution` eller `Access`) |
| H | Upstream Connected Switch Name Details |
| I | Upstream IP Address |
| J | Uplink Port #s (lokalt interface) |
| K | SFP used (mediatyp från `show interfaces status`) |

Stackobjekt, portkanaler (`Po`), virtuella interface och applikationsportar (`Ap`) är inte egna fysiska switchar eller accessportar. Expansionsmoduler på IE/1783 räknas till samma fysiska enhet. Upplänkar hämtas från **CDP**, inte interfacebeskrivningen. Om en medlem har flera länkar grupperas de i motsvarande celler i H–K.

> För enheter som även är core-switchar kan rollen vara `Core` och en upstream-anslutning ändå visas om CDP-grannen finns i core-listan.

## 6. Inloggning, temporära filer och hostnycklar

**Hantering av lösenord i korthet:**

1. Användarnamn och lösenord anges i en inloggningsruta (`Get-Credential`) varje gång du startar en scanning (menyval 2 eller 3). De lagras aldrig i en konfigurationsfil, i scriptet eller i projektmappen.
2. Uppgifterna hålls bara i minnet under den pågående scanningen.
3. När scanningen är klar, avbryts eller misslyckas rensas uppgifterna ur scriptets variabler. Nästa körning kräver ny inloggning.
4. För varje SSH-anslutning skapas en tillfällig lösenordsfil som Plink läser (`-pwfile`). Filen ligger i en låst temporär mapp och raderas direkt när anslutningen är klar. Lösenordet skickas aldrig som kommandoradsargument.

- **Ingen permanent `credentials.xml` behövs i v2.2.** Uppgifter anges på nytt för varje scanning och hålls som `PSCredential` under körningen.
- Plink kräver en kortlivad lösenordsfil via `-pwfile`. Scriptet skapar en unik temporär mapp med begränsad NTFS-ACL, och försöker radera den när anslutningen avslutas, även vid fel.
- Lösenordet ligger **tillfälligt i klartext på disk**. Radering innebär inte garanterad säker radering; använd bara detta upplägg om det accepteras av er säkerhetspolicy. Skydda användarkonton, tempkatalog och hostmaskin.
- För nya hostnycklar använder verktyget **TOFU (Trust On First Use)** och sparar det observerade SHA-256-fingerprintet i `hostkeys.json`. Detta innebär att den **första** anslutningen inte verifieras via en oberoende källa.
- Om en tidigare sparad nyckel byts ut ska anslutningen misslyckas i stället för att acceptera ändringen automatiskt. Utred och verifiera förändringen innan eventuell återregistrering.
- För högre säkerhetskrav: använd kontrollerade hostnycklar och överväg SSH-nyckelautentisering där plattform och AAA stödjer det.

## 7. Kommandon och plattformar

Scriptet hämtar primärt information genom:

```text
terminal length 0
show version
show switch
show inventory
show interfaces status
show cdp neighbors detail
```

Alla modeller stöder inte `show switch`. En fristående IE-switch kan därför rapporteras som medlem `1`. Stöd finns för olika inventorynamn (`1`, `Switch 1`, `Switch 1 Chassis`, `Chassis`) och för två- respektive tredelade interfacenamn (`Gi1/9`, `Gi1/0/49`).

## 8. Felsökning

| Symptom | Kontroll/åtgärd |
|---|---|
| `ImportExcel saknas` | Kontrollera `Modules/ImportExcel/.../ImportExcel.psd1` och att samtliga modulfiler är med. |
| `plink.exe hittades inte` | Kontrollera `putty/plink.exe`. |
| `Configured password was not accepted` | Prova Plink interaktivt med samma konto; verifiera lösenord, användarnamnsformat och AAA-behörighet. |
| `Cannot confirm a host key` | Kontrollera hostnyckelhanteringen och om `hostkeys.json` kan skrivas/läsas. Ändrade hostnycklar ska utredas, inte ignoreras. |
| Felaktig modell/serienummer | Jämför `show inventory` och `show version` med `StackMembers`-fliken. |
| Fel portantal | Kontrollera access-/uplinkindelning och expansionsmoduler. |
| Inga uplinks i Excel | Kontrollera CDP samt exakt kortnamn i `coreswitches.txt`. |
| `UnauthorizedAccess` vid scriptstart | Kontrollera `Get-ExecutionPolicy -List` och företagets policy. |

**Loggning i v2.2:** `ScanLog` ligger i den exporterade Excelrapporten. Scriptet skapar inte automatiskt en separat persistent `.log`-fil; om exporten misslyckas kan loggen därför saknas. Detta är en möjlig förbättring för en senare version.

## 9. Rekommenderad första körning

1. Skapa `switches.csv` och `coreswitches.txt` från exempelfilerna.
2. Kontrollera att `putty/plink.exe` och den kompletta ImportExcel-modulen finns på plats.
3. Kör scriptet och välj **3** för en känd switch.
4. Kontrollera `Inventory`, `CoreUplinks` och `ScanLog`.
5. Kör därefter **2** för hela switchparken.

---

**Underhåll:** Håll `switches.csv` uppdaterad med aktuella IP-adresser och `coreswitches.txt` med aktuella upstream-/core-namn. Lagra inte användarlösenord eller känsliga konfigurationsbackuper i projektkatalogen.
