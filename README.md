# Intune Management Tool

Een alleen-lezen Windows-tool met een Intune-achtige interface. Je ziet per policy, script, app of profiel aan wie het is toegewezen, apart voor include en exclude. Daarnaast bevat de tool een dashboard, ASR-overzicht, zoeken per groep of gebruiker, detail- en instellingenweergave en export. Er wordt niets in Intune gewijzigd.

## Starten

Dubbelklik op `Start-IntuneTool.bat`. De eerste keer wordt gevraagd of de module `Microsoft.Graph.Authentication` mag worden geïnstalleerd. Klik daarna op **Aanmelden met Microsoft**. Na het inloggen worden de gegevens automatisch geladen. Met **Gegevens laden** ververs je ze, en met **Annuleren** onderbreek je een lopende actie.

## De interface

Links staat de navigatie, net als in Intune. Rechts staat de inhoud.

| Groep | Onderdelen |
|---|---|
| Overzicht | Dashboard, Aandachtspunten, Alle assignments, Zoeken op groep/gebruiker |
| Apparaten | Configuration profiles, Compliance policies, Platform scripts, Remediations, Windows updates, Autopilot profielen, Enrollment, Apparaten |
| Endpoint security | Alle endpoint security, ASR-regels, Attack surface reduction, Antivirus, Firewall, Disk encryption, EDR, Account protection, Security baselines |
| Apps | Apps, App protection policies, App configuration policies |
| Groepen en filters | Groepen in gebruik, Assignment filters |

Achter elk onderdeel staat het aantal items. In elke lijst kun je zoeken, filteren op platform en alleen items tonen die niet (volledig) zijn toegewezen. Kolommen sorteer je door op de kop te klikken. Dubbelklik op een regel voor details.

## Wat de tool laat zien

**Assignments.** Elke regel heeft een kolom met opgenomen groepen (include) en een kolom met uitgesloten groepen (exclude). Assignment filters staan erbij als `[filter: Naam (include)]`, en bij apps de intent (Required, Available, Uninstall). De status is *Toegewezen*, *Niet toegewezen* of *Alleen uitsluiting*.

**Dashboard.** Kaarten met totalen, niet-toegewezen items, groepen in gebruik, fouten en ASR-conflicten, plus een tabel per type. Klik op een kaart om naar de bijbehorende lijst te gaan.

**Aandachtspunten.** Items die niet zijn toegewezen of alleen uitsluitingen hebben, verwijzingen naar groepen die niet meer bestaan, groepen die zowel include als exclude zijn, en ASR-regels met conflicterende modi. Met de knop *Ledenaantal ophalen* worden ook groepen zonder leden herkend.

**ASR-regels.** Elke Attack Surface Reduction-regel per policy met modus (Block, Audit, Warn, ...), include en exclude. Komt dezelfde regel in meerdere policies voor, dan staat er *Overlap* of *CONFLICT* bij met de betrokken policies. Met *Scan ook Settings catalog* worden ook alle Settings catalog-profielen doorzocht, omdat ASR daar vaak ook wordt ingesteld.

**Zoeken op groep/gebruiker.** Typ het begin van een naam of UPN en kies een resultaat. Je ziet alles wat eraan is toegewezen, ook via bovenliggende (geneste) groepen en via "Alle gebruikers/apparaten". De kolom *Effect* laat zien of een uitsluiting wint van een opname. Vanuit *Groepen in gebruik* kom je met een dubbelklik direct in deze weergave.

**Details.** Dubbelklik op een item voor tabbladen met assignments, eigenschappen en instellingen (Settings catalog met de echte Intune-namen, administrative templates, endpoint security, OMA-URI en overige eigenschappen). Bij platform scripts en remediations zie je de scriptinhoud. Bij een apparaat zie je de status van zijn configuration- en compliance-policies.

**Export.** Elke lijst exporteer je naar Excel (.xlsx, Excel hoeft niet geïnstalleerd te zijn), CSV (puntkomma) of een HTML-rapport met zoeken en sorteren. Vanaf het dashboard maak je een volledig rapport, in Excel met meerdere bladen, waaronder een plat overzicht met één regel per assignment voor draaitabellen.

## Rechten

Alle rechten zijn alleen-lezen (delegated):

| Recht | Gebruikt voor |
|---|---|
| `DeviceManagementConfiguration.Read.All` | profielen, compliance, scripts, endpoint security, filters, updates |
| `DeviceManagementApps.Read.All` | apps, app protection, app configuration |
| `DeviceManagementServiceConfig.Read.All` | Autopilot en enrollment |
| `DeviceManagementManagedDevices.Read.All` | apparaten en policy-status per apparaat |
| `Group.Read.All` | groepsnamen |
| `GroupMember.Read.All` | ledenaantal en lidmaatschappen |
| `User.Read.All` | gebruikers zoeken |

Vereist je tenant admin-consent voor de Microsoft Graph PowerShell-app, dan moet een beheerder dat eenmalig goedkeuren. Is een onderdeel niet beschikbaar (bijvoorbeeld door een ontbrekend recht of licentie), dan meldt de tool dat na het laden en blijven de overige onderdelen werken.

## Opbouw

| Bestand | Inhoud |
|---|---|
| `IntuneTool.ps1` | opstarten, navigatie opbouwen, events koppelen |
| `lib\Graph.ps1` | Graph-aanroepen met retry, alle bronnen, groepen en filters, laden |
| `lib\Settings.ps1` | instellingen, scripts en ASR-regels uitlezen |
| `lib\Analysis.ps1` | dashboard, aandachtspunten, groepen- en filteroverzicht, zoeken |
| `lib\Export.ps1` | xlsx-, csv- en html-export |
| `lib\Ui.ps1` | weergaven, tabellen, dashboard, detailvensters, aanmelden |
| `ui\MainWindow.xaml`, `ui\Resources.xaml` | layout en stijlen |

## Uitbreiden

Een nieuw type toevoegen is meestal één blok in `$script:Sources` in `lib\Graph.ps1` (endpoint, type, subtype, platform). Een nieuw navigatie-onderdeel is één regel in `$script:Views` in `lib\Ui.ps1`.

## Goed om te weten

- De tool gebruikt de beta-endpoints van Graph. Die kunnen op termijn veranderen.
- Het laden gebeurt in het hoofdvenster. De statusbalk toont de voortgang en het venster blijft bediend via **Annuleren**. Bij grote tenants kan het even duren.
- Assignments worden zo veel mogelijk in één keer opgehaald met `$expand=assignments`. Lukt dat niet voor een bron, dan valt de tool terug op een aanroep per item.
- Het ledenaantal van groepen wordt alleen op verzoek opgehaald, omdat dit per groep een aparte aanroep kost.
- Apparaten worden pas geladen als je die pagina opent.
