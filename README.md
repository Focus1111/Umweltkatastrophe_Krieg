# Notfallvorrat – Umweltkatastrophe & Krisenvorsorge

Sammlung aus Richtwerten, Rechenhilfen, Checklisten, Druckvorlagen und Medikamenten-Tools für die Vorsorge eines Haushalts. Ziel ist eine **10-tägige Selbstversorgung** ohne Strom, ohne Leitungswasser und ohne Versorgung von außen. Zusätzlich gibt es Richtwerte für 30 und 90 Tage.

> **Hinweis:** Die Werte sind Richtwerte und keine amtlichen Vorgaben. Laut Projektdateien wurden die Quellen bei der Erstellung **nicht online nachgeprüft**. Vor dem Ausdruck oder der Nutzung bitte die Angaben beim BBK und mit Ärztin oder Arzt bzw. Apotheke prüfen.

## Inhalt

- [Zielgruppe und Umfang](#zielgruppe-und-umfang)
- [Projektstruktur](#projektstruktur)
- [Schnellstart](#schnellstart)
- [Grundannahmen](#grundannahmen)
- [Bedarf pro Person](#bedarf-pro-person)
- [Lebensmittelvorrat für 10 Tage](#lebensmittelvorrat-für-10-tage)
- [Vorrat für mehrere Personen](#vorrat-für-mehrere-personen)
- [Ausrüstung und Haushaltsvorrat](#ausrüstung-und-haushaltsvorrat)
- [Wasser](#wasser)
- [Medikamente](#medikamente)
- [Lagerkontrolle und Einkauf](#lagerkontrolle-und-einkauf)
- [Notfallplan ohne Strom](#notfallplan-ohne-strom)
- [Sonderbedarf und Ergänzungen](#sonderbedarf-und-ergänzungen)
- [Sicherheitshinweise](#sicherheitshinweise)
- [Checkliste](#checkliste)
- [Quellen](#quellen)
- [Haftungsausschluss](#haftungsausschluss)

## Zielgruppe und Umfang

Das Projekt richtet sich an Haushalte, die sich auf einen Ausfall von Wasser-, Strom- und Lebensmittelversorgung vorbereiten möchten. Es deckt ab:

- Berechnung des Bedarfs nach Körpergewicht, Personenzahl und Zeitraum
- Einkaufs- und Ausrüstungsliste mit Kosten
- Lagerkontrolle mit Haltbarkeitsdatum
- Medikamentenvorrat mit Reichweite, Rezeptgültigkeit und Kühlpflicht
- Notfallplan für Stromausfall und fehlendes Wasser
- Druckfertige Unterlagen und Notfallkarten

## Projektstruktur

| Datei | Typ | Inhalt |
|---|---|---|
| `Notfallvorrat_Rechner.xlsx` | Excel | Hauptrechner mit 8 Blättern (siehe unten) |
| `Medikamenten_Vorrat.xlsx` | Excel | Medikamentenvorrat pro Person (3 Blätter) |
| `Notfallvorrat_pro_Person.md` | Markdown | Ausführliche Quelle der Richtwerte und Texte |
| `Notfallvorrat_Ratgeber.docx` | Word | Ratgeber mit Kapiteln 1 bis 6 (Bezugsrahmen bis Wasser) |
| `Notfallvorrat_Praesentation.pptx` | PowerPoint | 10 Folien zu Bedarf, Vorrat, Ausrüstung und Lagerkontrolle |
| `Notfallvorrat_Druckversion.pdf` | PDF | 7 Seiten: Bedarf, Berechnung, Ausrüstung, Checkliste, Notfallplan |
| `Notfallvorrat_Gesamt.pdf` | PDF | 30 Seiten: Gesamtdokument |
| `Notfallvorrat_30_Tage_pro_Person.pdf` | PDF | 2 Seiten: Überlebensvorrat für 30 Tage |
| `Notfallvorrat_90_Tage_pro_Person.pdf` | PDF | 3 Seiten: Überlebensvorrat für 90 Tage (Ziel über Monate) |
| `Notfallkarte.pdf` | PDF | 1 Seite: Notrufnummern und Kernwerte pro Person |
| `Medikamenten_Notfallkarte.pdf` | PDF | 5 Seiten: Hinweise zu Vorrat, Lagerung, Kühlung und Notfalldaten pro Person |
| `Medikamentenplan_ausfuellbar.pdf` | PDF | 4 Seiten: ausfüllbarer Medikamentenplan für Personen 1 bis 4 |
| `diagramm_bedarf.png` | Grafik | Visualisierung des Bedarfs |
| `diagramm_lebensmittel.png` | Grafik | Visualisierung der Lebensmittelmengen |
| `diagramm_wasser.png` | Grafik | Visualisierung des Wasserbedarfs |

### Aufbau des Hauptrechners (`Notfallvorrat_Rechner.xlsx`)

| Blatt | Funktion |
|---|---|
| **Bedarf** | Eingaben (gelbe Felder): Körpergewicht, Aktivitätsfaktor, kcal je kg, Wasser je kg, Hygienewasser, Zeitraum. Berechnet Energie, Trinkwasser, Lebensmittel und Deckungsgrad. Enthält außerdem einen Vergleich für 65, 97,5 und 130 kg. |
| **Haushalt** | Name und Gewicht pro Person eintragen. Berechnet Lebensmittel, Trinkwasser und Hygienewasser je Person sowie Summen. Der Zeitraum kommt aus dem Blatt Bedarf. |
| **Lebensmittel** | Basisliste für 2.000 kcal pro Tag und 10 Tage mit Summen. |
| **Ausruestung** | Ausrüstung mit Spalten „Pro Person“, „Pro Haushalt“ und „Gesamt“. Die Anzahl Personen kommt automatisch aus dem Blatt Haushalt. |
| **Einkaufsliste** | Menge, Einheit, Preis je Einheit und Summe. Summe wird automatisch berechnet. |
| **Lagerkontrolle** | Kaufdatum und Haltbar-bis-Datum. Status: ok, bald prüfen (90 Tage oder weniger) oder ABGELAUFEN. |
| **Wasservergleich** | Liter pro Person und Tag: BBK (2 L), Richtwerte nach Gewicht und Sphere (15 L). |
| **Referenzen** | Quellen und Annahmen. |

### Aufbau der Medikamententabelle (`Medikamenten_Vorrat.xlsx`)

| Blatt | Funktion |
|---|---|
| **Vorrat** | Eine Zeile pro Medikament: Person, Name, Wirkstoff und Stärke, Verbrauch pro Tag, Vorrat in Stück. Berechnet werden Reichweite in Tagen, Enddatum, Status Vorrat und Status Rezept. |
| **Kühlung** | Lagerung bei Stromausfall, z. B. Insulin (Kühlbox mit Kühlakkus), Augentropfen, Zäpfchen, Tabletten (trocken, unter 25 °C). |
| **Hinweise** | Anleitung und Statuserklärung. |

Statuswerte im Blatt Vorrat:

- **NACHBESTELLEN:** Vorrat reicht weniger als 14 Tage
- **bald prüfen:** Vorrat reicht weniger als 30 Tage
- **Rezept erneuern:** Rezept läuft in weniger als 30 Tagen ab
- **ABGELAUFEN:** Packung oder Rezept ist abgelaufen

## Schnellstart

1. **Personen erfassen:** Im Blatt *Haushalt* Name und Körpergewicht jeder Person eintragen. Den Zeitraum im Blatt *Bedarf* festlegen (Standard: 10 Tage).
2. **Bedarf prüfen:** Das Blatt *Bedarf* zeigt Energie, Wasser und Lebensmittelmenge.
3. **Einkaufsliste erstellen:** Mengen aus dem Blatt *Lebensmittel* bzw. *Haushalt* übernehmen, Preise im Blatt *Einkaufsliste* eintragen.
4. **Ausrüstung abhaken:** Blatt *Ausruestung* mit Spalte „Erledigt“ nutzen.
5. **Lagerbestand überwachen:** Haltbarkeit im Blatt *Lagerkontrolle* eintragen und regelmäßig prüfen.
6. **Medikamente pflegen:** `Medikamenten_Vorrat.xlsx` aktualisieren und den ausfüllbaren Medikamentenplan (`Medikamentenplan_ausfuellbar.pdf`) pflegen.
7. **Ausdrucken:** `Notfallvorrat_Druckversion.pdf` und `Notfallkarte.pdf` ausdrucken und im Notfallordner ablegen.

## Grundannahmen

| Parameter | Wert | Bemerkung |
|---|---|---|
| Zeitraum | 10 Tage (Basis), zusätzlich 30 und 90 Tage | Keine Versorgung von außen |
| Körpergewicht | 65 bis 130 kg pro Person | Beispiele: 65, 97,5 und 130 kg |
| Grundumsatz | 24 kcal pro kg und Tag | Annahme |
| Aktivitätsfaktor (Krise) | 1,3 (Bereich 1,2 bis 1,4) | Im Rechner als Eingabe änderbar |
| Trinkwasser | 35 ml pro kg und Tag | Trinken und Kochen |
| Hygienewasser | 2 L pro Tag | Zuschlag |
| Lebensmittelbasis | 2.000 kcal pro Tag | Skalierung siehe unten |
| Salz | ca. 1 Teelöffel pro Tag (ca. 150 g für 30 Tage) | Kochen und Würzen |

## Bedarf pro Person

Formeln:

- Energiebedarf (kcal/Tag) = kg × 24 × Aktivitätsfaktor
- Trinkwasser (L/Tag) = kg × 35 ml ÷ 1.000

| Körpergewicht | Energiebedarf (kcal/Tag) | Trinkwasser (L/Tag) | Wasser 10 Tage (L) | Wasser + Hygiene 10 Tage (L) |
|---|---|---|---|---|
| 65 kg | ca. 2.030 | ca. 2,3 | ca. 23 | ca. 43 |
| 97,5 kg | ca. 3.040 | ca. 3,4 | ca. 34 | ca. 54 |
| 130 kg | ca. 4.060 | ca. 4,6 | ca. 46 | ca. 66 |

Die Werte stimmen mit der Notfallkarte (`Notfallkarte.pdf`) überein. Die Kennzahl „Energie“ ist dort für 65 kg mit 2.028 kcal, für 97,5 kg mit 3.042 kcal angegeben.

Bei Hitze, Krankheit oder Schwerarbeit steigt der Bedarf. Für die Planung lieber etwas mehr einkalkulieren.

## Lebensmittelvorrat für 10 Tage

Basis: 2.000 kcal pro Tag und Person.

| Lebensmittel | Menge | ca. kg | ca. kcal (10 Tage) | Anmerkung |
|---|---|---|---|---|
| Reis, Nudeln, Haferflocken | 1,8 kg | 1,8 | 6.300 | Grundnahrung, lange haltbar |
| Hülsenfrüchte (Linsen, Bohnen) | 0,7 kg | 0,7 | 2.400 | Eiweiß |
| Konserven (Gemüse, Fisch, Fleisch) | 2,5 kg | 2,5 | 2.500 | Mit Dosenöffner |
| Speiseöl | 0,35 L | 0,32 | 3.080 | Energiedichte |
| Zucker oder Honig | 0,3 kg | 0,3 | 1.200 | Schnelle Energie |
| Nüsse, Müsli, Trockenfrüchte | 0,3 kg | 0,3 | 1.350 | Snacks |
| H-Milch oder Milchpulver | 2,0 L | 2,06 | 1.300 | Auch zum Kochen |
| Knäckebrot oder Zwieback | 0,5 kg | 0,5 | 1.750 | Brot-Ersatz |
| **Summe** | | **ca. 8,5 kg** | **ca. 19.880** | ca. 2.000 kcal pro Tag |

### Skalierung nach Körpergewicht

Formel: **Menge = Basismenge × (Energiebedarf ÷ 2.000 kcal)**

| Körpergewicht | Faktor | Lebensmittel für 10 Tage |
|---|---|---|
| 65 kg | x 1,01 | ca. 8,6 kg |
| 97,5 kg | x 1,52 | ca. 12,9 kg |
| 130 kg | x 2,03 | ca. 17,2 kg |

Die Excel-Datei rechnet diese Skalierung automatisch aus.

## Vorrat für mehrere Personen

Formel: **Menge = Menge pro Person × Anzahl Personen × (Tage ÷ 10)**

Bei unterschiedlichen Gewichten wird jede Person einzeln berechnet und die Ergebnisse werden addiert.

**Beispiel:** 3 Personen (97,5 kg, 65 kg, 130 kg) für 10 Tage:

- Lebensmittel: ca. 38,7 kg (ca. 39 kg)
- Wasser inkl. Hygiene: ca. 162 L

Im Blatt *Haushalt* trägst du Name und Gewicht jeder Person ein. Summen und Personenzahl werden automatisch berechnet.

## Ausrüstung und Haushaltsvorrat

Formel für die Ausrüstung: **Gesamt = Pro Person × Anzahl Personen + Pro Haushalt**

### Wasser-Technik

| Artikel | Pro Person | Pro Haushalt | Hinweis |
|---|---|---|---|
| Trinkwasser-Kanister (10 L) | – | 2 | Getrennt lagern, Ablaufdatum beachten |
| Wasserfilter | – | 1 | Für Oberflächen- oder Leitungswasser |
| Entkeimungstabletten | – | 1 Packung | Nach Packungsanleitung |

### Licht und Strom

| Artikel | Pro Person | Pro Haushalt | Hinweis |
|---|---|---|---|
| Kurbel-Taschenlampe | 1 | – | Ohne Batterien |
| Stirnlampe | 1 | – | Hände bleiben frei |
| Kurbelradio mit USB | – | 1 | Empfang von Warnmeldungen |
| Solar-Ladegerät (faltbar, USB) | – | 1 | Nur bei Sonne, lädt langsam |
| Powerbank (mind. 10.000 mAh) | 1 | – | Vorher voll laden |
| Batterien AA | – | 24 | Für Radio und Lampen |
| Batterien AAA | – | 12 | Für Stirnlampen und Kleingeräte |

### Kochen und Wärme

| Artikel | Pro Person | Pro Haushalt | Hinweis |
|---|---|---|---|
| Gaskocher | – | 1 | Nur im Freien oder gut belüftet |
| Gaskartuschen | – | 6 | Ersatz für den Gaskocher |
| Trockenbrennstoff-Kocher (z. B. Esbit) | – | 1 | Reserve, nur gut belüftet |
| Trockenbrennstoff-Würfel | 20 | – | ca. 2 Würfel pro Mahlzeit |
| Kochtopf, Geschirr, Besteck | 1 Satz | – | |
| Dosenöffner (manuell) | – | 1 | Für Konserven |
| Feuerzeug | – | 2 | |
| Wasserfeste Streichhölzer | – | 2 Packungen | |
| Kerzen | – | 10 | Brandgefahr, nie unbeaufsichtigt brennen lassen |
| Schlafsack oder Decke | 1 | – | |
| Warme Kleidung und feste Schuhe | 1 Satz | – | Zusätzlich zur Alltagskleidung |

### Sicherheit

| Artikel | Pro Person | Pro Haushalt | Hinweis |
|---|---|---|---|
| CO-Melder | – | 1 je Gerät | CO-Warnung beim Kochen wichtig |
| Rauchmelder | – | 1 je Gerät | Funktion prüfen |
| Feuerlöscher | – | 1 | Druckanzeige und Prüfdatum kontrollieren |
| Notfalltasche (Rucksack) | 1 | – | Griffbereit, mit Dokumentenkopien |

### Gesundheit und Hygiene

| Artikel | Pro Person | Pro Haushalt | Hinweis |
|---|---|---|---|
| Erste-Hilfe-Set (DIN 13157) | – | 1 | |
| Desinfektionsmittel | – | 1 | |
| Medikamente für 14 Tage | 1 Satz | – | Ärztliche Rücksprache |
| Hygieneartikel | 1 Satz | – | Seife, Zahnpflege, Feuchttücher, Menstruationsprodukte |
| FFP2-Masken | 10 | – | |
| Toilettenpapier | 8 Rollen | – | |
| Müllbeutel | mind. 10 | – | |

### Dokumente und Geld

| Artikel | Pro Person | Pro Haushalt | Hinweis |
|---|---|---|---|
| Kopien von Ausweis, Impfpass, Versicherung | 1 Satz | – | Wasserdicht verpacken |
| Bargeld in kleinen Scheinen | – | 1 Satz | Kartenzahlung kann ausfallen |

## Wasser

Die BBK-Empfehlung deckt nur Trinken und einfaches Kochen ab. Bei längerer Versorgungsunterbrechung sollte man sich am Sphere-Richtwert orientieren.

| Maßstab | Liter pro Person und Tag |
|---|---|
| BBK Trinkwasser (Trinken und Kochen) | 2,0 |
| 65 kg (Berechnung) | ca. 2,3 |
| 97,5 kg (Berechnung) | ca. 3,4 |
| 130 kg (Berechnung) | ca. 4,6 |
| Sphere-Minimum (inkl. Hygiene) | 15,0 |

Empfehlung:

- **Mindestens 20 L Trinkwasser pro Person** (BBK: 2 L pro Tag × 10 Tage)
- **Empfohlen: 30 bis 50 L pro Person** inklusive Hygiene und Kochen

### Wasser gewinnen

- Warmwasserboiler (Kaltwasser-Ablass) nutzen, Badewanne und Eimer füllen.
- Regenwasser filtern und abkochen.
- Schnee schmelzen, Oberflächenwasser filtern.
- **Abkochen:** 1 Minute (3 Minuten über 2.000 m Höhe).
- **SODIS:** mind. 6 Stunden in der Sonne in transparenten PET-Flaschen.

## Medikamente

- Pro Person eine Liste mit Name, Wirkstoff und Stärke, Dosis, Einnahmezeit und Vorrat in Tagen führen.
- Rezepte oder Kopien der Verordnungen im Notfallordner aufbewahren.
- Rezepte rund 1 bis 2 Wochen vor dem Ende des Vorrats anfordern. Die Apotheke kann bei der Planung helfen.
- Ersatzbrille, Kontaktlinsen-Pflege und Hörgerätebatterien mitführen.
- **Kühlpflichtige Medikamente** (z. B. Insulin, Augentropfen, Zäpfchen, flüssige Arzneien): Kühlbox mit Kühlakkus nutzen und Packungsbeilage beachten. Insulin nicht einfrieren.
- **Trockene Tabletten und Kapseln:** trocken, unter 25 °C, in der Originalpackung lagern.
- **Notfallmedikamente** (z. B. Adrenalin-Pen bei Allergien) nur nach ärztlicher Verordnung.
- Medikamentenplan (`Medikamentenplan_ausfuellbar.pdf`) bzw. `Medikamenten_Vorrat.xlsx` regelmäßig aktualisieren.
- Die Medizinische Notfallkarte (`Medikamenten_Notfallkarte.pdf`) enthält Felder für Name, Geburtsdatum, Blutgruppe, Allergien, Vorerkrankungen, Hausarzt, Apotheke, Notfallkontakt, Hilfsmittel und Patientenverfügung.

## Lagerkontrolle und Einkauf

- Im Blatt *Lagerkontrolle* für jeden Artikel Kaufdatum und Haltbar-bis-Datum eintragen.
- Der Status zeigt **ok**, **bald prüfen** (90 Tage oder weniger) oder **ABGELAUFEN**.
- Im Blatt *Einkaufsliste* Preise je Einheit eintragen. Die Summe wird automatisch berechnet.
- Vorrat trocken, kühl und dunkel lagern. Bei Hochwassergefahr im Obergeschoss lagern.

## Notfallplan ohne Strom

### Stufenplan

| Zeitraum | Maßnahmen |
|---|---|
| Stunde 0 bis 1 | Notfallkarte und Notfalltasche bereitlegen, Wasser sichern, Kühlschrank und Gefriertruhe geschlossen halten, Stromkreise prüfen |
| Tag 1 bis 3 | Warnmeldungen über Radio und NINA verfolgen, Wasser abkochen oder filtern, sparsam kochen, Kühlgut zuerst verbrauchen |
| Tag 4 bis 10 | Vorrat einteilen, Kochen nur an gut belüfteten Stellen, Hygiene sicherstellen, Kontakt mit Nachbarn halten |
| Ab Tag 10 | Versorgungslage prüfen, Vorrat nachlegen, Stromversorgung ausbauen |

### Verhalten im Haushalt

- Kühlschrank geschlossen halten (ca. 4 Stunden kühl). Gefriergut hält etwa 24 bis 48 Stunden.
- Heizung und Gasgeräte nie als Dauerwärme in Innenräumen nutzen.
- Eimer-Toilette bei fehlendem Wasser verwenden.
- Haus-Check: Hauptabsperrungen für Wasser, Gas und Strom kennen und beschriften. Werkzeug (Schraubenschlüssel) bereitlegen. Sicherungskasten beschriften.

### Stromversorgung ausbauen

1. **Stufe 1:** Powerbanks und Kurbelgeräte
2. **Stufe 2:** Mobile Powerstation mit Solarpanel
3. **Stufe 3:** Photovoltaik mit Speicher durch Fachbetrieb
4. **Stufe 4:** Notstromaggregat nur im Freien betreiben

Einspeisung ins Hausnetz nur über eine fest installierte Umschalteinrichtung.

## Sonderbedarf und Ergänzungen

- **Babys (0 bis 12 Monate):** Säuglingsnahrung für 10 Tage, Windeln, Feuchttücher, Fläschchen. Zubereitung nur mit abgekochtem oder abgefülltem Wasser. Stillende haben einen höheren Energiebedarf.
- **Kinder:** Portionen nach Alter, Kinderfiebersaft (Dosierung nach Gewicht), Beschäftigung (Bücher, Spiele), Schnuller und Ersatzkleidung.
- **Pflegebedürftige:** Inkontinenzmaterial, Pflegehilfsmittel, Medikamentenliste, Kopien von Pflegegrad und Versicherung.
- **Haustiere:** Futter für 10 Tage (Menge nach Packungsangabe), Wasser, Transportbox, Leine, Impfpass, benötigte Medikamente.
- **Kommunikation:** Papierliste mit Notrufen (110, 112), Familie, Nachbarn, Hausarzt, Apotheke, Versicherungen und Bank im Geldbeutel und im Notfallordner. Treffpunkt außerhalb des Wohnorts festlegen. Warn-App **NINA** installieren und Benachrichtigungen aktivieren. Ersatz-Powerbank im Rucksack.
- **Mobilität:** Auto bei Warnung mindestens halb voll tanken. E-Auto mindestens 50 % laden. Kraftstoff nur in zugelassenen Kanistern lagern und örtliche Vorschriften beachten. Starthilfekabel, Warndreieck, Warnweste und Decke ins Auto. Fahrrad-Reparaturset mit Flickzeug, Luftpumpe und Multitool.
- **Werkzeug:** Multitool, Taschenmesser, Gewebeband, Seil (ca. 10 m), Arbeitshandschuhe, Kabel und Adapter, Wasserkanister mit Zapfhahn.
- **Schlüssel:** Kopien bei Nachbarn oder Familie hinterlegen.
- **Sehhilfen und Hörgeräte:** Ersatzbrille und Hörgerätebatterien mitführen.

## Sicherheitshinweise

- **Gaskocher und Trockenbrennstoff** nur im Freien oder in gut belüfteten Räumen verwenden. Kohlenmonoxid ist geruchlos und in geschlossenen Räumen lebensgefährlich.
- **CO-Melder** einbauen und regelmäßig prüfen.
- **Kerzen** nie unbeaufsichtigt brennen lassen.
- **Solar-Ladegeräte** funktionieren nur bei Sonne und laden langsam. Powerbank und Kurbelradio ergänzen sie.
- **Batterien und Brennstoff:** Verbrauch hängt von den Geräten ab. Bei Bedarf mehr einplanen.

## Checkliste

- [ ] Wasser: mind. 20 L pro Person, empfohlen 30 bis 50 L inkl. Hygiene und Kochen
- [ ] Lebensmittel für 10 Tage (ca. 2.000 kcal pro Tag und Person)
- [ ] Kochmöglichkeit ohne Strom (Gaskocher oder Trockenbrennstoff, gut belüftet)
- [ ] Licht, Radio, Batterien, Powerbank
- [ ] Erste-Hilfe-Set, Medikamente, Hausapotheke
- [ ] Bargeld und Dokumentenkopien
- [ ] Hygieneartikel und Müllbeutel
- [ ] Decken, warme Kleidung, Schlafsack
- [ ] Notfalltasche pro Person
- [ ] Haus-Check: Absperrungen kennen, Melder prüfen
- [ ] NINA installiert, Papierliste mit Notrufen im Geldbeutel

## Quellen

- **BBK:** Ratgeber für Notfallvorsorge und Broschüre „Vorsorge für Krisen- und Katastrophenlagen“
- **Sphere-Standard:** Kapitel Wasser, Sanitär und Ernährung (Mindestwasser 15 L pro Person und Tag inkl. Hygiene, Mindestenergie 2.100 kcal pro Tag)
- **DIN 13157:** Erste-Hilfe-Set

Die Annahmen (Grundumsatz 24 kcal/kg, Aktivitätsfaktor 1,3, Wasser 35 ml/kg) sind Richtwerte aus der Projektdokumentation und wurden nicht gegen eine amtliche Quelle geprüft.

## Haftungsausschluss

Die Inhalte dienen der allgemeinen Information und ersetzen keine ärztliche, apothekerliche oder behördliche Beratung. Richtwerte können je nach Person und Situation abweichen. Medikamente und Notfallmedikamente nur nach Verordnung und Rücksprache mit Ärztin oder Arzt verwenden.
