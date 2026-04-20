/** SharePoint TA-Liste item */
export interface ITaItem {
  ID: number;
  Title: string;           // TA-Nr.
  field_1?: string;         // StatusOld (legacy)
  field_2?: string;         // Bemerkung
  field_4?: string;         // Erstellungsdatum
  field_5?: string;         // Wunschtermin
  field_6?: string;         // geplanter Termin
  field_8?: string;         // Kunde
  field_9?: string;         // Endkunde
  field_10?: number;        // Kundennummer
  field_11?: string;        // Anwendung
  field_12?: string;        // Material
  field_13?: number;        // Priorität (1=Niedrig, 2=Mittel, 3=Hoch)
  field_14?: number;        // Potenzial [jato]
  field_15?: number;        // Chance [%]
  field_16?: string;        // Kategorie
  field_17?: string;        // Budget bei Start
  field_18?: number;        // IST Kosten TA
  field_19?: number;        // geplante Kosten TA
  field_20?: number;        // Differenz
  field_21?: string;        // Grund für Verschiebung
  field_22?: string;        // Ursprünglicher Termin
  Erledigungsdatum?: string; // Erledigungsdatum (IST)
  Status?: string;          // Choice: Termin planen | läuft planmäßig | prüfen | überfällig | abgeschlossen
  Ersteller?: { Title: string; EMail: string };
  Verantwortlicher?: { Title: string; EMail: string };
  Modified?: string;
  Aufgabenstellung?: string; // Aufgabenstellung (Note)
  Projektnummer?: number;    // VPro-Nummer aus Projektliste
}

export interface IKategorieItem {
  ID: number;
  Title: string; // Kategorie
  Email?: string; // Ziel-Email bei TA-Anlage
}

export interface IKundenAnwendungItem {
  ID: number;
  Title: string; // Kunde
  Anwendung: string;
}

/** SharePoint Projekt-Liste item */
export interface IProjektItem {
  ID: number;
  Title: string;            // VPro
  kontaktNr?: string;        // Kontaktnummer (Spalte "Kontakt")
  field_1?: string;          // Kunde
  field_2?: string;          // Name
  field_3?: string;          // Name 2
  field_7?: string;          // Bezeichnung (Anwendung)
  field_8?: number;          // Potential
  field_10?: number;         // Chance (Anfrage)
  field_11?: number;         // Budget
  field_12?: string;         // SegCode
  field_13?: string;         // Endkunde
  field_16?: string;         // Material
  field_18?: number;         // Ziel VK (Preis)
  field_21?: string;         // SOP (BeginnDat)
  vps?: string;              // VPS (Priorität: 1=Niedrig, 2=Mittel, 3=Hoch)
}

/** New TA form data */
export interface INewTaForm {
  kunde: string;
  endkunde: string;
  kontaktNr: string;
  aufgabe: string;
  bemerkung: string;
  anwendung: string;
  material: string;
  kategorie: string;
  potenzial: string;
  chance: string;
  budget: string;
  wunschtermin: string;
  verantwortlicherId?: number | null;
  zielpreis: string;
  sop: string;
  segCode: string;
  antwortIn: "Deutsch" | "Englisch";
  prioritaet: string;
  projektNr: string;
}

/** Status constants */
export const STATUS_VALUES = [
  'Termin planen',
  'läuft planmäßig',
  'prüfen',
  'überfällig',
  'abgeschlossen'
] as const;

export type TaStatus = typeof STATUS_VALUES[number];

/** Postponement reasons */
export const VERSCHIEBUNG_GRUENDE = [
  'TA-Aufgabe nachträglich erweitert',
  'Fehlende Infos / Rückmeldung Kunde',
  'Verzögerung durch EMS',
  'Verzögerung durch externen Lieferanten',
  'Verzögerung durch internen Lieferanten',
  'Kapazitätsengpass im SG',
  'Kapazitätsengpass im TKS',
  'Materialverfügbarkeit intern',
  'Materialverfügbarkeit extern',
  'Umverteilung Prioritäten',
  'Technischer Fehler TA-Liste',
  'Selbstorganisation',
  'Musterungstermin verschoben'
];

/** Delay reasons for initial planning */
export const INITIAL_DELAY_REASONS = [
  'Rückfragen',
  'Selbstorganisation',
  'Komplexität',
  'Kapazität der Abteilung',
  'kein Projektbudget vorhanden',
  'Technische Probleme(Nur TA Liste)'
];

/** View enum for navigation */
export enum AppView {
  Dashboard = 'dashboard',
  NeueTa = 'neue-ta',
  TerminPlanen = 'termin-planen',
  TaDetail = 'ta-detail',
  AlleTas = 'alle-tas',
  Statistik = 'statistik',
  Settings = 'settings'
}

/** KPI data */
export interface IKpiData {
  overdue: number;
  plan: number;
  onTrack: number;
  review: number;
}
