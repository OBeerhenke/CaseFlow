import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/attachments';
import '@pnp/sp/files'; // Added import for getFileByServerRelativeUrl / getFileByServerRelativePath
import '@pnp/sp/site-users/web';
import '@pnp/sp/profiles';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ITaItem, IProjektItem, INewTaForm, IKpiData, IKategorieItem, IKundenAnwendungItem } from '../models/types';

const TA_LIST = 'Technische_Anfragen';
const PROJEKT_LIST = 'Projekt-Liste';
const KATEGORIEN_LIST = 'TA_Kategorien';
const KUNDEN_ANWENDUNGEN_LIST = 'TA_Kunden_Anwendungen';
const CONFIG_LIST = 'TA_Config';

function pad2(n: number): string {
    return n < 10 ? '0' + n : '' + n;
}

export class SharePointService {
    private static _instance: SharePointService;

    public static init(context: WebPartContext): void {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        sp.setup({ spfxContext: context as any });
        SharePointService._instance = new SharePointService();
    }

    public static get instance(): SharePointService {
        return SharePointService._instance;
    }

    // ─── TA-Liste CRUD ─────────────────────────────────

    public async getAllTAs(): Promise<ITaItem[]> {
        return sp.web.lists.getByTitle(TA_LIST).items
            .select(
                'ID', 'Title', 'field_1', 'field_2', 'field_4', 'field_5', 'field_6',
                'field_8', 'field_9', 'field_10', 'field_11', 'field_12', 'field_13',
                'field_14', 'field_15', 'field_16', 'field_17', 'field_18', 'field_19',
                'field_20', 'field_21', 'field_22', 'Erledigungsdatum', 'Status', 'Modified',
                'Aufgabenstellung', 'Projektnummer',
                'Ersteller/Title', 'Ersteller/EMail',
                'Verantwortlicher/Title', 'Verantwortlicher/EMail'
            )
            .expand('Ersteller', 'Verantwortlicher')
            .top(2000)
            .orderBy('Modified', false)
            .get();
    }

    public async getSiteUsers(): Promise<any[]> {
        // PrincipalType 1 = User. Also filter out system accounts which typically have null/empty email or specific titles.
        return sp.web.siteUsers
            .filter("PrincipalType eq 1 and Email ne ''")
            .select("Id", "Title", "Email")
            .get();
    }

    public async searchUsers(query: string): Promise<any[]> {
        if (!query || query.length < 3) return [];
        try {
            const results = await sp.profiles.clientPeoplePickerSearchUser({
                AllowEmailAddresses: true,
                AllowMultipleEntities: false,
                MaximumEntitySuggestions: 20,
                PrincipalSource: 15, // PrincipalSource.All
                PrincipalType: 1,    // PrincipalType.User
                QueryString: query
            });
            // Only show users whose email local part contains a dot (e.g. John.Doe@...)
            // This filters out system/service accounts with numeric or coded email prefixes
            return results.filter((u: any) => {
                const email: string = u.EntityData?.Email || u.Description || '';
                if (!email || !email.includes('@')) return false;
                const localPart = email.split('@')[0];
                return localPart.includes('.');
            });
        } catch (e) {
            console.error("Error searching users:", e);
            return [];
        }
    }

    public async ensureUser(loginName: string): Promise<number> {
        try {
            const user = await sp.web.ensureUser(loginName);
            return user.data.Id;
        } catch (e) {
            console.error("Error ensuring user:", e);
            throw e;
        }
    }

    public async getOpenTAs(): Promise<ITaItem[]> {
        return sp.web.lists.getByTitle(TA_LIST).items
            .select(
                'ID', 'Title', 'field_2', 'field_4', 'field_5', 'field_6',
                'field_8', 'field_9', 'field_10', 'field_11', 'field_12', 'field_13',
                'field_14', 'field_15', 'field_16', 'field_17', 'field_18', 'field_19',
                'field_20', 'field_21', 'field_22', 'Erledigungsdatum', 'Status', 'Modified',
                'Ersteller/Title', 'Ersteller/EMail',
                'Verantwortlicher/Title', 'Verantwortlicher/EMail'
            )
            .expand('Ersteller', 'Verantwortlicher')
            .filter("Status ne 'abgeschlossen'")
            .top(2000)
            .orderBy('Modified', false)
            .get();
    }

    public async getTaById(id: number): Promise<ITaItem> {
        return sp.web.lists.getByTitle(TA_LIST).items.getById(id)
            .select(
                'ID', 'Title', 'field_1', 'field_2', 'field_4', 'field_5', 'field_6',
                'field_8', 'field_9', 'field_10', 'field_11', 'field_12', 'field_13',
                'field_14', 'field_15', 'field_16', 'field_17', 'field_18', 'field_19',
                'field_20', 'field_21', 'field_22', 'Erledigungsdatum', 'Status', 'Modified',
                'Ersteller/Title', 'Ersteller/EMail',
                'Verantwortlicher/Title', 'Verantwortlicher/EMail'
            )
            .expand('Ersteller', 'Verantwortlicher')
            .get();
    }

    public async createTA(form: INewTaForm, taNr: string): Promise<number> {
        const today = new Date();
        const dateStr = today.getFullYear() + '-' + pad2(today.getMonth() + 1) + '-' + pad2(today.getDate());

        const potVal = form.potenzial ? parseFloat(form.potenzial.replace(',', '.')) : null;
        const chanceVal = form.chance ? parseFloat(form.chance.replace(',', '.')) : null;

        // Convert DD.MM.YYYY from our form into YYYY-MM-DD for SharePoint (Edm.DateTime)
        let wunschIso = '';
        if (form.wunschtermin) {
            const wtParts = form.wunschtermin.split('.');
            if (wtParts.length === 3) {
                wunschIso = `${wtParts[2]}-${wtParts[1]}-${wtParts[0]}`;
            }
        }

        // Get the current user to set as Ersteller
        const currentUser = await sp.web.currentUser.get();

        const item: any = {
            Title: taNr,
            field_8: form.kunde,
            field_9: form.endkunde,
            field_10: form.kontaktNr ? parseInt(form.kontaktNr, 10) : null,
            Aufgabenstellung: form.aufgabe,
            field_2: form.bemerkung,
            field_11: form.anwendung,
            field_12: form.material,
            field_16: form.kategorie,
            field_13: form.prioritaet ? parseInt(form.prioritaet, 10) : null, // Assuming field_13 is Priority
            field_14: !isNaN(potVal as any) ? potVal : null,
            field_15: !isNaN(chanceVal as any) ? chanceVal : null,
            ErstellerId: currentUser.Id,
            VerantwortlicherId: form.verantwortlicherId || null,
            field_17: form.budget,
            Zielpreis: !isNaN(parseFloat(form.zielpreis)) ? parseFloat(form.zielpreis) : null, // NEW Field
            SOP: form.sop,                                // NEW Field
            SegCode: form.segCode,                        // NEW Field
            AntwortIn: form.antwortIn,                    // NEW Field
            Projektnummer: form.projektNr ? parseInt(form.projektNr, 10) : null,
            field_5: wunschIso || undefined,
            field_4: dateStr,
            Status: 'Termin planen'
        };

        // Omitting null or empty fields cleanly for SharePoint
        Object.keys(item).forEach(key => {
            if (item[key] === null || item[key] === undefined || item[key] === '') {
                delete item[key];
            }
        });

        try {
            const result = await sp.web.lists.getByTitle(TA_LIST).items.add(item);
            return result.data.ID as number;
        } catch (e: any) {
            console.error("Error creating TA:", e);
            if (e.isHttpRequestError) {
                const data = await e.response.json();
                console.error("SharePoint REST API Error details:", data);
                alert(`Fehler beim Speichern in SharePoint:\n${data.error?.message?.value || e.message}`);
            } else {
                alert(`Fehler beim Speichern: ${e.message}`);
            }
            throw e;
        }
    }

    public async updateTA(id: number, fields: Partial<Record<string, string | number | undefined>>): Promise<void> {
        await sp.web.lists.getByTitle(TA_LIST).items.getById(id).update(fields);
    }

    public async getAttachments(id: number): Promise<{ FileName: string; ServerRelativeUrl: string }[]> {
        return sp.web.lists.getByTitle(TA_LIST).items.getById(id).attachmentFiles.get();
    }

    public async addAttachment(id: number, fileName: string, content: ArrayBuffer): Promise<void> {
        await sp.web.lists.getByTitle(TA_LIST).items.getById(id).attachmentFiles.add(fileName, content);
    }

    public async deleteAttachment(id: number, fileName: string): Promise<void> {
        await sp.web.lists.getByTitle(TA_LIST).items.getById(id).attachmentFiles.getByName(fileName).delete();
    }

    private convertToIso(dateStr: string): string {
        if (!dateStr) return '';
        const parts = dateStr.split('.');
        if (parts.length === 3) {
            return `${parts[2]}-${parts[1]}-${parts[0]}`;
        }
        return dateStr;
    }

    public async setTermin(id: number, termin: string, verantwortlicherId?: number, delayReason?: string): Promise<void> {
        const updateData: any = {
            field_6: this.convertToIso(termin),
            Status: 'läuft planmäßig'
        };
        if (verantwortlicherId) {
            updateData.VerantwortlicherId = verantwortlicherId;
        }
        if (delayReason) {
            updateData.field_21 = delayReason;
        }
        await this.updateTA(id, updateData);
    }

    public async verschiebeTermin(id: number, neuerTermin: string, grund: string, alterTermin: string, urspruenglicherTermin?: string): Promise<void> {
        const updateData: Record<string, string> = {
            field_6: this.convertToIso(neuerTermin),
            field_21: grund
        };
        // Ursprünglichen Termin nur bei der ersten Verschiebung setzen
        if (!urspruenglicherTermin) {
            updateData.field_22 = this.convertToIso(alterTermin);
        }
        await this.updateTA(id, updateData);
    }

    public async getKpiData(tas: ITaItem[]): Promise<IKpiData> {
        return {
            overdue: tas.filter(t => t.Status === 'überfällig').length,
            plan: tas.filter(t => t.Status === 'Termin planen').length,
            onTrack: tas.filter(t => t.Status === 'läuft planmäßig').length,
            review: tas.filter(t => t.Status === 'prüfen').length
        };
    }

    /** Auto-detect TA statuses based on planned date and update them if needed */
    public async evaluateStatuses(tas: ITaItem[]): Promise<number> {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const updates: { id: number; status: string }[] = [];

        const pruefenTageRaw = await this.getConfigValue('PruefenTage', '3');
        const pruefenTage = parseInt(pruefenTageRaw, 10) || 3;

        for (const ta of tas) {
            if (ta.Status === 'abgeschlossen') continue;

            let expectedStatus = ta.Status;

            if (!ta.field_6) {
                expectedStatus = 'Termin planen';
            } else {
                let planDate: Date;
                if (ta.field_6.includes('-')) {
                    planDate = new Date(ta.field_6);
                } else {
                    const parts = ta.field_6.split('.');
                    if (parts.length === 3) {
                        planDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
                    } else {
                        planDate = new Date();
                    }
                }
                planDate.setHours(0, 0, 0, 0);

                // Prüfen-Schwelle aus TA_Config (PruefenTage) — Werktage
                const pruefenDeadline = this.addWorkdays(today, pruefenTage);

                if (planDate <= today) {
                    expectedStatus = 'überfällig';
                } else if (planDate <= pruefenDeadline) {
                    expectedStatus = 'prüfen';
                } else {
                    expectedStatus = 'läuft planmäßig';
                }
            }

            if (ta.Status !== expectedStatus) {
                ta.Status = expectedStatus;
                updates.push({ id: ta.ID, status: expectedStatus });
            }
        }

        // Batch status updates sequentially to avoid flooding SharePoint with parallel requests
        // that each re-trigger Power Automate flows
        for (const u of updates) {
            try {
                await this.updateTA(u.id, { Status: u.status });
            } catch (e) {
                console.error("Auto-status update failed for TA", u.id, e);
            }
        }

        return updates.length;
    }

    /**
     * Add N workdays (Mon–Fri) to a given date.
     * Replicates Excel's WORKDAY() function.
     */
    private addWorkdays(start: Date, days: number): Date {
        const result = new Date(start);
        let added = 0;
        while (added < days) {
            result.setDate(result.getDate() + 1);
            const dow = result.getDay();
            if (dow !== 0 && dow !== 6) { // Skip Sat (6) and Sun (0)
                added++;
            }
        }
        return result;
    }

    public async getNextTaNumber(): Promise<string> {
        const now = new Date();
        const yy = pad2(now.getFullYear() % 100);
        const mm = pad2(now.getMonth() + 1);
        const prefix = `TA${yy}${mm}`;

        // Alle TAs des aktuellen Monats abfragen
        const items = await sp.web.lists.getByTitle(TA_LIST).items
            .select('Title')
            .filter(`startswith(Title,'${prefix}')`)
            .top(5000)
            .get();

        let maxSeq = 0;
        for (const item of items) {
            const title: string = item.Title || '';
            // Letzten 2 Zeichen = laufende Nummer
            const seqStr = title.substring(prefix.length);
            const seq = parseInt(seqStr, 10);
            if (!isNaN(seq) && seq > maxSeq) {
                maxSeq = seq;
            }
        }

        return prefix + pad2(maxSeq + 1);
    }

    // ─── Projekt-Liste ─────────────────────────────────

    public async getProjekte(): Promise<IProjektItem[]> {
        try {
            // Lade die CSV Datei direkt aus der Dokumentenbibliothek "Projektliste"
            // Dies ist deutlich schneller als eine SharePoint Liste mit 1700 Elementen abzufragen
            const fileUrl = '/sites/TechnischeAnfragen/Projektliste/Projektliste.csv';

            // PnPjs: Dateiinhalt als Text (Blob/Text) abrufen
            const blob: Blob = await sp.web.getFileByServerRelativePath(fileUrl).getBlob();
            const text = await blob.text();

            // Zeilen aufteilen (unterstützt CRLF Windows und LF Unix)
            const lines = text.split(/\r?\n/);
            const projekte: IProjektItem[] = [];

            // Hilfsfunktion zum sauberen Trennen von CSV-Spalten (berücksichtigt Anführungszeichen)
            const splitCsvRow = (row: string, delimiter: string = ';'): string[] => {
                const result: string[] = [];
                let insideQuotes = false;
                let currentValue = '';

                for (let i = 0; i < row.length; i++) {
                    const char = row[i];
                    if (char === '"') {
                        insideQuotes = !insideQuotes;
                    } else if (char === delimiter && !insideQuotes) {
                        result.push(currentValue);
                        currentValue = '';
                    } else {
                        currentValue += char;
                    }
                }
                result.push(currentValue);
                return result;
            };

            // Zeile 0 ist der Header. Wir parsen diese, um die Spalten dynamisch zu finden.
            const headers = splitCsvRow(lines[0], ';').map(h => h.replace(/"/g, '').trim().toLowerCase());

            const getColIdx = (keywords: string[], defaultIdx: number) => {
                const found = headers.findIndex(h => keywords.some(k => h.includes(k)));
                return found !== -1 ? found : defaultIdx;
            };

            const idxKunde = getColIdx(['kunde'], 1);
            const idxName = getColIdx(['name'], 2);
            // Default 4 corresponds to VBA Column 5 (E) which was VproNumber usually
            const idxVpro = getColIdx(['vpro', 'projektnummer', 'anfrage-nr', 'kontakt-nr', 'kontaktnr'], 4);
            const idxSelektion = getColIdx(['selektion'], 6);
            // Fallback for Bezeichnung if Selektion not found? Actually they use both.
            // Let's ensure idxSelektion is used for the dropdown base!
            const idxKontakt = getColIdx(['kontakt'], 0);
            const idxPotential = getColIdx(['potential', 'potenzial'], 8);
            const idxChance = getColIdx(['chance'], 9);
            const idxBudget = getColIdx(['budget'], 10);
            const idxEndkunde = getColIdx(['endkunde'], 13);
            const idxMaterial = getColIdx(['material'], 16);
            const idxZielVk = getColIdx(['ziel vk', 'zielpreis'], 17);
            const idxSop = getColIdx(['sop', 'beginndat'], 20);
            const idxSegCode = getColIdx(['branch', 'branchcode', 'segcode', 'segment', 'seg'], 21); // Prioritize Branch names based on recent data inspection
            const idxVps = getColIdx(['vps'], 22);

            for (let i = 1; i < lines.length; i++) {
                const line = lines[i].trim();
                if (!line) continue;

                const cols = splitCsvRow(line, ';');
                if (cols.length >= 8) {

                    // Helper to safely parse German numbers (e.g. "5.000,50" -> 5000.50)
                    const parseNum = (val: string) => {
                        if (!val) return undefined;
                        // Remove quotes, remove thousand-separator dots, replace comma with dot
                        const clean = val.replace(/"/g, '').replace(/\./g, '').replace(',', '.').trim();
                        if (clean.length === 0) return undefined;
                        const parsed = parseFloat(clean);
                        return isNaN(parsed) ? undefined : parsed;
                    };

                    projekte.push({
                        ID: parseInt(cols[0].replace(/"/g, '').trim()) || i,
                        Title: cols[idxVpro] ? cols[idxVpro].replace(/"/g, '').trim() : '', // VproNumber
                        kontaktNr: cols[idxKontakt] ? cols[idxKontakt].replace(/"/g, '').trim() : '',
                        field_1: cols[idxKunde] ? cols[idxKunde].replace(/"/g, '').trim() : '',
                        field_2: cols[idxName] ? cols[idxName].replace(/"/g, '').trim() : '',
                        field_7: cols[idxSelektion] ? cols[idxSelektion].replace(/"/g, '').trim() : '', // Selektion is our dropdown value
                        field_8: parseNum(cols[idxPotential]) || 0,
                        field_10: parseNum(cols[idxChance]),
                        field_11: parseNum(cols[idxBudget]),
                        field_13: cols[idxEndkunde] ? cols[idxEndkunde].replace(/"/g, '').trim() : '',
                        field_16: cols[idxMaterial] ? cols[idxMaterial].replace(/"/g, '').trim() : undefined,
                        field_18: parseNum(cols[idxZielVk]),
                        field_21: cols[idxSop] ? cols[idxSop].replace(/"/g, '').trim() : undefined,
                        field_12: cols[idxSegCode] ? cols[idxSegCode].replace(/"/g, '').trim() : undefined,
                        vps: cols[idxVps] ? cols[idxVps].replace(/"/g, '').trim() : undefined
                    });
                }
            }

            // Optional: Sortieren nach Kundenname
            projekte.sort((a, b) => (a.field_1 || '').localeCompare(b.field_1 || ''));
            return projekte;

        } catch (error) {
            console.error("Fehler beim Laden der Projektliste.csv (Bitte sicherstellen, dass die Datei existiert):", error);
            return []; // Fallback, damit die App nicht crasht
        }
    }

    public async getDistinctKunden(projekte: IProjektItem[]): Promise<string[]> {
        const kundenSet = new Set<string>();
        projekte.forEach(p => {
            if (p.field_1) kundenSet.add(p.field_1);
        });
        return Array.from(kundenSet).sort();
    }

    // ─── Hilfstabellen (Kategorien & Kunden-Anwendungen) ─────────

    public async getKategorien(): Promise<IKategorieItem[]> {
        return sp.web.lists.getByTitle(KATEGORIEN_LIST).items
            .select('ID', 'Title', 'Email')
            .top(500)
            .orderBy('Title', true)
            .get();
    }

    public async addKategorie(title: string, email?: string): Promise<void> {
        const item: Record<string, string> = { Title: title };
        if (email) item.Email = email;
        await sp.web.lists.getByTitle(KATEGORIEN_LIST).items.add(item);
    }

    public async updateKategorie(id: number, fields: { Email?: string }): Promise<void> {
        await sp.web.lists.getByTitle(KATEGORIEN_LIST).items.getById(id).update(fields);
    }

    public async deleteKategorie(id: number): Promise<void> {
        await sp.web.lists.getByTitle(KATEGORIEN_LIST).items.getById(id).delete();
    }

    public async getKundenAnwendungen(kunde?: string): Promise<IKundenAnwendungItem[]> {
        let query = sp.web.lists.getByTitle(KUNDEN_ANWENDUNGEN_LIST).items
            .select('ID', 'Title', 'Anwendung')
            .top(2000)
            .orderBy('Title', true);

        if (kunde) {
            query = query.filter(`Title eq '${kunde.replace(/'/g, "''")}'`);
        }

        return query.get();
    }

    public async addKundenAnwendung(kunde: string, anwendung: string): Promise<void> {
        await sp.web.lists.getByTitle(KUNDEN_ANWENDUNGEN_LIST).items.add({
            Title: kunde,
            Anwendung: anwendung
        });
    }

    public async deleteKundenAnwendung(id: number): Promise<void> {
        await sp.web.lists.getByTitle(KUNDEN_ANWENDUNGEN_LIST).items.getById(id).delete();
    }

    // ─── Config ─────────────────────────────────────────

    public async getConfigValue(key: string, defaultValue: string = ''): Promise<string> {
        try {
            const items = await sp.web.lists.getByTitle(CONFIG_LIST).items
                .filter(`Title eq '${key}'`)
                .select('Value')
                .top(1)
                .get();

            if (items.length > 0) {
                return items[0].Value;
            }
        } catch (e) {
            console.warn(`Could not read config key ${key}. Using default.`, e);
        }
        return defaultValue;
    }

    public async setConfigValue(key: string, value: string): Promise<void> {
        const list = sp.web.lists.getByTitle(CONFIG_LIST);
        const existing = await list.items.filter(`Title eq '${key}'`).select('ID').top(1).get();

        if (existing.length > 0) {
            await list.items.getById(existing[0].ID).update({ Value: value });
        } else {
            await list.items.add({ Title: key, Value: value });
        }
    }
}
