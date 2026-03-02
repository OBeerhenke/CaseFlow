import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/files'; // Added import for getFileByServerRelativeUrl / getFileByServerRelativePath
import '@pnp/sp/site-users/web';
import '@pnp/sp/profiles';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ITaItem, IProjektItem, INewTaForm, IKpiData } from '../models/types';

const TA_LIST = 'Technische_Anfragen';
const PROJEKT_LIST = 'Projekt-Liste';

function pad2(n: number): string {
    return n < 10 ? '0' + n : '' + n;
}
function pad3(n: number): string {
    if (n < 10) return '00' + n;
    if (n < 100) return '0' + n;
    return '' + n;
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
                'field_20', 'field_21', 'field_22', 'Status', 'Modified',
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
            return results;
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
                'field_20', 'field_21', 'field_22', 'Status', 'Modified',
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
                'field_20', 'field_21', 'field_22', 'Status', 'Modified',
                'Ersteller/Title', 'Ersteller/EMail',
                'Verantwortlicher/Title', 'Verantwortlicher/EMail'
            )
            .expand('Ersteller', 'Verantwortlicher')
            .get();
    }

    public async createTA(form: INewTaForm, taNr: string): Promise<void> {
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
            field_2: form.aufgabe,
            field_11: form.anwendung,
            field_12: form.material,
            field_16: form.kategorie,
            field_14: !isNaN(potVal as any) ? potVal : null,
            field_15: !isNaN(chanceVal as any) ? chanceVal : null,
            ErstellerId: currentUser.Id,
            VerantwortlicherId: form.verantwortlicherId || null,
            field_17: form.budget,
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
            await sp.web.lists.getByTitle(TA_LIST).items.add(item);
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

    private convertToIso(dateStr: string): string {
        if (!dateStr) return '';
        const parts = dateStr.split('.');
        if (parts.length === 3) {
            return `${parts[2]}-${parts[1]}-${parts[0]}`;
        }
        return dateStr;
    }

    public async setTermin(id: number, termin: string, verantwortlicherId?: number): Promise<void> {
        const updateData: any = {
            field_6: this.convertToIso(termin),
            Status: 'läuft planmäßig'
        };
        if (verantwortlicherId) {
            updateData.VerantwortlicherId = verantwortlicherId;
        }
        await this.updateTA(id, updateData);
    }

    public async verschiebeTermin(id: number, neuerTermin: string, grund: string, alterTermin: string): Promise<void> {
        await this.updateTA(id, {
            field_22: this.convertToIso(alterTermin),
            field_6: this.convertToIso(neuerTermin),
            field_21: grund
        });
    }

    public async getKpiData(tas: ITaItem[]): Promise<IKpiData> {
        return {
            overdue: tas.filter(t => t.Status === 'überfällig').length,
            plan: tas.filter(t => t.Status === 'Termin planen').length,
            onTrack: tas.filter(t => t.Status === 'läuft planmäßig').length,
            review: tas.filter(t => t.Status === 'prüfen').length
        };
    }

    /** Auto-detect overdue TAs and update their status */
    public async checkOverdue(tas: ITaItem[]): Promise<number> {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        let count = 0;

        for (const ta of tas) {
            if (ta.Status !== 'überfällig' && ta.Status !== 'abgeschlossen' && ta.field_6) {
                const parts = ta.field_6.split('.');
                if (parts.length === 3) {
                    const planDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
                    if (planDate < today) {
                        await this.updateTA(ta.ID, { Status: 'überfällig' });
                        ta.Status = 'überfällig';
                        count++;
                    }
                }
            }
        }
        return count;
    }

    public async getNextTaNumber(): Promise<string> {
        const items = await sp.web.lists.getByTitle(TA_LIST).items
            .select('ID')
            .top(1)
            .orderBy('ID', false)
            .get();

        const nextNum = items.length > 0 ? items[0].ID + 1 : 1;
        const year = new Date().getFullYear();
        return 'TA-' + year + '-' + pad3(nextNum);
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

            // Zeile 0 ist der Header ("Kontakt";"Kunde";...). Wir starten bei i = 1.
            for (let i = 1; i < lines.length; i++) {
                const line = lines[i].trim();
                if (!line) continue;

                const cols = splitCsvRow(line, ';');
                if (cols.length >= 8) {
                    projekte.push({
                        ID: parseInt(cols[0].replace(/"/g, '').trim()) || i, // "Kontakt" als ID
                        Title: cols[0].replace(/"/g, '').trim(),      // "Kontakt" 
                        field_1: cols[1].replace(/"/g, '').trim(),    // "Kunde"
                        field_2: cols[2].replace(/"/g, '').trim(),    // "Name"
                        field_3: cols[3].replace(/"/g, '').trim(),    // "Name 2"
                        field_7: cols[7].replace(/"/g, '').trim(),    // "Bezeichnung"
                        field_8: parseFloat(cols[8].replace(/"/g, '').replace(',', '.')) || 0, // "Potential"
                        field_13: cols[13] ? cols[13].replace(/"/g, '').trim() : '' // "Endkunde"
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
}
