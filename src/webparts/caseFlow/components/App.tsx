import * as React from 'react';
import styles from './App.module.scss';
import { ICaseFlowProps } from './ICaseFlowProps';
import { ICaseItem, IProjektItem, IKpiData, INewCaseForm, AppView } from '../models/types';
import { SharePointService } from '../services/SharePointService';
import { ConfigService } from '../services/ConfigService';
import { LabelService } from '../services/LabelService';
import { ThemeService } from '../services/ThemeService';
import NavBar from './NavBar';
import Dashboard from './Dashboard';
import NewCase from './NewCase';
import Schedule from './Schedule';
import CaseDetail from './CaseDetail';
import CaseList from './CaseList';
import Analytics from './Analytics';
import Settings from './Settings';

interface IAppState {
  currentView: AppView;
  selectedCaseId: number | undefined;
  previousView: AppView | undefined;
  previousFilterStatus: string;
  filterStatus: string;
  cases: ICaseItem[];
  projekte: IProjektItem[];
  users: any[];
  kpi: IKpiData;
  loading: boolean;
  toast: string;
  currentUserNickname: string;
  config: Record<string, string>;
}

export default class App extends React.Component<ICaseFlowProps, IAppState> {
  private svc: SharePointService;

  constructor(props: ICaseFlowProps) {
    super(props);
    this.svc = SharePointService.instance;
    this.state = {
      currentView: AppView.Dashboard,
      selectedCaseId: undefined,
      previousView: undefined,
      previousFilterStatus: '',
      filterStatus: '',
      cases: [],
      projekte: [],
      users: [],
      kpi: { overdue: 0, plan: 0, onTrack: 0, review: 0 },
      loading: true,
      toast: '',
      currentUserNickname: '',
      config: {}
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData();
    try {
      const nn = await this.svc.getNickname(this.props.userLoginName);
      if (nn) this.setState({ currentUserNickname: nn });
    } catch (e) {
      console.warn('Failed to load current user nickname', e);
    }
  }

  private async loadData(): Promise<void> {
    this.setState({ loading: true });
    try {
      const [cases, projekte, users, config] = await Promise.all([
        this.svc.getAllCases(),
        this.svc.getProjekte(),
        this.svc.getSiteUsers(),
        ConfigService.instance.getAll()
      ]);

      // Auto-detect overdue
      await this.svc.evaluateStatuses(cases);

      // Enrich cases with SharePoint user nicknames (short codes)
      await this.svc.enrichWithNicknames(cases);

      const kpi = await this.svc.getKpiData(cases);
      this.setState({ cases, projekte, users, kpi, config, loading: false });
    } catch (err) {
      console.error('Error loading data:', err);
      this.setState({ loading: false });
      this.showToast('Fehler beim Laden der Daten');
    }
  }

  private showToast = (message: string): void => {
    this.setState({ toast: message });
    setTimeout(() => this.setState({ toast: '' }), 3000);
  }

  private navigate = (view: AppView, filterStatus?: string): void => {
    this.setState({
      currentView: view,
      filterStatus: filterStatus || '',
      selectedCaseId: undefined
    });
  }

  private selectCase = (id: number): void => {
    this.setState({
      previousView: this.state.currentView,
      previousFilterStatus: this.state.filterStatus,
      currentView: AppView.CaseDetail,
      selectedCaseId: id
    });
  }

  private goBack = (): void => {
    if (this.state.currentView === AppView.CaseDetail && this.state.previousView) {
      this.setState({
        currentView: this.state.previousView,
        selectedCaseId: undefined,
        filterStatus: this.state.previousFilterStatus,
        previousView: undefined,
        previousFilterStatus: ''
      });
      return;
    }

    this.setState({
      currentView: AppView.Dashboard,
      selectedCaseId: undefined,
      filterStatus: '',
      previousView: undefined,
      previousFilterStatus: ''
    });
  }

  private handleCreateCase = async (form: INewCaseForm, files: File[]): Promise<void> => {
    try {
      const caseNr = await this.svc.getNextCaseNumber();
      const newId = await this.svc.createCase(form, caseNr);
      // Upload attachments if any
      for (const file of files) {
        const buffer = await file.arrayBuffer();
        await this.svc.addAttachment(newId, file.name, buffer);
      }
      this.showToast(`${caseNr} erfolgreich angelegt!`);
      await this.loadData();
      this.setState({ currentView: AppView.Dashboard });
    } catch (err) {
      console.error('Error creating case:', err);
      this.showToast('Fehler beim Anlegen');
    }
  }

  private handleSetTermin = async (id: number, termin: string, verantwortlicherId?: number, delayReason?: string): Promise<void> => {
    try {
      await this.svc.setTermin(id, termin, verantwortlicherId, delayReason);
      this.showToast('Termin erfolgreich geplant');
      await this.loadData();
      if (this.state.selectedCaseId === id) {
        this.setState({ currentView: AppView.Dashboard });
      }
    } catch (err) {
      console.error('Error planning termin:', err);
      this.showToast('Fehler beim Planen des Termins');
    }
  }

  private handleSaveCase = async (id: number, fields: Partial<Record<string, string | number | undefined>>): Promise<void> => {
    try {
      await this.svc.updateCase(id, fields);
      this.showToast('Gespeichert');
      await this.loadData();
      this.setState({ currentView: AppView.Dashboard });
    } catch (err) {
      console.error('Error saving case:', err);
      this.showToast('Fehler beim Speichern');
    }
  }

  private handleVerschieben = async (id: number, neuerTermin: string, grund: string, alterTermin: string, urspruenglicherTermin?: string): Promise<void> => {
    try {
      await this.svc.verschiebeTermin(id, neuerTermin, grund, alterTermin, urspruenglicherTermin);
      this.showToast('Termin verschoben');
      await this.loadData();
      this.setState({ currentView: AppView.Dashboard });
    } catch (err) {
      console.error('Error shifting termin:', err);
      this.showToast('Fehler beim Verschieben');
    }
  }

  private renderView(): React.ReactNode {
    const { currentView, cases, projekte, kpi, selectedCaseId, filterStatus } = this.state;

    switch (currentView) {
      case AppView.Dashboard:
        return (
          <Dashboard
            cases={cases}
            kpi={kpi}
            config={this.state.config}
            userName={this.props.userDisplayName}
            onNavigate={this.navigate}
            onSelectCase={this.selectCase}
          />
        );

      case AppView.NewCase:
        return (
          <NewCase
            projekte={projekte}
            users={this.state.users}
            onSubmit={this.handleCreateCase}
            onBack={this.goBack}
            onSearchUsers={async (q) => await this.svc.searchUsers(q)}
            onEnsureUser={async (login) => await this.svc.ensureUser(login)}
          />
        );

      case AppView.Schedule:
        return (
          <Schedule
            cases={cases}
            onSelectCase={this.selectCase}
            onBack={this.goBack}
          />
        );

      case AppView.CaseDetail: {
        const caseItem = cases.find(t => t.ID === selectedCaseId);
        if (!caseItem) {
          return (
            <div className={styles.content}>
              <div className={styles.emptyState}>
                <span>{LabelService.getEntityLabelSingular(this.state.config)} nicht gefunden</span>
                <button className={styles.btnPrimary} onClick={this.goBack}>Zurück</button>
              </div>
            </div>
          );
        }
        return (
          <CaseDetail
            caseItem={caseItem}
            onBack={this.goBack}
            onSave={this.handleSaveCase}
            onVerschieben={this.handleVerschieben}
            onSetTermin={this.handleSetTermin}
            onSearchUsers={async (q) => await this.svc.searchUsers(q)}
            onEnsureUser={async (login) => await this.svc.ensureUser(login)}
            currentUserDisplayName={this.props.userDisplayName}
            currentUserLoginName={this.props.userLoginName}
            currentUserNickname={this.state.currentUserNickname}
          />
        );
      }

      case AppView.CaseList:
        return (
          <CaseList
            cases={cases}
            initialFilter={filterStatus}
            onSelectCase={this.selectCase}
            onBack={this.goBack}
          />
        );

      case AppView.Analytics:
        return (
          <Analytics
            cases={cases}
            onSelectCase={this.selectCase}
            onBack={this.goBack}
          />
        );

      case AppView.Settings: {
        const kundenList = Array.from(new Set(projekte.map(p => p.field_2 || p.field_1).filter(Boolean))) as string[];
        return (
          <Settings
            onBack={this.goBack}
            kundenList={kundenList}
          />
        );
      }

      default:
        return undefined;
    }
  }

  public render(): React.ReactElement<ICaseFlowProps> {
    const { loading, currentView, toast } = this.state;

    return (
      <div className={styles.caseFlowApp} style={ThemeService.buildStyle(this.state.config)}>
        {loading ? (
          <div className={styles.loading}>
            <div className={styles.spinner} />
            Daten werden geladen...
          </div>
        ) : (
          <>
            <NavBar activeView={currentView} onNavigate={this.navigate} config={this.state.config} />
            {this.renderView()}

            {/* Mobile FAB: create a new case */}
            {currentView !== AppView.NewCase && (
              <button
                className={styles.fab}
                onClick={() => this.navigate(AppView.NewCase)}
                title={LabelService.getCreateActionLabel(this.state.config)}
              >
                +
              </button>
            )}
          </>
        )}

        {/* Toast notification */}
        {toast && <div className={styles.toastSuccess}>{toast}</div>}
      </div>
    );
  }
}
