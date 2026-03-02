import * as React from 'react';
import styles from './TaManagement.module.scss';
import { ITaManagementProps } from './ITaManagementProps';
import { ITaItem, IProjektItem, IKpiData, INewTaForm, AppView } from '../models/types';
import { SharePointService } from '../services/SharePointService';
import NavBar from './NavBar';
import Dashboard from './Dashboard';
import NeueTa from './NeueTa';
import TerminPlanen from './TerminPlanen';
import TaDetail from './TaDetail';
import AlleTas from './AlleTas';

interface IAppState {
  currentView: AppView;
  selectedTaId: number | undefined;
  filterStatus: string;
  tas: ITaItem[];
  projekte: IProjektItem[];
  users: any[];
  kpi: IKpiData;
  loading: boolean;
  toast: string;
}

export default class TaManagement extends React.Component<ITaManagementProps, IAppState> {
  private svc: SharePointService;

  constructor(props: ITaManagementProps) {
    super(props);
    this.svc = SharePointService.instance;
    this.state = {
      currentView: AppView.Dashboard,
      selectedTaId: undefined,
      filterStatus: '',
      tas: [],
      projekte: [],
      users: [],
      kpi: { overdue: 0, plan: 0, onTrack: 0, review: 0 },
      loading: true,
      toast: ''
    };
  }

  public async componentDidMount(): Promise<void> {
    await this.loadData();
  }

  private async loadData(): Promise<void> {
    this.setState({ loading: true });
    try {
      const [tas, projekte, users] = await Promise.all([
        this.svc.getAllTAs(),
        this.svc.getProjekte(),
        this.svc.getSiteUsers()
      ]);

      // Auto-detect overdue
      await this.svc.checkOverdue(tas);

      const kpi = await this.svc.getKpiData(tas);
      this.setState({ tas, projekte, users, kpi, loading: false });
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
      selectedTaId: undefined
    });
  }

  private selectTa = (id: number): void => {
    this.setState({
      currentView: AppView.TaDetail,
      selectedTaId: id
    });
  }

  private goBack = (): void => {
    this.setState({
      currentView: AppView.Dashboard,
      selectedTaId: undefined,
      filterStatus: ''
    });
  }

  private handleCreateTA = async (form: INewTaForm): Promise<void> => {
    try {
      const taNr = await this.svc.getNextTaNumber();
      await this.svc.createTA(form, taNr);
      this.showToast(`${taNr} erfolgreich angelegt!`);
      await this.loadData();
      this.setState({ currentView: AppView.Dashboard });
    } catch (err) {
      console.error('Error creating TA:', err);
      this.showToast('Fehler beim Anlegen der TA');
    }
  }

  private handleSetTermin = async (id: number, termin: string, verantwortlicherId?: number): Promise<void> => {
    try {
      await this.svc.setTermin(id, termin, verantwortlicherId);
      this.showToast('Termin erfolgreich geplant');
      await this.loadData();
      if (this.state.selectedTaId === id) {
        this.setState({ currentView: AppView.Dashboard });
      }
    } catch (err) {
      console.error('Error planning termin:', err);
      this.showToast('Fehler beim Planen des Termins');
    }
  }

  private handleSaveTa = async (id: number, fields: Partial<Record<string, string | number | undefined>>): Promise<void> => {
    try {
      await this.svc.updateTA(id, fields);
      this.showToast('Gespeichert');
      await this.loadData();
      this.setState({ currentView: AppView.Dashboard });
    } catch (err) {
      console.error('Error saving TA:', err);
      this.showToast('Fehler beim Speichern');
    }
  }

  private handleVerschieben = async (id: number, neuerTermin: string, grund: string, alterTermin: string): Promise<void> => {
    try {
      await this.svc.verschiebeTermin(id, neuerTermin, grund, alterTermin);
      this.showToast('Termin verschoben');
      await this.loadData();
      // Refresh TA detail
      this.setState({ selectedTaId: id });
    } catch (err) {
      console.error('Error shifting termin:', err);
      this.showToast('Fehler beim Verschieben');
    }
  }

  private renderView(): React.ReactNode {
    const { currentView, tas, projekte, kpi, selectedTaId, filterStatus } = this.state;

    switch (currentView) {
      case AppView.Dashboard:
        return (
          <Dashboard
            tas={tas}
            kpi={kpi}
            userName={this.props.userDisplayName}
            onNavigate={this.navigate}
            onSelectTa={this.selectTa}
          />
        );

      case AppView.NeueTa:
        return (
          <NeueTa
            projekte={projekte}
            users={this.state.users}
            onSubmit={this.handleCreateTA}
            onBack={this.goBack}
            onSearchUsers={async (q) => await this.svc.searchUsers(q)}
            onEnsureUser={async (login) => await this.svc.ensureUser(login)}
          />
        );

      case AppView.TerminPlanen:
        return (
          <TerminPlanen
            tas={tas}
            onSetTermin={this.handleSetTermin}
            onBack={this.goBack}
            onSearchUsers={async (q) => await this.svc.searchUsers(q)}
            onEnsureUser={async (login) => await this.svc.ensureUser(login)}
          />
        );

      case AppView.TaDetail: {
        const ta = tas.find(t => t.ID === selectedTaId);
        if (!ta) {
          return (
            <div className={styles.content}>
              <div className={styles.emptyState}>
                <span>TA nicht gefunden</span>
                <button className={styles.btnPrimary} onClick={this.goBack}>Zurück</button>
              </div>
            </div>
          );
        }
        return (
          <TaDetail
            ta={ta}
            onBack={this.goBack}
            onSave={this.handleSaveTa}
            onVerschieben={this.handleVerschieben}
          />
        );
      }

      case AppView.AlleTas:
        return (
          <AlleTas
            tas={tas}
            initialFilter={filterStatus}
            onSelectTa={this.selectTa}
            onBack={this.goBack}
          />
        );

      default:
        return undefined;
    }
  }

  public render(): React.ReactElement<ITaManagementProps> {
    const { loading, currentView, toast } = this.state;

    return (
      <div className={styles.taApp}>
        {loading ? (
          <div className={styles.loading}>
            <div className={styles.spinner} />
            Daten werden geladen...
          </div>
        ) : (
          <>
            {this.renderView()}
            <NavBar activeView={currentView} onNavigate={this.navigate} />
          </>
        )}

        {/* Toast notification */}
        {toast && <div className={styles.toastSuccess}>{toast}</div>}
      </div>
    );
  }
}
