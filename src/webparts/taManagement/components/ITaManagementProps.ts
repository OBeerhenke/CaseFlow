import { AppView } from '../models/types';

export interface ITaManagementProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userLoginName: string;
}

export interface IAppState {
  currentView: AppView;
  selectedTaId: number | undefined;
  filterStatus: string;
}
