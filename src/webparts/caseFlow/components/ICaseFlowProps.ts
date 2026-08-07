import { AppView } from '../models/types';

export interface ICaseFlowProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  userLoginName: string;
}

export interface IAppState {
  currentView: AppView;
  selectedCaseId: number | undefined;
  filterStatus: string;
}
