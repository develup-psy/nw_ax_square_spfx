export interface INwAxSquareProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  env?: string | null;
  isWebView?: boolean;
}
