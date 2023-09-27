import { IItems } from '@pnp/sp/items';

export interface IProjectMapProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  listItems: IItems,
  mapApiKey: string,
  mapDataListName: string,
  startLat: number,
  startLon: number
}
