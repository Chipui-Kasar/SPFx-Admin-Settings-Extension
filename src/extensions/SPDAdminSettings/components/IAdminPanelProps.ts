import { IUserCustomActionInfo } from "@pnp/sp/user-custom-actions/types";

export interface IAdminPanelProps {
  topString: string;
  context: any;
}
export interface IExtendedUserCustomActionInfo extends IUserCustomActionInfo {
  ClientSideComponentProperties?: string;
}

export interface ICustomCSSProperties {
  fullWidth: boolean;
  spacing: number;
  alignment: string;
}
