import '@pnp/sp/user-custom-actions';

declare module '@pnp/sp/user-custom-actions' {
  interface IUserCustomActionInfo {
    ClientSideComponentId: string;
    ClientSideComponentProperties: string;
  }
}
