import { IUserProfile } from "@pnp/sp/profiles";


export interface IProfile extends IUserProfile {
  Id: number;
  DisplayName: string;
  Email: string;
  Title: string;
  PersonalUrl: string;
  PictureUrl: string;
  WorkPhone: string;
  CellPhone: string;
  HireDate: string;
  BirthDate: string;
}
