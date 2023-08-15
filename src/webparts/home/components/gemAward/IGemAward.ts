import { IFileInfo } from "@pnp/sp/files";
import { ImageFit } from "office-ui-fabric-react";

export interface IGemAward extends Partial<IFileInfo> {
  title: string
  imageSrc: string;
  Email: string;
  description?: string;
  PersonalUrl?: string;
  showDetailsOnHover: boolean;
  imageFit: ImageFit;
}