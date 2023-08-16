import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IHomeProps {
	context: WebPartContext
}

export interface ISiteUser {
	Id: number;
	Title: string;
	JobTitle: string;
	EMail: string;
	WorkPhone: string;
	MobilePhone: string;
	Office: string;
	Name: string;
	UserUrl: string;
	HireDate: string;
	BirthDate: string;
}

