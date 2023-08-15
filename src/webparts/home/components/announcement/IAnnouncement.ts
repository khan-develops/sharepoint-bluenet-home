export interface IAnnouncement {
    Id: number;
    Title: string;
    Description: string;
    Date: string;
    Url: string;
    imageSource: string;
    IsActive: boolean;
    ImageFile: string;
    DocumentLink: {
        Url: string;
    };
    ImageLink: {
        Url: string;
        DataUrl: string | ArrayBuffer;
        Description: string;
    };
}