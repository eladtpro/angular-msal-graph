import { User as GraphUser } from '@microsoft/microsoft-graph-types';

export class User {
    constructor(user: GraphUser) {
        this.displayName = user.displayName;
        // Prefer the mail property, but fall back to userPrincipalName
        this.mail = user.mail || user.userPrincipalName;
        // this.timeZone = user.mailboxSettings.timeZone;
        // Use default avatar
        this.avatar = '/assets/no-profile-photo.png';
    }

    displayName: string;
    mail: string;
    avatar: string;
    timeZone: string;
}