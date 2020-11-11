import { Injectable } from '@angular/core';
import { Client } from '@microsoft/microsoft-graph-client';
import { Event, User as GraphUser } from '@microsoft/microsoft-graph-types';
import { AuthenticationService } from './authentication.service';
import { User } from '../models/user';

@Injectable({
  providedIn: 'root'
})
export class GraphService {

  private client: Client;

  constructor(private auth: AuthenticationService) {
    this.client = Client.init({
      debugLogging: true,
      authProvider: async (done) => {
        this.auth.getAccessToken()
          .then(token => done(null, token))
          .catch(reason => done(reason, null))
      }
    });
  }

  async getProfile(): Promise<GraphUser> {
    if (!this.auth.authenticated()) return null;

    // Get the user from Graph (GET /me)
    let graphUser: GraphUser = await this.client
      .api('/me')
      // .select('displayName,mail,mailboxSettings,userPrincipalName')
      .select('displayName,mail')
      .get();

    let user = new User(graphUser);
    return user;
  }

  // https://docs.microsoft.com/en-us/graph/api/invitation-post?view=graph-rest-1.0&tabs=http
  async sendInvitation(mail: string) {
    await this.sendInvitationGraph(mail);




  }


  async sendInvitationGraph(mail: string) {
    const request = {
      "invitedUserEmailAddress": mail,
      "inviteRedirectUrl": "http://localhost:4200/invitation",
      // "invitedUserDisplayName": "Elad Tal",
      // "invitedUserMessageInfo": { "@odata.type": "microsoft.graph.invitedUserMessageInfo" },
      "sendInvitationMessage": true,
      "inviteRedeemUrl": "http://localhost:4200/auth",
      // "status": "string",
      // "invitedUser": { "@odata.type": "microsoft.graph.user" },
      // "invitedUserType": "string"
    }

    let invitation = await this.client.api('/invitations').post(request)
      .then(result => console.log(result))
      .catch(error => console.error(error));
  }

  // POST https://graph.microsoft.com/v1.0/invitations
  // Content-type: application/json
  // Content-length: 551

  // {
  //   "invitedUserEmailAddress": "yyy@test.com",
  //   "inviteRedirectUrl": "https://myapp.contoso.com"
  // }

  async getCalendarView(start: string, end: string, timeZone: string): Promise<Event[]> {
    if (!this.auth.authenticated()) return null;

    try {
      // GET /me/calendarview?startDateTime=''&endDateTime=''
      // &$select=subject,organizer,start,end
      // &$orderby=start/dateTime
      // &$top=50
      let result = await this.client
        .api('/me/calendarview')
        // .header('Prefer', `outlook.timezone="${timeZone}"`)
        .query({
          startDateTime: start,
          endDateTime: end
        })
        .select('subject,organizer,start,end')
        .orderby('start/dateTime')
        .top(50)
        .get();

      return result.value;
    } catch (error) {
      console.error('Could not get events', JSON.stringify(error, null, 2));
    }
  }

  // <AddEventSnippet>
  async addEventToCalendar(newEvent: Event): Promise<void> {
    try {
      // POST /me/events
      await this.client
        .api('/me/events')
        .post(newEvent);
    } catch (error) {
      throw Error(JSON.stringify(error, null, 2));
    }
  }


}
