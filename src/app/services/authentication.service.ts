import { Injectable } from '@angular/core';
import { BroadcastService, MsalService } from '@azure/msal-angular';
import { Account, AuthenticationParameters } from 'msal'
import { BehaviorSubject, Observable } from 'rxjs';

const AuthParams: AuthenticationParameters = {
  redirectUri: 'http://localhost:4200/auth',
  scopes: [
    // "user.read",
    // "mailboxsettings.read",
    // "calendars.readwrite",
    'User.Invite.All',
    'User.ReadWrite.All', 
    'Directory.ReadWrite.All',
    // 'Calendars.Read', 
    // 'Calendars.ReadWrite'
  ]
};

@Injectable({
  providedIn: 'root'
})
export class AuthenticationService {

  private logged$: BehaviorSubject<Account> = new BehaviorSubject(null);

  logged(): Observable<Account> {
    return this.logged$.asObservable();
  }

  constructor(private msal: MsalService, private broadcastService: BroadcastService) {

    this.msal.handleRedirectCallback((authError, response) => {
      if (authError) {
        console.error('Redirect Error: ', authError.errorMessage);
        return;
      }
      console.log('Redirect Success: ', response.accessToken);
    });

    this.broadcastService.subscribe('msal:loginSuccess', () => {
      this.logged$.next(this.getAccount());
    });
  }

  private isIE = () => window.navigator.userAgent.indexOf('MSIE ') > -1 || window.navigator.userAgent.indexOf('Trident/') > -1;

  getAccount = (): Account => this.msal.getAccount();

  authenticated = (): boolean => !!this.getAccount();

  login = (): void => {
    if (this.isIE())
      this.msal.loginRedirect();
    else
      this.msal.loginPopup()
        .catch((reason) => {
          console.error('Login failed', JSON.stringify(reason, null, 2));
        });
  }

  logout = (): void => this.msal.logout();

  getAccessToken = (): Promise<string> =>
    new Promise((resolve, reject) => {
      this.msal.acquireTokenSilent(AuthParams)
        .then(result => resolve(result.accessToken))
        .catch(error => {
          //Acquire token silent failure, and send an interactive request
          console.error(error.errorMessage, error);
          if (error.errorMessage.indexOf("interaction_required") === -1)
            reject(error);

          if (this.isIE())
            this.msal.acquireTokenRedirect(AuthParams)
          else
            this.msal.acquireTokenPopup(AuthParams)
              .then(response => resolve(response.accessToken))
              .catch(reject);
        });
    });
}
