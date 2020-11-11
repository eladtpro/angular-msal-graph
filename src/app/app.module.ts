import { BrowserModule } from '@angular/platform-browser';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { NgModule } from '@angular/core';

import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatButtonModule } from '@angular/material/button';
import { MatListModule } from '@angular/material/list';
import { MatToolbarModule } from '@angular/material/toolbar';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';

import { Logger, LogLevel, CryptoUtils } from 'msal';
import { MsalModule, MsalInterceptor } from '@azure/msal-angular';
import { HTTP_INTERCEPTORS, HttpClientModule } from '@angular/common/http';
import { ProfileComponent } from './components/profile/profile.component';
import { HomeComponent } from './components/home/home.component';
import { LoggedOutComponent } from './components/logged-out/logged-out.component';
import { InvitationComponent } from './components/invitation/invitation.component';
import { CalendarComponent } from './components/calendar/calendar.component';

const isIE = window.navigator.userAgent.indexOf('MSIE ') > -1 || window.navigator.userAgent.indexOf('Trident/') > -1;

@NgModule({
  declarations: [
    AppComponent,
    ProfileComponent,
    HomeComponent,
    LoggedOutComponent,
    InvitationComponent,
    CalendarComponent
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    BrowserAnimationsModule,
    HttpClientModule,
    MatToolbarModule,
    MatFormFieldModule,
    MatInputModule,
    MatButtonModule,
    MatListModule,
    AppRoutingModule,
    MsalModule.forRoot({
      auth: {
        clientId: '838d477f-42c9-4701-9490-eb6fef98bee4',
        authority: 'https://login.microsoftonline.com/03a0f57d-d561-42fe-a38b-cd46ed17d941/',
        redirectUri: 'http://localhost:4200/auth',
        postLogoutRedirectUri: "http://localhost:4200/goodbye",
        navigateToLoginRequestUrl: true
      },
      cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: isIE, // set to true for IE 11
      },
      system: {
        logger: new Logger((level: LogLevel, message: string, containsPii: boolean): void => {
          if (containsPii) {
            return;
          }
          switch (level) {
            case LogLevel.Error:
              console.error(message);
              return;
            case LogLevel.Info:
              console.info(message);
              return;
            case LogLevel.Verbose:
              console.debug(message);
              return;
            case LogLevel.Warning:
              console.warn(message);
              return;
          }
        }, {
          level: LogLevel.Info,
          piiLoggingEnabled: false,
          correlationId: `[MsalAngular(${CryptoUtils.createNewGuid()})]`
        })
      }
    },
      {
        popUp: !isIE,
        consentScopes: [
          // 'user.read',
          'offline_access',
          // 'email',
          'openid',
          'profile',
          'User.Invite.All',
          'User.ReadWrite.All', 
          'Directory.ReadWrite.All',
          // 'User.ReadWrite.All', 
          // 'Directory.ReadWrite.All',
          // 'Calendars.Read', 
          // 'Calendars.ReadWrite'
          // 'https://b2cpm.onmicrosoft.com/orca/read',
          // 'https://b2cpm.onmicrosoft.com/orca/write',
          // 'https://b2cpm.onmicrosoft.com/orca/user_impersonation'
        ],
        unprotectedResources: [],
        protectedResourceMap: [
          // [environment.graph, ['user.read']]
        ],
        extraQueryParameters: {}
      })
  ],
  providers: [
    {
      provide: HTTP_INTERCEPTORS,
      useClass: MsalInterceptor,
      multi: true
    }
  ],
  bootstrap: [AppComponent]
})
export class AppModule { }
