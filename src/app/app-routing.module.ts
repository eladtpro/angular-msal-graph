import { InvitationComponent } from './components/invitation/invitation.component';
import { LoggedOutComponent } from './components/logged-out/logged-out.component';
import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';
import { ProfileComponent } from './components/profile/profile.component';
import { MsalGuard } from '@azure/msal-angular';
import { HomeComponent } from './components/home/home.component';

const routes: Routes = [
  { path: 'profile', component: ProfileComponent, canActivate: [MsalGuard] },
  { path: 'auth', redirectTo: 'profile' },
  { path: 'invitation', component: InvitationComponent },
  { path: 'goodbye', component: LoggedOutComponent },
  { path: '', component: HomeComponent }
];

@NgModule({
  imports: [RouterModule.forRoot(routes, { useHash: false })],
  exports: [RouterModule]
})
export class AppRoutingModule { }
