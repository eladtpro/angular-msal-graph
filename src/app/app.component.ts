import { GraphService } from './services/graph.service';
import { AuthenticationService } from './services/authentication.service';
import { Component, OnInit } from '@angular/core';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  title = 'Syndi';
  // mail = 'hakam@spectory.com';
  mail = 'eladt.pro@gmail.com';;
  isIframe = false;
  loggedIn = false;

  constructor(private auth: AuthenticationService, private graph: GraphService) { }

  ngOnInit() {
    this.isIframe = window !== window.parent && !window.opener;
    this.checkAccount();
    this.auth.logged().subscribe(account => this.checkAccount());
  }

  checkAccount() {
    this.loggedIn = this.auth.authenticated();
  }

  login = () => this.auth.login();

  logout() {
    this.auth.logout();
  }

  invite() {
    this.graph.sendInvitation(this.mail);
  }

  calendar() {
    this.graph.getCalendarView('2020-11-01T19:00:00-08:00', '2020-11-02T19:00:00-08:00', '');
  }
}
