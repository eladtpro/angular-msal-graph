import { User } from '@microsoft/microsoft-graph-types';
import { GraphService } from '../../services/graph.service';
import { Component, OnInit } from '@angular/core';

@Component({
  selector: 'app-profile',
  templateUrl: './profile.component.html',
  styleUrls: ['./profile.component.css']
})
export class ProfileComponent implements OnInit {
  profile: User;

  constructor(private graph: GraphService) { }

  ngOnInit() {
    this.graph.getProfile()
      .then(user => this.profile = user)
      .catch(reason => console.error(reason));
  }
}
