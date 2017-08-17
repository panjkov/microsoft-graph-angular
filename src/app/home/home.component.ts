import { Component, OnInit, OnDestroy } from '@angular/core';
import { Subscription } from 'rxjs/Subscription';

import { User } from './user.model';
import { Event } from './user.model';
import { HomeService } from './home.service';
import { AuthService } from '../auth/auth.service';

@Component({
  selector: 'app-home',
  template: `
    <table>
      <tr>
        <th align="left">Name</th>
        <th align="left">Email</th>
      </tr>
      <tr *ngFor="let event of events">
        <td>{{event.subject}}</td>
        <td>{{event.organizer.emailAddress.address}}</td>
      </tr>
    </table>
    <button (click)="onAddEventToExcel()">Write to Excel</button>
    <button (click)="onLogout()">Logout</button>
  `
})
export class HomeComponent implements OnInit, OnDestroy {
  users: User[];
  events: Event[];
  subsGetUsers: Subscription;
  subsGetEvents: Subscription;
  subsAddContactToExcel: Subscription;
  subsAddEventToExcel: Subscription;

  constructor(
    private homeService: HomeService,
    private authService: AuthService
  ) { }

  ngOnInit() {
    //this.subsGetUsers = this.homeService.getUsers().subscribe(users => this.users = users);
    this.subsGetEvents = this.homeService.getEvents().subscribe(events => this.events = events);
  }

  ngOnDestroy() {
    this.subsGetUsers.unsubscribe();
    this.subsAddContactToExcel.unsubscribe();
    this.subsAddEventToExcel.unsubscribe();
  }

  onAddContactToExcel() {
    this.subsAddContactToExcel = this.homeService.addContactToExcel(this.users).subscribe();
  }

  onAddEventToExcel() {
    this.subsAddEventToExcel = this.homeService.addEventToExcel(this.events).subscribe();
  }

  onLogout() {
    this.authService.logout();
  }
}
