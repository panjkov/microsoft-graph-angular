import 'rxjs/add/operator/catch';
import 'rxjs/add/operator/map';
import { Injectable } from '@angular/core';
import { Http } from '@angular/http';
import { Observable } from 'rxjs/Observable';

import { extractData, handleError } from '../shared/http-helper';
import { HttpService } from '../shared/http.service';
import { Event } from './event.model';

@Injectable()
export class HomeService {
  url = 'https://graph.microsoft.com/v1.0';
  file = 'demo.xlsx';
  table = 'Table1';

  constructor(
    private http: Http,
    private httpService: HttpService) {
  }

  getEvents(): Observable<Event[]> {
    return this.http
      .get(`${this.url}/me/events?$select=subject,organizer`, this.httpService.getAuthRequestOptions())
      .map(extractData)
      .catch(handleError);
  }

  addEventToExcel(events: Event[]) {
    const calendarEvents = [];

    events.forEach(event => {
      calendarEvents.push([event.subject, event.organizer.emailAddress.address]);
    });

    const calendarEventRequestBody = {
      index: null,
      values: calendarEvents
    };


    const body = JSON.stringify(calendarEventRequestBody);

    return this.http
      .post(
        `${this.url}/me/drive/root:/${this.file}:/workbook/tables/${this.table}/rows/add`,
        body,
        this.httpService.getAuthRequestOptions()
      )
      .map(extractData)
      .catch(handleError);
  }

}
