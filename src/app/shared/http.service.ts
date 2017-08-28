import { Injectable } from '@angular/core';
import { Headers } from '@angular/http';
import * as hello from 'hellojs/dist/hello.all.js';

@Injectable()
export class HttpService {
  getAuthRequestOptions() {
    const msft = hello('msft').getAuthResponse();
    const authHeaders = new Headers({ 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + msft.access_token });

    return { headers: authHeaders };
  }
  getAccessToken() {
    const msft = hello('msft').getAuthResponse();
    const accessToken = msft.access_token;
    return accessToken;
  }
}
