export class User {
  displayName: string;
  emailAddresses: [{
    address: string
  }];
}

export class Event {
  subject: string;
  organizer: { 
    emailAddress: { 
        address: string 
        } };
}