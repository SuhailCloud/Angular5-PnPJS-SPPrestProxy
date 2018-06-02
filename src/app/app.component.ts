import { Injectable, isDevMode, OnInit } from '@angular/core';
import { Component } from '@angular/core';
import { sp } from '@pnp/sp';
import { loadPageContext } from 'sp-rest-proxy/dist/utils/env';


@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  title = 'Angular 5 - @PnP - SP Prest Proxy';

  siteUrl = '';
  accountName = '';
  displayName = '';
  email = '';
  webTitle = '';
  constructor() {
    if (isDevMode()) {
      this.siteUrl = 'http://localhost:8082';
      console.log('Running in Developer Environment. SharePoint Rest Endpoint:' + this.siteUrl);
    } else {
      this.siteUrl = (<any>window)._spPageContextInfo.siteAbsoluteUrl;
      console.log('Running in Production Environment. SharePoint Rest Endpoint:' + this.siteUrl);
    }

    sp.setup({
      sp: {

        baseUrl: this.siteUrl,
        headers: {
          'Accept': 'application/json;odata=verbose'
        }
      }
    }
    );


  }
  ngOnInit() {
    // Using @Pnp Library
    sp.profiles.myProperties
      .get()
      .then((response) => {
        this.accountName = response.AccountName;
        this.displayName = response.DisplayName;
        this.email = response.Email;
      }).catch((error) => {
        console.log(error);
      });

    // Using SP Proxy
    loadPageContext().then(async _ => {
      this.webTitle = _spPageContextInfo.webTitle;

    }).catch((error) => {
      console.log(error);
    });
  }
}
