import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'ImecRedirectUsersApplicationCustomizerStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

const LOG_SOURCE: string = 'ImecRedirectUsersApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IImecRedirectUsersApplicationCustomizerProperties {
  // This is an example; replace with your own property
  groupWhitelist: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ImecRedirectUsersApplicationCustomizer
  extends BaseApplicationCustomizer<IImecRedirectUsersApplicationCustomizerProperties> {

  public CheckUserPermissions() {
    var absoluteUri = this.context.pageContext.web.absoluteUrl;

    this.context.spHttpClient.get(absoluteUri + "/_api/Web/CurrentUser?$select=ID",
      SPHttpClient.configurations.v1)
      .then((userResponse: SPHttpClientResponse) => {
        userResponse.json().then((user: any) => {
          var userId = user.Id;
          console.log(userId);
          this.context.spHttpClient.get(absoluteUri + "/_api/Web/GetUserById(" + userId + ")/Groups",
            SPHttpClient.configurations.v1)
            .then((groupResponse: SPHttpClientResponse) => {
              groupResponse.json().then((groupsData: any) => {
                var groups = groupsData.value;
                console.log(groups);

                groups.forEach(group => {
                  // Check if group contains 'Owners' no redirect will be done.
                  if (group.Title.indexOf(this.properties.groupWhitelist) === -1) {
                    // Redirect
                    if(window.location.href.indexOf('_layouts') !== -1){
                      window.location.href = this.context.pageContext.web.absoluteUrl;
                    }
                  }
                });
              });
            });
        });
      });
  }

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.CheckUserPermissions();

    return Promise.resolve();
  }
}
