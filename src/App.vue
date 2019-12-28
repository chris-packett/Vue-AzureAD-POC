<template>
  <div>
    <h3>Hello, {{ displayName }}</h3>
    <p>You are a {{ jobTitle }}. People can reach your mobile phone at {{ mobilePhone }}</p>
  </div>
</template>

<script>
import * as Msal from 'msal';
import { msalConfig, graphConfig } from './config/msalConfig';

const request = { scopes: ['user.read'] };

const msalUserAgentApplication = new Msal.UserAgentApplication(msalConfig);

msalUserAgentApplication.handleRedirectCallback((err, response) => {
  if (err) {
    console.log(err);
  } else if (response.tokenType === "access_token") {
    this.callMSGraph(graphConfig.graphMeEndpoint, response.accessToken, this.graphAPICallback);
  }
});

export default {
  name: 'app',
  data: function () {
    return {
      displayName: '',
      jobTitle: '',
      mobilePhone: ''
    };
  },
  mounted: function () {
    if (msalUserAgentApplication.getAccount()) {
      this.showWelcomeMessage();
  
      this.acquireTokenPopupAndCallMSGraph();
    } else {
      msalUserAgentApplication
        .loginRedirect(request)
        .then(() => {
          this.showWelcomeMessage();

          this.acquireTokenPopupAndCallMSGraph();
        })
        .catch(err => {
          console.log(err);
        });
    }
  },
  methods: {
    acquireTokenPopupAndCallMSGraph: function () {
      msalUserAgentApplication
        .acquireTokenSilent(request)
        .then(resp => {

          this.callMSGraph(graphConfig.graphMeEndpoint, resp.accessToken, this.graphAPICallback);

        }).catch(err => {

            console.log(err);

            if (this.requiresInteraction(err.errorCode)) {

              msalUserAgentApplication
                .acquireTokenPopup(request)
                .then(resp => {

                  this.callMSGraph(graphConfig.graphMeEndpoint, resp.accessToken, this.graphAPICallback);

                }).catch(err => {

                    console.log(err);

                });
            }
        });
    },
    callMSGraph: function (url, accessToken, callback) {
      var xmlHttp = new XMLHttpRequest();

      xmlHttp.onreadystatechange = function () {
        if (this.readyState == 4 && this.status == 200) {
          callback(JSON.parse(this.responseText));
        }
      }

      xmlHttp.open("GET", url, true);
      xmlHttp.setRequestHeader('Authorization', 'Bearer ' + accessToken);
      xmlHttp.send();
    },
    graphAPICallback: function (data) {
      this.displayName = data.displayName;
      this.jobTitle = data.jobTitle;
      this.mobilePhone = data.mobilePhone;
    },
    showWelcomeMessage: function () {
      console.log(`Welcome ${msalUserAgentApplication.getAccount().userName} to Microsoft Graph API!`);
    },
    requiresInteraction: function (errorCode) {
      if (!errorCode || !errorCode.length) return false;

      return errorCode === "consent_required" ||
          errorCode === "interaction_required" ||
          errorCode === "login_required";
    }
  }
}
</script>
