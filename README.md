Extract calendar availablity for Google Calendar, as a string you can copy into an email.

Supports multiple timezones and configurations.

Code is based on the [Google API Example](https://developers.google.com/calendar/api/quickstart/python), and only runs locally.

Must add `client_secret.json` to work.

Steps to set up access to a calendar:
1. Enable Google Calendar API (see steps [here](https://developers.google.com/identity/protocols/oauth2/native-app#enable-apis))
2. Create authorization credentials OAuth  2.0 Client ID for a "Desktop application":
   * Go to the [Credentials](https://console.developers.google.com/apis/credentials)
   * Select *Create credentials > OAuth client ID*.
   * Set the application type to *Desktop app*
3. Download JSON and save the credentials as client_secret.json
