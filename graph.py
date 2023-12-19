import msal
from datetime import date, timedelta
import requests

class Graph:
    def __init__(self, settings, token_cache):
        self.settings = settings
        self.token_cache = token_cache
        authority = "https://login.microsoftonline.com/{}".format(self.settings["tenant_id"])
        self.app = msal.PublicClientApplication(
            self.settings["client_id"], authority=authority,
            token_cache=token_cache
        )

    def get_token(self):
        result = None

        accounts = self.app.get_accounts(username=self.settings["username"])
        if accounts:
            chosen = accounts[0]  # Assuming the end user chose this one to proceed
            result = self.app.acquire_token_silent(self.settings["scopes"], account=chosen)

        if not result:
            result = self.app.acquire_token_interactive(
                self.settings["scopes"],
                login_hint=self.settings["username"],
            )

        if "access_token" in result:
            return {"token": result["access_token"], "source": result["token_source"]}
        else:
            print("Token acquisition failed: %s" % (result["error_description"]))
            return None

    def get_calendar_entries(self):
        # Calling graph using the access token
        start = date.today() - timedelta(days=self.settings["days_history"])
        end = date.today() + timedelta(days=self.settings["days_future"])
        graph_response = requests.get(
            f"https://graph.microsoft.com/v1.0/me/calendarview?startdatetime={start}T00:00:00.000Z&enddatetime={end}T23:59:59.999Z&$orderby=start/dateTime&$top:50",
            headers={
                'Authorization': 'Bearer ' + self.get_token()["token"],
                'Prefer': 'outlook.timezone="India Standard Time"',
                'Prefer': 'outlook.body-content-type="text"'
            },)
        if graph_response.status_code != 200:
            print(graph_response.text)
            return None

        return graph_response.json()
