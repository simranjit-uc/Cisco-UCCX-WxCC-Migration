# Cisco UCCX to Webex Contact Center (WxCC) Configuration Migration Tool

This tool automates the migration of **Cisco Unified Contact Center Express (UCCX)** configurations to **Webex Contact Center (WxCC)**. It extracts UCCX data, transforms it into WxCC compatible format, and pushes the configuration into a WxCC tenant using Cisco's REST APIs. The link to the video demo of the app is given at the bottom.


## Project Overview

### Migration Flow:

1. **Extract**: Pulls configuration from UCCX using REST APIs and stores it in an Excel file.
2. **Transform**: Converts UCCX configuration to WxCC-compatible format and saves it in a new Excel file.
3. **Authenticate**: Performs OAuth 2.0 authentication with Cisco Webex APIs.
4. **Push**: Uploads configuration to WxCC tenant via REST APIs.


### File Structure

| File Name          | Purpose |
|--------------------|---------|
| `main.py`          | Entry point. Coordinates the overall workflow and calls other modules. |
| `CCX_Sheet.py`     | Connects to UCCX, extracts configuration using REST APIs, and saves it to `CCX-Details.xlsx`. |
| `WxCC_Sheet.py`    | Reads `CCX-Details.xlsx`, transforms data into WxCC format, and writes to `WxCC Details.xlsx`. |
| `WxCC.py`          | Pushes data from `WxCC Details.xlsx` to WxCC tenant via REST APIs. It also triggers OAuth via `Client_OAuth.py`. |
| `Client_OAuth.py`  | Manages OAuth 2.0 flow to obtain access token from Cisco. |
| `Web_Server.py`    | Local HTTP server listening on port `5963`. Captures redirect with `Auth Code` and `State` from Cisco's auth service. |


### Prerequisites

- Python 3.7+
- Python libraries - requests, Openpyxl, xmltodict
- Cisco UCCX with REST API access enabled
- WxCC tenant with API access and registered integration for OAuth
- A browser available for handling OAuth redirect
- Local certificates to interact with Cisco Authorization server over HTTPS. You can use OpenSSL to generate local self signed certificates. This will give you the private key and the certificate which will then need to be used while wrapping the socket request. Once you have the private key and the certificate then update their information in the **"keyfile"** and **"certfile"** variables in `Web_Server.py` file.

---
## Customization

### Environment Variables

You will need to create the following "Environment Variables":

* **CCX_INSTANCE** : FQDN of on-prem UCCX instance
* **CCX_TOKEN** : Base64 value of the username/password combination of an account that has access to UCCX REST APIs
* **WxCC_INSTANCE** : FQDN of WxCC Tenant
* **ORG_ID** : WxCC Organization ID
* **WxCC_CLIENT_ID** : Client ID that you get when you register your app on `https://developer.webex-cx.com/my-apps`
* **WxCC_CLIENT_SECRET** : Client Secret value that you get when you register your app on `https://developer.webex-cx.com/my-apps`
* **WxCC_AUTH_URL** : Authorization URL you get after you have registered your app on `https://developer.webex-cx.com/my-apps`
* **WxCC_TOKEN_URL** : Set it to `https://webexapis.com/v1/access_token`
* **WxCC_REDIRECT_URI** : Set it to the URL where you want Cisco to redirect OAuth response. In the case of this example, it will be `https://localhost:5963`

---

## How to Run

**1. Clone the repository**
  ```bash
   git clone https://github.com/simranjit-uc/Cisco-UCCX-WxCC-Migration.git
  ```

> Ensure you whitelist `http://localhost:5963` as a valid **redirect URI** in your local system.

**2. Start the Migration**
* Move to the newly generated directory.
   ```bash
   cd Cisco-UCCX-WxCC-Migration
   ```
* Run the local web server first. This is where the Cisco Auth server will be redirecting the user to as a part of the OAuth process. This needs to be run first so that the incoming auth-code and state values from the Cisco Auth server can be captured.
  ```bash
  python3 Web_Server.py
  ```
* Open a new session and run the main file to start the migration process.
   ```bash
   python3 main.py
   ```

**3. Follow Prompts**

   The app will guide you through the following steps with step by step update:

  * Initiating OAuth via browser (token is returned via `Web_Server.py`)
  * Connecting to UCCX (via Basic Auth)
  * Generating `CCX-Details.xlsx`
  * Transforming to `WxCC Details.xlsx`
  * Uploading configuration to WxCC

---

## OAuth Flow (Explained)

1. `Main.py` calls `Client_OAuth.py` to start the OAuth process.
2. `Client_OAuth.py`:

   * Opens a browser pointing to Cisco OAuth URL.
   * Redirect URI is set to `http://localhost:5963`.
3. `Web_Server.py`:

   * Listens on port `5963` for incoming HTTP redirect from Cisco.
   * The communication between this app and Cisco's Auth server is happening over SSl/TLS. This is where the locally generated certificate and its private key will come in.
   * Captures `auth_code` and `state` values.
4. `Client_OAuth.py` exchanges `auth_code` for access/refresh tokens and returns them to `WxCC.py`.

> You must approve the authentication in the browser for the flow to complete.

---

## Output Files

* `CCX-Details.xlsx`: Raw UCCX configuration as extracted via API.
* `WxCC Details.xlsx`: Transformed configuration compatible with WxCC.

---

## Example Use Case - Video Demo

[[Watch the demo]](https://www.youtube.com/watch?v=gK3W_2sHtIs)

************ **Sample Screenshot** ************
---

[![Watch the demo](https://learnuccollab.wordpress.com/wp-content/uploads/2025/05/thumbnail-wp-e1748122953676.jpg)](https://www.youtube.com/watch?v=gK3W_2sHtIs)

In the above sample use-case, we are migrating the following UCCX configuration for a typical "Telecom Company" to a WxCC tenant.

* 7 Applications
* 7 CSQs
* 7 Teams
* 9 Skills
* 31 Wrap-Up and Reason Codes
* 2 Phone books with 15 Contacts

The app does the following in less than a minute:

1. Extracts all 63 odd items from UCCX and logs them.
2. Transforms them into a WxCC compatible format.
3. Maps Teams, CSQs and Skills to WxCC Team, Queue and Skill Profile concepts.
4. Translates Reason/Wrapup Codes and Phonebooks into Aux Codes and Address Books.

---

## Future Contributions

Contributions are welcome! In order to do that, please:

1. Fork this repo and clone it to work locally.
2. Create a new branch.
3. Add your new feature and test it.
4. Commit your changes.
5. Open a pull request with a clear description.

---

## License

MIT License. See [LICENSE](LICENSE) for more details.

---

## Contact

If you want to support or provide any suggestions or simply bounce any technical ideas, please reach out to me at the following:

* 🌐 Website: [learnuccollab.com](https://learnuccollab.com)
* 🔗 LinkedIn: [Simranjit Singh](https://www.linkedin.com/in/simranjit-singh-455751b9)
* 📺 YouTube: [@learnuccollab](https://www.youtube.com/@learnuccollab)
  
