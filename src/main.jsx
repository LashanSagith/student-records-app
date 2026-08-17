import React from "react";
import ReactDOM from "react-dom/client";
import { PublicClientApplication, EventType } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { msalConfig } from "./authConfig";
import App from "./App";
import "./index.css";

function renderFatalError(err) {
  // If MSAL fails to start (e.g. CLIENT_ID / TENANT_ID in authConfig.js are
  // still placeholders, or don't match a real Azure App Registration), show
  // a real message instead of leaving a blank white page with only a
  // console error nobody sees.
  const root = document.getElementById("root");
  root.innerHTML = `
    <div style="min-height:100vh;display:flex;align-items:center;justify-content:center;background:#E9EBEE;font-family:sans-serif;padding:24px;">
      <div style="background:#fff;border:1px solid #DADCE0;border-radius:12px;padding:28px;max-width:480px;text-align:left;">
        <div style="color:#B23A48;font-weight:700;font-size:15px;margin-bottom:8px;">The app couldn't start</div>
        <div style="color:#616771;font-size:13px;line-height:1.6;margin-bottom:14px;">
          This almost always means <code>CLIENT_ID</code> / <code>TENANT_ID</code> in
          <code>src/authConfig.js</code> are still the placeholder text, or don't match
          a real Azure App Registration yet. See README.md → "Azure App Registration".
        </div>
        <div style="background:#F6E3E5;border-radius:8px;padding:10px 12px;color:#B23A48;font-size:11px;font-family:monospace;word-break:break-all;">
          ${(err && (err.errorMessage || err.message)) || String(err)}
        </div>
      </div>
    </div>`;
}

try {
  const msalInstance = new PublicClientApplication(msalConfig);

  // Keep MSAL's "active account" pointed at whoever most recently signed in,
  // so a page refresh doesn't lose track of who's logged in.
  msalInstance
    .initialize()
    .then(() => {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);
      }

      msalInstance.addEventCallback((event) => {
        if (
          (event.eventType === EventType.LOGIN_SUCCESS || event.eventType === EventType.ACQUIRE_TOKEN_SUCCESS) &&
          event.payload?.account
        ) {
          msalInstance.setActiveAccount(event.payload.account);
        }
      });

      ReactDOM.createRoot(document.getElementById("root")).render(
        <React.StrictMode>
          <MsalProvider instance={msalInstance}>
            <App />
          </MsalProvider>
        </React.StrictMode>
      );
    })
    .catch(renderFatalError);
} catch (err) {
  renderFatalError(err);
}
