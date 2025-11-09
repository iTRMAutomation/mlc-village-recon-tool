import React, { useEffect, useMemo, useRef, useState } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import metlifecareLogo from "./assets/metlifecare-logo.svg";

/**
 * Photo Report Uploader – Mobile‑friendly SPA
 * Stack: React + MSAL (browser) + Microsoft Graph
 * Auth: Delegated (users sign in with Entra ID)
 * Data: Upload photos to SharePoint library folder and create ONE list item
 *       per submission containing multiple photo URLs.
 *
 * ──────────────────────────────────────────────────────────────────────────────
 * WHAT'S NEW (Village dropdown)
 *   • Auto-detects the **Village** column on your target list
 *   • If the column is a **Choice** type, the component renders a <select>
 *     populated from the column's configured choices via Graph.
 *   • If the column is not Choice (e.g., Text) or choices cannot be read,
 *     it gracefully falls back to a free‑text <input> (original behavior).
 *
 * Also retained recent changes/fixes:
 *   • FIX: Proper folder create endpoint at root (no `root::/children`)
 *   • Popup‑only auth (iframe‑safe)
 *   • Runtime discovery of internal list field names
 *   • Photo column type detection (single URL vs JSON array)
 *   • Append extra photo URLs into Notes for readability
 *   • Diagnostics panel with path builder self-tests
 */

// ── UTILS (declared BEFORE any use) ──────────────────────────────────────────
const toStr = (v: any, fallback: string = "") => (v == null ? fallback : String(v));
function normalizeRedirectUri(v: any) {
  const s = toStr(v, "");
  if (!s) return "";
  return /\/$/.test(s) ? s : s + "/";
}
const isGuid = (s: any) => /^{?[0-9a-fA-F]{8}(-?[0-9a-fA-F]{4}){3}-?[0-9a-fA-F]{12}}?$/.test(toStr(s));
const trimSlashes = (p: any) => toStr(p).replace(/^\/+|\/+$/g, "");
const isInIframe = (() => { try { return window.self !== window.top; } catch { return true; } })();

const NZ_TIME_ZONE = "Pacific/Auckland";

const formatDateTimeLocal = (date: Date, timeZone: string) => {
  const formatter = new Intl.DateTimeFormat("en-NZ", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hourCycle: "h23",
  });
  const parts = formatter.formatToParts(date).reduce((acc, part) => {
    if (part.type !== "literal") acc[part.type] = part.value;
    return acc;
  }, {} as Record<string, string>);
  const year = parts.year ?? "0000";
  const month = parts.month ?? "01";
  const day = parts.day ?? "01";
  const hour = parts.hour ?? "00";
  const minute = parts.minute ?? "00";
  return year + "-" + month + "-" + day + "T" + hour + ":" + minute;
};

const getTimeZoneOffset = (date: Date, timeZone: string) => {
  const formatter = new Intl.DateTimeFormat("en-NZ", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hourCycle: "h23",
  });
  const parts = formatter.formatToParts(date).reduce((acc, part) => {
    if (part.type !== "literal") acc[part.type] = part.value;
    return acc;
  }, {} as Record<string, string>);
  const year = Number(parts.year ?? "0");
  const month = Number(parts.month ?? "1");
  const day = Number(parts.day ?? "1");
  const hour = Number(parts.hour ?? "0");
  const minute = Number(parts.minute ?? "0");
  const second = Number(parts.second ?? "0");
  const asUtc = Date.UTC(year, month - 1, day, hour, minute, second);
  return asUtc - date.getTime();
};

const zonedDateTimeStringToUtc = (value: string, timeZone: string) => {
  if (!value) return new Date();
  const segments = value.split("T");
  if (segments.length !== 2) return new Date(value);
  const datePart = segments[0];
  const timePart = segments[1];
  const dateBits = datePart.split("-");
  const timeBits = timePart.split(":");
  if (dateBits.length !== 3 || timeBits.length < 2) return new Date(value);
  const year = Number(dateBits[0]);
  const month = Number(dateBits[1]);
  const day = Number(dateBits[2]);
  const hour = Number(timeBits[0]);
  const minute = Number(timeBits[1]);
  if ([year, month, day, hour, minute].some((n) => Number.isNaN(n))) {
    return new Date(value);
  }
  const naiveUtc = new Date(Date.UTC(year, month - 1, day, hour, minute));
  const offset = getTimeZoneOffset(naiveUtc, timeZone);
  return new Date(naiveUtc.getTime() - offset);
};
const formatDateTimeForSharePoint = (value: string, timeZone: string) => {
  const date = zonedDateTimeStringToUtc(value, timeZone);
  const parts = new Intl.DateTimeFormat("en-NZ", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hourCycle: "h23",
  }).formatToParts(date).reduce((acc, part) => {
    if (part.type !== "literal") acc[part.type] = part.value;
    return acc;
  }, {} as Record<string, string>);
  const offsetMinutesTotal = Math.round(getTimeZoneOffset(date, timeZone) / 60000);
  const sign = offsetMinutesTotal >= 0 ? "+" : "-";
  const absMinutes = Math.abs(offsetMinutesTotal);
  const offsetHours = Math.floor(absMinutes / 60);
  const offsetMinutes = absMinutes % 60;
  const offset = `${sign}${String(offsetHours).padStart(2, "0")}:${String(offsetMinutes).padStart(2, "0")}`;
  const year = parts.year ?? "0000";
  const month = parts.month ?? "01";
  const day = parts.day ?? "01";
  const hour = parts.hour ?? "00";
  const minute = parts.minute ?? "00";
  const second = parts.second ?? "00";
  return `${year}-${month}-${day}T${hour}:${minute}:${second}${offset}`;
};
// ── ⬇️ REQUIRED CONFIG – EDIT THESE VALUES ⬇️ ────────────────────────────────
const CONFIG = {
  // NOTE: Browser apps cannot use certificate credential flow. Delegated MSAL is used here.
  tenantId: "3a508ea9-e4ed-4d0b-b4e0-9e903e30072b",
  clientId: "495c7630-3fe2-4296-9213-101932aa2d27",
  redirectUri: normalizeRedirectUri((window as any)?.location?.origin ?? ""), // Ensure trailing '/'

  // Graph base URL (NO version here) + version separately
  graphBaseUrl: "https://graph.microsoft.com",
  graphVersion: "v1.0",

  // Your SharePoint site
  spSiteHostname: "metlifecare.sharepoint.com",

  spSitePath: "/sites/windows11upgrade",

  // Target List & Library
  spListIdOrName: "W11VillageReconList",
  spLibraryDriveIdOrName: "Documents", // or "Shared Documents" – confirm in UI

  // Subfolder inside the library where images should land
  libraryFolderPath: "VillageReconPhotos", // created under YYYY/MM
};

// Scopes required (delegated). Keep minimal but sufficient for upload + list write.
const GRAPH_SCOPES = [
  "User.Read",
  "Sites.ReadWrite.All",
  "Files.ReadWrite.All",
  "offline_access",
];

const lastSegment = toStr(CONFIG.spSitePath).split("/").filter(Boolean).slice(-1)[0] || "";

// ── MSAL SETUP ───────────────────────────────────────────────────────────────
const msal = new PublicClientApplication({
  auth: {
    clientId: toStr(CONFIG.clientId),
    authority: `https://login.microsoftonline.com/${toStr(CONFIG.tenantId)}`,
    redirectUri: toStr(CONFIG.redirectUri),
  },
  cache: { cacheLocation: "localStorage" },
  system: { allowRedirectInIframe: false }, // popup‑only flows
});

// ── ENV/NETWORK PROBES ──────────────────────────────────────────────────────
async function probeGraphReachable() {
  const base = toStr(CONFIG.graphBaseUrl, "https://graph.microsoft.com").replace(/\/$/, "");
  const ver = toStr(CONFIG.graphVersion, "v1.0").replace(/^\/+|\/+$/g, "");
  const metaUrl = `${base}/${ver}/$metadata`;
  try {
    const res = await fetch(metaUrl, { method: "GET", mode: "cors" });
    return { ok: true, status: res.status, metaUrl };
  } catch (e) {
    return { ok: false, error: String(e), metaUrl };
  }
}

async function probeSharePointHost(hostname: string) {
  const host = toStr(hostname);
  if (!host) return { ok: false, error: "spSiteHostname is empty" };
  try {
    await fetch(`https://${host}/_api/v2.1/drive/root`, { mode: "no-cors" });
    return { ok: true };
  } catch (e) { return { ok: false, error: String(e) }; }
}

// ── GRAPH CLIENT ─────────────────────────────────────────────────────────────
function useGraphClient(account: any) {
  return useMemo(() => {
    if (!account) return null;
    return Client.initWithMiddleware({
      baseUrl: toStr(CONFIG.graphBaseUrl, "https://graph.microsoft.com"),
      defaultVersion: toStr(CONFIG.graphVersion, "v1.0"),
      authProvider: {
        getAccessToken: async () => {
          const current = account || msal.getAllAccounts()[0];
          try {
            const silent = await msal.acquireTokenSilent({ account: current, scopes: GRAPH_SCOPES });
            return silent.accessToken;
          } catch {
            const inter = await msal.acquireTokenPopup({ account: current, scopes: GRAPH_SCOPES });
            return inter.accessToken;
          }
        },
      },
    });
  }, [account]);
}

function rewrapNetworkError(err: any, contextPath: string) {
  const msg = String(err?.message || err);
  const isNetwork = /Failed to fetch|NetworkError|TypeError|Load failed|ERR_NETWORK/i.test(msg);
  if (!isNetwork) return err;
  const hints = [
    `Endpoint: ${contextPath}`,
    `Check CONFIG.spSiteHostname ('${toStr(CONFIG.spSiteHostname)}') and spSitePath ('${toStr(CONFIG.spSitePath)}')`,
    `If running on HTTP, ensure corporate policies/firewalls allow cross-origin calls to graph.microsoft.com and *.sharepoint.com`,
    `Confirm your Entra app's SPA Redirect URI matches this origin: ${toStr((window as any)?.location?.origin)}/`,
  ].join("\\n• ");
  return new Error(`Network/CORS issue when calling Graph. Original: ${msg}\\n• ${hints}`);
}

function useSafeGraph(graph: any) {
  return {
    get: async (path: string) => { try { return await graph.api(path).get(); } catch (e) { throw rewrapNetworkError(e, path); } },
    post: async (path: string, body: any) => { try { return await graph.api(path).post(body); } catch (e) { throw rewrapNetworkError(e, path); } },
    patch: async (path: string, body: any) => { try { return await graph.api(path).patch(body); } catch (e) { throw rewrapNetworkError(e, path); } },
    putContent: async (path: string, blob: Blob) => { try { return await graph.api(path).put(blob); } catch (e) { throw rewrapNetworkError(e, path); } },
    createUploadSession: async (path: string, body: any) => { try { return await graph.api(path).post(body); } catch (e) { throw rewrapNetworkError(e, path); } },
  };
}

// ── UI HELPERS ───────────────────────────────────────────────────────────────
type FieldProps = {
  label: string;
  hint?: string;
  full?: boolean;
  controlId: string;
  children: React.ReactNode;
};

const Field: React.FC<FieldProps> = ({ label, hint, full, controlId, children }) => (
  <div className={`form-field${full ? " form-field--full" : ""}`}>
    <label className="form-label" htmlFor={controlId}>{label}</label>
    {hint ? <p className="form-hint">{hint}</p> : null}
    <div className="form-input">
      {React.Children.map(children, (child, index) => {
        if (index === 0 && React.isValidElement(child) && !child.props.id) {
          return React.cloneElement(child, { id: controlId });
        }
        return child;
      })}
    </div>
  </div>
);

type ButtonVariant = "primary" | "secondary" | "ghost" | "success";

type ButtonProps = React.ButtonHTMLAttributes<HTMLButtonElement> & {
  variant?: ButtonVariant;
};

const Button: React.FC<ButtonProps> = ({ variant = "primary", className = "", children, ...props }) => {
  const classes = ["btn", `btn-${variant}`, props.disabled ? "btn-disabled" : "", className].filter(Boolean).join(" ");
  return (
    <button {...props} className={classes}>
      {children}
    </button>
  );
};

const palette = {
  background: "#F6F1EA",
  surface: "#FFFFFF",
  primary: "#1F214B",
  accent: "#5A5FF4",
  accentSoft: "#E9E8FF",
  text: "#202134",
  textMuted: "#5C5E75",
  border: "rgba(38, 39, 70, 0.14)",
  success: "#2F8F75",
  onPrimary: "#F7F7FF",
};

const GLOBAL_STYLE_ID = "village-uploader-styles";

function ensureGlobalStyles() {
  if (typeof document === "undefined" || document.getElementById(GLOBAL_STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = GLOBAL_STYLE_ID;
  style.textContent = `
    :root {
      --brand-background: ${palette.background};
      --brand-surface: ${palette.surface};
      --brand-primary: ${palette.primary};
      --brand-accent: ${palette.accent};
      --brand-accent-soft: ${palette.accentSoft};
      --brand-text: ${palette.text};
      --brand-text-muted: ${palette.textMuted};
      --brand-border: ${palette.border};
      --brand-success: ${palette.success};
      --brand-on-primary: ${palette.onPrimary};
    }
    *, *::before, *::after { box-sizing: border-box; }
    html, body { margin: 0; min-height: 100%; font-family: "Inter", "Segoe UI", "SF Pro Display", -apple-system, BlinkMacSystemFont, "Helvetica Neue", sans-serif; background: var(--brand-background); color: var(--brand-text); }
    body { -webkit-font-smoothing: antialiased; line-height: 1.55; }
    #root { min-height: 100%; }
    .app-shell { min-height: 100vh; padding: clamp(24px, 4vw, 56px) clamp(18px, 5vw, 64px); background: var(--brand-background); }
    .app-layout { width: min(960px, 100%); margin: 0 auto; display: grid; gap: clamp(24px, 3vw, 32px); }
    .app-header { display: flex; flex-wrap: wrap; justify-content: space-between; align-items: flex-start; gap: 24px; padding: clamp(24px, 4vw, 36px); border-radius: 28px; background: var(--brand-primary); color: var(--brand-on-primary); box-shadow: 0 34px 60px -45px rgba(20, 21, 50, 0.78); }
    .app-brand { display: flex; align-items: center; gap: clamp(16px, 3vw, 28px); min-width: 0; }
    .app-heading { display: grid; gap: 6px; min-width: 0; }
    .app-logo { width: clamp(128px, 18vw, 208px); height: auto; border-radius: 20px; flex-shrink: 0; box-shadow: 0 24px 42px -30px rgba(8, 9, 30, 0.6); background: rgba(255, 255, 255, 0.08); }
    .app-tagline { font-size: 0.75rem; letter-spacing: 0.28em; text-transform: uppercase; opacity: 0.72; margin: 0 0 8px; display: block; }
    .app-title { margin: 0; font-size: clamp(1.8rem, 3.4vw, 2.6rem); font-weight: 600; }
    .app-subtitle { margin: 8px 0 0; max-width: 420px; opacity: 0.78; font-size: 0.95rem; line-height: 1.6; }
    .auth-block { display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }
    .auth-status { font-size: 0.8rem; opacity: 0.65; }
    .auth-identity { display: inline-flex; align-items: center; padding: 6px 14px; border-radius: 999px; background: rgba(255, 255, 255, 0.12); color: rgba(255, 255, 255, 0.86); font-size: 0.85rem; font-weight: 500; }
    .app-main { display: grid; gap: clamp(24px, 3vw, 32px); }
    .panel { background: var(--brand-surface); border-radius: 26px; padding: clamp(24px, 3.5vw, 36px); box-shadow: 0 40px 74px -58px rgba(24, 25, 60, 0.55); display: grid; gap: 24px; border: 1px solid rgba(255, 255, 255, 0.7); }
    .panel-header { display: flex; align-items: flex-start; justify-content: space-between; gap: 16px; }
    .panel-title { margin: 0; font-size: clamp(1.2rem, 2.2vw, 1.6rem); font-weight: 600; color: var(--brand-primary); display: block; }
    .panel-subtitle { margin: 6px 0 0; font-size: 0.95rem; color: var(--brand-text-muted); max-width: 460px; display: block; }
    .diagnostics-panel .panel-header { gap: 18px; }
    .diagnostics-header { align-items: center; }
    .diagnostics-header .btn { flex-shrink: 0; }
    .diagnostics-toggle { display: flex; align-items: flex-start; gap: 16px; border: none; background: none; padding: 0; color: inherit; cursor: pointer; text-align: left; flex: 1; }
    .diagnostics-toggle:hover { opacity: 0.92; }
    .diagnostics-toggle:focus-visible { outline: none; box-shadow: 0 0 0 4px rgba(90, 95, 244, 0.25); border-radius: 18px; }
    .diagnostics-toggle-icon { width: 34px; height: 34px; border-radius: 50%; border: 1px solid rgba(38, 39, 70, 0.16); display: grid; place-items: center; background: rgba(38, 39, 70, 0.05); transition: transform 0.2s ease, background 0.2s ease, border-color 0.2s ease; color: var(--brand-primary); flex-shrink: 0; }
    .diagnostics-toggle-icon::before { content: ""; border: solid currentColor; border-width: 0 2px 2px 0; display: inline-block; padding: 6px; transform: rotate(45deg); transition: transform 0.2s ease; }
    .diagnostics-panel--open .diagnostics-toggle-icon { background: rgba(90, 95, 244, 0.12); border-color: rgba(90, 95, 244, 0.3); }
    .diagnostics-panel--open .diagnostics-toggle-icon::before { transform: rotate(135deg); }
    .diagnostics-text { display: grid; gap: 6px; align-content: center; }
    .diagnostics-content { display: grid; gap: 20px; }
    .diagnostics-content[hidden] { display: none; }
    .panel-badge { align-self: flex-start; padding: 6px 16px; border-radius: 999px; background: var(--brand-accent-soft); color: var(--brand-primary); font-weight: 600; font-size: 0.85rem; }
    .form-grid { display: grid; gap: 18px 24px; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); }
    .form-field { display: grid; gap: 10px; }
    .form-field--full { grid-column: 1 / -1; }
    .form-label { font-size: 0.75rem; letter-spacing: 0.08em; text-transform: uppercase; font-weight: 600; color: var(--brand-text-muted); margin: 0; }
    .form-hint { margin: -4px 0 0; font-size: 0.8rem; color: var(--brand-text-muted); }
    .form-input { display: grid; gap: 12px; }
    .form-control { width: 100%; border-radius: 18px; border: 1px solid var(--brand-border); padding: 12px 16px; font-size: 0.95rem; background: rgba(255, 255, 255, 0.75); transition: border-color 0.15s ease, box-shadow 0.15s ease, background 0.15s ease, transform 0.12s ease; color: var(--brand-text); }
    .form-control:focus-visible { outline: none; border-color: rgba(90, 95, 244, 0.75); box-shadow: 0 0 0 4px rgba(90, 95, 244, 0.2); background: #fff; }
    textarea.form-control { resize: vertical; min-height: 140px; }
    select.form-control { appearance: none; background-image: linear-gradient(135deg, transparent 0%, transparent calc(100% - 1.2rem), rgba(32, 33, 52, 0.08) calc(100% - 1.2rem), rgba(32, 33, 52, 0.08) 100%), url('data:image/svg+xml;utf8,<svg xmlns="http://www.w3.org/2000/svg" width="12" height="8" viewBox="0 0 12 8"><path fill="%23202134" d="M1.41.59 6 5.17 10.59.59 12 2l-6 6-6-6z"/></svg>'); background-repeat: no-repeat, no-repeat; background-position: right 18px center, right 18px center; background-size: 1px 100%, 12px 8px; padding-right: 48px; }
    .file-input { border-style: dashed; border-color: rgba(38, 39, 70, 0.2); background: rgba(38, 39, 70, 0.035); cursor: pointer; }
    .file-input:hover { border-color: rgba(38, 39, 70, 0.38); background: rgba(38, 39, 70, 0.06); }
    .file-input::file-selector-button { margin-right: 16px; border: none; border-radius: 14px; padding: 10px 16px; font-weight: 600; background: linear-gradient(135deg, var(--brand-accent), var(--brand-primary)); color: var(--brand-on-primary); cursor: pointer; transition: transform 0.15s ease, box-shadow 0.2s ease; }
    .file-input::file-selector-button:hover { transform: translateY(-1px); box-shadow: 0 16px 28px -18px rgba(38, 39, 70, 0.45); }
    .photo-grid { display: grid; gap: 12px; grid-template-columns: repeat(auto-fill, minmax(96px, 128px)); justify-content: flex-start; }
    .photo-thumb { width: 100%; height: 100%; aspect-ratio: 1; object-fit: cover; border-radius: 16px; box-shadow: 0 24px 40px -32px rgba(38, 39, 70, 0.5); }
    .btn { display: inline-flex; align-items: center; justify-content: center; gap: 0.5rem; border: none; border-radius: 999px; padding: 0.75rem 1.6rem; font-weight: 600; font-size: 0.95rem; cursor: pointer; transition: transform 0.15s ease, box-shadow 0.2s ease, opacity 0.2s ease, background 0.2s ease, color 0.2s ease; text-decoration: none; }
    .btn:focus-visible { outline: none; box-shadow: 0 0 0 4px rgba(90, 95, 244, 0.25); }
    .btn-primary { background: linear-gradient(135deg, var(--brand-accent), var(--brand-primary)); color: var(--brand-on-primary); box-shadow: 0 22px 38px -24px rgba(38, 39, 70, 0.78); }
    .btn-primary:hover { transform: translateY(-1px); box-shadow: 0 28px 46px -28px rgba(38, 39, 70, 0.78); }
    .btn-secondary { background: rgba(38, 39, 70, 0.08); color: var(--brand-primary); border: 1px solid rgba(38, 39, 70, 0.16); }
    .btn-secondary:hover { background: rgba(38, 39, 70, 0.12); border-color: rgba(38, 39, 70, 0.26); }
    .btn-ghost { background: rgba(255, 255, 255, 0.14); color: rgba(255, 255, 255, 0.92); border: 1px solid rgba(255, 255, 255, 0.35); }
    .btn-ghost:hover { background: rgba(255, 255, 255, 0.24); }
    .btn-success { background: var(--brand-success); color: #f4fffb; }
    .btn-success:hover { transform: translateY(-1px); box-shadow: 0 22px 40px -26px rgba(47, 143, 117, 0.55); }
    .btn-disabled, .btn:disabled { opacity: 0.55; cursor: not-allowed; transform: none; box-shadow: none; }
    .activity-log { background: rgba(38, 39, 70, 0.035); border: 1px solid rgba(38, 39, 70, 0.12); border-radius: 20px; padding: 18px 20px; min-height: 80px; white-space: pre-wrap; font-size: 0.92rem; color: var(--brand-text-muted); }
    .diagnostics-list { list-style: none; margin: 0; padding: 0; display: grid; gap: 10px; }
    .diagnostics-list li { padding: 12px 16px; border-radius: 16px; background: rgba(38, 39, 70, 0.04); color: var(--brand-text-muted); font-size: 0.9rem; }
    .diagnostics-list li strong { color: var(--brand-primary); }
    .diagnostics-json { background: rgba(38, 39, 70, 0.03); border-radius: 16px; padding: 16px; font-family: "SFMono-Regular", "Consolas", "Liberation Mono", monospace; font-size: 0.85rem; color: var(--brand-primary); white-space: pre-wrap; }
    .app-footer { text-align: center; font-size: 0.82rem; color: var(--brand-text-muted); padding-bottom: 12px; }
    @media (max-width: 640px) {
      .app-header { padding: 22px; border-radius: 22px; }
      .app-brand { flex-direction: column; align-items: flex-start; }
      .app-logo { width: min(220px, 100%); }
      .panel { padding: 22px; border-radius: 22px; }
      .form-grid { grid-template-columns: 1fr; }
      .panel-header { flex-direction: column; align-items: flex-start; }
      .auth-block { width: 100%; justify-content: space-between; }
      .diagnostics-header { flex-direction: column; align-items: stretch; }
      .diagnostics-toggle { width: 100%; }
      .diagnostics-header .btn { width: 100%; justify-content: center; }
    }
  `;
  document.head.appendChild(style);
}

// ── PATH HELPERS (fix root children path) ────────────────────────────────────
function childrenEndpoint(driveId: string, parentPath: string) {
  const clean = trimSlashes(parentPath);
  return clean
    ? `/drives/${driveId}/root:/${encodeURI(clean)}:/children`
    : `/drives/${driveId}/root/children`;
}

// ── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [msalReady, setMsalReady] = useState(false);
  const [account, setAccount] = useState<any>(null);
  const [title, setTitle] = useState("");
  const [village, setVillage] = useState("");
  const [notes, setNotes] = useState("");
  const [capturedOn, setCapturedOn] = useState(() => formatDateTimeLocal(new Date(), NZ_TIME_ZONE));
  const [files, setFiles] = useState<File[]>([]);
  const [busy, setBusy] = useState(false);
  const [log, setLog] = useState<string[]>([]);
  const [previews, setPreviews] = useState<{ key: string; url: string }[]>([]);
  const isMobileBrowser = useMemo(() => {
    if (typeof navigator === "undefined") return false;
    const ua = navigator.userAgent || "";
    return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(ua);
  }, []);
  const [diagnosticsOpen, setDiagnosticsOpen] = useState(false);
  const inputRef = useRef<HTMLInputElement | null>(null);
  useEffect(() => {
    return () => {
      previews.forEach((preview) => URL.revokeObjectURL(preview.url));
    };
  }, [previews]);

  // Resolved IDs after diagnostics/first run
  const [resolved, setResolved] = useState<{ siteId: string | null; listId: string | null; driveId: string | null }>({ siteId: null, listId: null, driveId: null });

  // NEW: Column choice caches (Village + CategoryOptions)
  const [titleChoices, setTitleChoices] = useState<string[] | null>(null);
  const [villageChoices, setVillageChoices] = useState<string[] | null>(null);
  const [choicesPrimed, setChoicesPrimed] = useState(false);

  useEffect(() => {
    if (!Array.isArray(titleChoices) || titleChoices.length === 0) return;
    setTitle((prev) => (titleChoices.includes(prev) ? prev : ""));
  }, [titleChoices]);

  useEffect(() => { ensureGlobalStyles(); }, []);

  useEffect(() => {
    const boot = async () => {
      try {
        await msal.initialize();
        setMsalReady(true);
        const accts = msal.getAllAccounts();
        setAccount(accts[0] ?? null);
      } catch (e: any) {
        console.error(e);
        setLog((l) => [...l, `${new Date().toLocaleTimeString()}: MSAL init failed → ${e.message}`]);
      }
    };
    boot();
  }, []);

  const graph = useGraphClient(account);
  const sgraph = graph ? useSafeGraph(graph) : null;
  useEffect(() => {
    if (!account || !msalReady || !sgraph || choicesPrimed) return;
    let cancelled = false;
    const preloadChoices = async () => {
      try {
        const currentSiteId = resolved.siteId || (await getSiteId());
        const currentListId = resolved.listId || (await resolveListId(currentSiteId, CONFIG.spListIdOrName));
        const currentDriveId = resolved.driveId || (await resolveDriveId(currentSiteId, CONFIG.spLibraryDriveIdOrName));
        const { byLower, choicesByName, metaByName } = await getColumns(currentSiteId, currentListId);
        const villageName = chooseField(byLower, metaByName, ["Village"], { requireWritable: true });
        const categoryOptionsName = chooseField(byLower, metaByName, ["CategoryOptions", "Category Options", "Category"]);
        const availableVillageChoices = villageName ? (choicesByName[villageName] ?? []) : [];
        const availableTitleChoices = categoryOptionsName ? (choicesByName[categoryOptionsName] ?? []) : [];
        if (cancelled) return;
        setResolved({ siteId: currentSiteId, listId: currentListId, driveId: currentDriveId });
        setVillageChoices(availableVillageChoices.length ? availableVillageChoices : null);
        setTitleChoices(availableTitleChoices.length ? availableTitleChoices : null);
        setChoicesPrimed(true);
      } catch (e: any) {
        if (cancelled) return;
        setChoicesPrimed(true);
        setTitleChoices(null);
        setVillageChoices(null);
        setLog((l) => [...l, `${new Date().toLocaleTimeString()}: Choice preload failed - ${e.message}`]);
      }
    };
    preloadChoices();
    return () => { cancelled = true; };
  }, [account, msalReady, sgraph, choicesPrimed, resolved.siteId, resolved.listId, resolved.driveId]);

  const addLog = (msg: string) => setLog((l) => [...l, `${new Date().toLocaleTimeString()}: ${msg}`]);

  const resetSharePointState = () => {
    setResolved({ siteId: null, listId: null, driveId: null });
    setTitleChoices(null);
    setVillageChoices(null);
    setChoicesPrimed(false);
  };

  const toggleDiagnostics = () => setDiagnosticsOpen((open) => !open);
  const scheduleMobileRefresh = () => {
    if (!isMobileBrowser || typeof window === "undefined") return;
    const reload = () => {
      try { window.location.reload(); } catch { /* noop */ }
    };
    const requestIdle = (window as any).requestIdleCallback;
    if (typeof requestIdle === "function") {
      requestIdle.call(window, reload, { timeout: 750 });
    } else {
      window.setTimeout(reload, 350);
    }
  };

  const signIn = async () => {
    if (!msalReady) {
      addLog("Auth is still initializing. Please retry in a moment.");
      return;
    }
    try {
      await msal.loginPopup({ scopes: GRAPH_SCOPES, prompt: "select_account" });
      const accts = msal.getAllAccounts();
      resetSharePointState();
      setAccount(accts[0] ?? null);
    } catch (err: any) {
      console.warn("Popup sign-in failed", err);
      addLog(`[Auth] Popup sign-in blocked (${err?.errorCode || err?.message || "unknown"}). Falling back to redirect.`);
      await msal.loginRedirect({ scopes: GRAPH_SCOPES, prompt: "select_account" });
    }
  };

  const signOut = async () => {
    if (!account || !msalReady) return;
    try {
      await msal.logoutPopup({ account });
      setAccount(null);
      resetSharePointState();
    } catch (err: any) {
      console.warn("Popup sign-out failed", err);
      addLog(`[Auth] Popup sign-out blocked (${err?.errorCode || err?.message || "unknown"}). Redirecting to complete sign-out.`);
      await msal.logoutRedirect({ account });
      resetSharePointState();
    }
  };

  // ── Graph helpers ─────────────────────────────────────────────────────────
  const getSiteId = async () => {
    const host = toStr(CONFIG.spSiteHostname);
    const path = toStr(CONFIG.spSitePath);
    const directPath = `/sites/${host}:${path}`;
    try {
      if (!host || !path) throw new Error("Missing spSiteHostname or spSitePath");
      const site = await sgraph!.get(directPath);
      return site.id as string;
    } catch (e: any) {
      addLog(`[Site Resolve] Direct lookup failed: ${e.message}`);
      addLog(`[Site Resolve] Trying discovery via /sites?search=${lastSegment || "(empty)"} ...`);
      const searchTerm = lastSegment || path || host;
      const res = await sgraph!.get(`/sites?search=${encodeURIComponent(toStr(searchTerm))}`);
      const candidates = res?.value || [];
      if (!candidates.length) throw new Error(`Could not discover site. Verify spSiteHostname ('${host}') and spSitePath ('${path}').`);
      const pathLower = path.toLowerCase();
      const hostLower = host.toLowerCase();
      const chosen =
        candidates.find((s: any) => (toStr(s.webUrl).toLowerCase().includes(pathLower)) && (toStr(s.siteCollection?.hostname).toLowerCase().includes(hostLower))) ||
        candidates.find((s: any) => toStr(s.webUrl).toLowerCase().includes(pathLower)) ||
        candidates[0];
      addLog(`[Site Resolve] Using discovered site: ${chosen.webUrl || chosen.id}`);
      return chosen.id as string;
    }
  };

  const resolveListId = async (siteId: string, idOrName: string) => {
    if (isGuid(idOrName)) return idOrName;
    const name = toStr(idOrName).replace(/'/g, "''");
    const res = await sgraph!.get(`/sites/${siteId}/lists?$filter=displayName eq '${name}'`);
    const found = res?.value?.[0];
    if (!found) throw new Error(`List '${idOrName}' not found in site.`);
    return found.id as string;
  };

  const resolveDriveId = async (siteId: string, idOrName: string) => {
    const drives = await sgraph!.get(`/sites/${siteId}/drives`);
    if (isGuid(idOrName)) return idOrName;
    const targetName = decodeURIComponent(toStr(idOrName)).toLowerCase();
    let found = (drives?.value || []).find((d: any) => toStr(d.name).toLowerCase() === targetName);
    if (!found) found = (drives?.value || []).find((d: any) => toStr(d.driveType).toLowerCase() === "documentlibrary");
    if (!found) {
      const names = (drives?.value || []).map((d: any) => d.name).join(", ");
      throw new Error(`Drive '${idOrName}' not found. Available: ${names}`);
    }
    return (found as any).id as string;
  };

  // ENHANCED: capture choice values
  const getColumns = async (siteId: string, listId: string) => {
    const cols = await sgraph!.get(`/sites/${siteId}/lists/${listId}/columns`);
    const byLower: Record<string, string> = {};
    const typeByName: Record<string, string> = {};
    const choicesByName: Record<string, string[]> = {};
    const metaByName: Record<string, { readOnly: boolean; hidden: boolean }> = {};
    const register = (key: string, internal: string) => {
      const normalized = toStr(key).trim().toLowerCase();
      if (!normalized) return;
      if (byLower[normalized]) return; // keep first hit to avoid overriding writable columns with read-only variants
      byLower[normalized] = internal;
    };
    for (const c of cols?.value || []) {
      const internalName = toStr(c?.name);
      if (!internalName) continue;
      const readOnly = Boolean((c as any)?.readOnly);
      const hidden = Boolean((c as any)?.hidden);
      metaByName[internalName] = { readOnly, hidden };
      register(internalName, internalName);
      register(internalName.replace(/_/g, " "), internalName);
      register(internalName.replace(/[^a-z0-9]/gi, ""), internalName);
      const displayName = toStr(c?.displayName);
      if (displayName) {
        register(displayName, internalName);
        register(displayName.replace(/\s+/g, ""), internalName);
        register(displayName.replace(/[^a-z0-9]/gi, ""), internalName);
      }
      if (c?.hyperlinkOrPicture) typeByName[internalName] = "hyperlinkOrPicture";
      else if (c?.text) typeByName[internalName] = "text";
      else if (c?.number) typeByName[internalName] = "number";
      else if (c?.dateTime) typeByName[internalName] = "dateTime";
      else if (c?.boolean) typeByName[internalName] = "boolean";
      else if (c?.choice) {
        typeByName[internalName] = "choice";
        const rawChoices = (c.choice?.choices ?? []).map((choice: any) => toStr(choice).trim()).filter(Boolean);
        choicesByName[internalName] = Array.from(new Set(rawChoices));
      }
      else if (c?.multiChoice) {
        typeByName[internalName] = "multiChoice";
        const rawChoices = (c.multiChoice?.choices ?? []).map((choice: any) => toStr(choice).trim()).filter(Boolean);
        choicesByName[internalName] = Array.from(new Set(rawChoices));
      }
      else if (c?.lookup) typeByName[internalName] = "lookup"; // not auto-populating in UI yet
      else typeByName[internalName] = "unknown";
    }
    return { byLower, typeByName, choicesByName, metaByName };
  };

  const chooseField = (byLower: Record<string, string>, metaByName: Record<string, { readOnly: boolean; hidden: boolean }>, candidates: string[], opts: { requireWritable?: boolean } = {}) => {
    for (const cand of candidates) {
      const key = toStr(cand).toLowerCase();
      const hit = byLower[key];
      if (!hit) continue;
      const meta = metaByName[hit] || { readOnly: false, hidden: false };
      if (opts.requireWritable && meta.readOnly) continue;
      if (meta.hidden) continue;
      return hit;
    }
    return null;
  };

  const formatPhotoValue = (typeByName: Record<string,string>, fieldName: string, urls: string[]) => {
    const t = typeByName[fieldName];
    if (t === "hyperlinkOrPicture") {
      return urls[0] || ""; // single URL
    }
    return JSON.stringify(urls); // Text/multiline fallback
  };

  const ensureFolder = async (driveId: string, baseFolderPath: string, stamp: Date = new Date()) => {

    const year = String(stamp.getFullYear());

    const month = String(stamp.getMonth() + 1).padStart(2, "0");

    const combined = trimSlashes(`${toStr(baseFolderPath)}/${year}/${month}`).replace(/\/{2,}/g, "/");



    let parentPath = "";

    for (const raw of combined.split("/")) {

      const segment = raw.trim();

      if (!segment) continue;

      const currentPath = parentPath ? `${parentPath}/${segment}` : segment;

      try {

        await sgraph!.get(`/drives/${driveId}/root:/${encodeURI(currentPath)}`);

      } catch {

        const endpoint = childrenEndpoint(driveId, parentPath);

        await sgraph!.post(endpoint, { name: segment, folder: {}, "@microsoft.graph.conflictBehavior": "replace" });

      }

      parentPath = currentPath;

    }

    return `${year}/${month}`;

  };




  const buildTargetFileName = (file: File, villageValue: string, stamp: Date) => {
    const originalName = toStr(file?.name, 'photo');
    const dot = originalName.lastIndexOf('.');
    const base = dot > -1 ? originalName.slice(0, dot) : originalName;
    const ext = dot > -1 ? originalName.slice(dot + 1) : '';
    const timestamp = [
      String(stamp.getFullYear()),
      String(stamp.getMonth() + 1).padStart(2, '0'),
      String(stamp.getDate()).padStart(2, '0'),
      String(stamp.getHours()).padStart(2, '0'),
      String(stamp.getMinutes()).padStart(2, '0'),
      String(stamp.getSeconds()).padStart(2, '0'),
    ].join('');
    const sanitize = (segment: string) => toStr(segment).trim().replace(/[^a-zA-Z0-9]+/g, '-').replace(/-+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60);
    const parts = [villageValue, base].map((p) => sanitize(p)).filter(Boolean);
    const safeBase = parts.length ? parts.join('_') : 'upload';
    const safeExt = ext ? ('.' + ext.toLowerCase()) : '';
    return timestamp + '_' + safeBase + safeExt;
  };
  const uploadOne = async (driveId: string, baseFolderPath: string, file: File, villageValue: string) => {

    const stamp = new Date();

    const pathSub = await ensureFolder(driveId, baseFolderPath, stamp);

    const fileName = buildTargetFileName(file, villageValue, stamp);

    const target = `${trimSlashes(toStr(baseFolderPath))}/${pathSub}/${fileName}`.replace(/^\/+/, "");



    if (file.size <= 3_999_999) {

      return await sgraph!.putContent(`/drives/${driveId}/root:/${encodeURI(target)}:/content`, file);

    } else {

      const session = await sgraph!.createUploadSession(`/drives/${driveId}/root:/${encodeURI(target)}:/createUploadSession`, { item: { "@microsoft.graph.conflictBehavior": "rename" } });

      const chunkSize = 5 * 1024 * 1024; // 5MB

      let start = 0;

      while (start < file.size) {

        const end = Math.min(start + chunkSize, file.size);

        const chunk = file.slice(start, end);

        let res: Response;

        try {

          res = await fetch(session.uploadUrl, { method: "PUT", headers: { "Content-Length": `${(chunk as any).size}`, "Content-Range": `bytes ${start}-${end - 1}/${file.size}` }, body: chunk });

        } catch (e) { throw rewrapNetworkError(e, "UploadSession PUT chunk"); }

        if (!res.ok && ![200,201,202].includes(res.status)) { const text = await res.text().catch(() => ""); throw new Error(`Upload chunk failed: HTTP ${res.status} ${res.statusText} ${text}`); }

        start = end;

      }

      const finalItem = await sgraph!.get(`/drives/${driveId}/root:/${encodeURI(target)}`);

      return finalItem;

    }

  };



  const createListItem = async (siteId: string, listId: string, fields: any) => {
    return await sgraph!.post(`/sites/${siteId}/lists/${listId}/items`, { fields });
  };

  const fileKey = (file: File) => `${file.name}::${file.size}::${file.lastModified}`;

  const onFilesChanged = (e: React.ChangeEvent<HTMLInputElement>) => {
    const picked = Array.from(e.target.files || []);
    if (picked.length === 0) return;
    setFiles((prev) => {
      const existingKeys = new Set(prev.map((file) => fileKey(file)));
      const next = [...prev];
      for (const file of picked) {
        const key = fileKey(file);
        if (existingKeys.has(key)) continue;
        existingKeys.add(key);
        next.push(file);
      }
      return next;
    });
    setPreviews((prev) => {
      const next = [...prev];
      const existingKeys = new Set(prev.map((item) => item.key));
      for (const file of picked) {
        const key = fileKey(file);
        if (existingKeys.has(key)) continue;
        const url = URL.createObjectURL(file);
        next.push({ key, url });
        existingKeys.add(key);
      }
      return next;
    });
    e.target.value = "";
  };

  const validateConfig = () => {
    const issues: string[] = [];
    const host = toStr(CONFIG.spSiteHostname);
    const path = toStr(CONFIG.spSitePath);
    const sharepointHostPattern = /(sharepoint\.(com|cn|de|mil|us)|sharepoint-df\.com)/i;
    if (!sharepointHostPattern.test(host)) {
      issues.push(`spSiteHostname '${host}' does not look like a SharePoint Online host.`);
    }
    if (!/^\/(sites|teams)\//i.test(path)) {
      issues.push(`spSitePath '${path}' should start with '/sites/' or '/teams/'.`);
    }
    if (!toStr(CONFIG.libraryFolderPath)) {
      issues.push("libraryFolderPath is empty - images will be stored at the library root/YYYY/MM.");
    }
    if (!toStr(CONFIG.graphBaseUrl)) {
      issues.push("graphBaseUrl is empty - defaulting to https://graph.microsoft.com");
    }
    if (!toStr(CONFIG.graphVersion)) {
      issues.push("graphVersion is empty - defaulting to v1.0");
    }
    if (!toStr(CONFIG.redirectUri)) {
      issues.push("redirectUri is empty - add your SPA redirect URI in Entra app registration.");
    }
    return issues;
  };

  const submit = async () => {
    if (!account) return alert("Please sign in first.");
    if (!files.length) return alert("Select at least one photo.");
    if (!msalReady) return alert("Auth is still initializing. Try again in a moment.");

    const cfgIssues = validateConfig();
    if (cfgIssues.length) { if (!confirm(`Config looks unusual:\\n- ${cfgIssues.join("\\n- ")}\\n\\nContinue anyway?`)) return; }

    setBusy(true); setLog([]);
    try {
      addLog("Probing network reachability...");
      const [gProbe, spProbe] = await Promise.all([ probeGraphReachable(), probeSharePointHost(CONFIG.spSiteHostname) ]);
      if (!gProbe.ok) addLog(`[Probe] Graph unreachable: ${gProbe.error || gProbe.status} (meta: ${gProbe.metaUrl})`);
      if (!spProbe.ok) addLog(`[Probe] SharePoint host unreachable: ${spProbe.error}`);

      addLog("Resolving site, list, library & columns...");
      const siteId = resolved.siteId || (await getSiteId());
      const listId = resolved.listId || (await resolveListId(siteId, CONFIG.spListIdOrName));
      const driveId = resolved.driveId || (await resolveDriveId(siteId, CONFIG.spLibraryDriveIdOrName));
      const { byLower, typeByName, choicesByName, metaByName } = await getColumns(siteId, listId);

      // Decide internal names (case-insensitive)
      const titleName = chooseField(byLower, metaByName, ["Title"], { requireWritable: true }) || "Title";
      const villageName = chooseField(byLower, metaByName, ["Village"], { requireWritable: true }) || null;
      const notesName = chooseField(byLower, metaByName, ["Notes","Note","Description"], { requireWritable: true }) || null;
      const capturedOnName = chooseField(byLower, metaByName, ["CapturedOn","Captured On","Captured_On"], { requireWritable: true }) || null;
      const photoFieldName = chooseField(byLower, metaByName, ["PhotoUrl","PhotoUrL","Photo URL","Photo_Url"], { requireWritable: true }) || null;
      const categoryOptionsName = chooseField(byLower, metaByName, ["CategoryOptions","Category Options","Category"]) || null;
      setResolved({ siteId, listId, driveId });

      // Capture Village choices for the dropdown UI when available
      const availableVillageChoices = villageName ? (choicesByName[villageName] ?? []) : [];
      if (availableVillageChoices.length) {
        setVillageChoices(availableVillageChoices);
      } else {
        setVillageChoices(null);
      }
      const availableTitleChoices = categoryOptionsName ? (choicesByName[categoryOptionsName] ?? []) : [];
      if (availableTitleChoices.length) {
        setTitleChoices(availableTitleChoices);
      } else {
        setTitleChoices(null);
      }

      if (!photoFieldName) throw new Error("Could not find a 'PhotoUrl'/'PhotoUrL' column on the list. Check the internal name in List Settings → Columns.");

      const villageMode = availableVillageChoices.length ? "choice" : "text/other";
      addLog(`Columns resolved → title: ${titleName}; village: ${villageName ?? "<none>"} (${villageMode}); notes: ${notesName ?? "<none>"}; capturedOn: ${capturedOnName ?? "<none>"}; photo: ${photoFieldName} (${typeByName[photoFieldName]})`);

      addLog(`Uploading ${files.length} file(s)...`);
      const uploadedItems: any[] = [];
      for (const f of files) {
        addLog(`Uploading: ${f.name} (${Math.round(f.size / 1024)} KB)`);
        const driveItem = await uploadOne(driveId!, CONFIG.libraryFolderPath, f, village || "");
        uploadedItems.push(driveItem);
        addLog(`Uploaded ✔ → ${driveItem?.webUrl ?? "(no url)"}`);
      }

      const photoUrls: string[] = uploadedItems.map((i) => i?.webUrl).filter(Boolean);
      const primaryUrl = photoUrls[0] || "";
      const extraUrls = photoUrls.slice(1);

      // Notes stay as-entered; photo URLs tracked solely in the photo column
      const augmentedNotes = notes || "";

      const isoDate = formatDateTimeForSharePoint(capturedOn, NZ_TIME_ZONE);

      const fields: any = {};
      fields[titleName] = title || (files[0]?.name || "Photo Report");
      if (villageName) fields[villageName] = village || "";
      if (notesName) fields[notesName] = augmentedNotes; else if (extraUrls.length) addLog("Warning: Notes field not found; additional URLs were not stored in list notes.");
      if (capturedOnName) fields[capturedOnName] = isoDate;
      fields[photoFieldName] = formatPhotoValue(typeByName, photoFieldName, [primaryUrl, ...extraUrls]);

      addLog("Creating list item with resolved field names...");
      const created = await createListItem(siteId!, listId!, fields);
      const newItemId = created?.id || created?.listItem?.id || "(unknown)";
      addLog(`List item created ✔ (id: ${newItemId})`);

      addLog("All done. ✨");
      alert("Submitted successfully.");
      setTitle(""); setVillage(""); setNotes(""); setFiles([]);
      setPreviews((prev) => {
        prev.forEach((item) => URL.revokeObjectURL(item.url));
        return [];
      });
      setCapturedOn(() => formatDateTimeLocal(new Date(), NZ_TIME_ZONE));
      setDiagnosticsOpen(false);
      if (inputRef.current) inputRef.current.value = "";
      scheduleMobileRefresh();
    } catch (e: any) {
      console.error(e);
      addLog(`Error: ${e.message}`);
      alert(`Something went wrong: ${e.message}`);
    } finally { setBusy(false); }
  };

  // ── DIAGNOSTICS (unchanged except minor notes) ─────────────────────────────
  const [diagBusy, setDiagBusy] = useState(false);
  const [diag, setDiag] = useState<any>({});

  const runDiagnostics = async () => {
    setDiagnosticsOpen(true);
    setDiagBusy(true); setDiag({}); setLog([]);
    try {
      addLog("[Diagnostics] Checking configuration...");
      const cfgIssues: string[] = [];
      ["tenantId","clientId","spSiteHostname","spSitePath","spListIdOrName","spLibraryDriveIdOrName","libraryFolderPath","graphBaseUrl","graphVersion","redirectUri"].forEach((k) => ((CONFIG as any)[k] == null || (CONFIG as any)[k] === "") && cfgIssues.push(`Missing/empty CONFIG.${k}`));
      const cfgWarnings = (() => {
        const warnings: string[] = [];
        const host = toStr(CONFIG.spSiteHostname);
        const path = toStr(CONFIG.spSitePath);
        const hostPattern = /(sharepoint\.(com|cn|de|mil|us)|sharepoint-df\.com)/i;
        if (!hostPattern.test(host)) {
          warnings.push(`spSiteHostname '${host}' does not look like a SharePoint Online host.`);
        }
        if (!/^\/(sites|teams)\//i.test(path)) {
          warnings.push(`spSitePath '${path}' should start with '/sites/' or '/teams/'.`);
        }
        if (!toStr(CONFIG.libraryFolderPath)) {
          warnings.push("libraryFolderPath is empty - images will be stored at the library root/YYYY/MM.");
        }
        return warnings;
      })();
      setDiag((d: any) => ({ ...d, configOk: cfgIssues.length === 0, cfgIssues, cfgWarnings }));

      addLog("[Diagnostics] Ensuring MSAL initialized...");
      setDiag((d: any) => ({ ...d, msalReady }));

      addLog("[Diagnostics] Detecting iframe environment & popup-only policy...");
      setDiag((d: any) => ({ ...d, isInIframe, allowRedirectInIframe: false, authFlow: "popup-only" }));

      addLog("[Diagnostics] Redirect URI vs origin (null-safe)...");
      const redirect = toStr(CONFIG.redirectUri); const origin = toStr((window as any)?.location?.origin);
      const redirectMatches = redirect.replace(/\/$/, "") === origin.replace(/\/$/, "");
      setDiag((d: any) => ({ ...d, redirectUri: redirect, origin, redirectMatches }));

      addLog("[Diagnostics] Probing network reachability (Graph & SharePoint)...");
      const [gProbe, spProbe] = await Promise.all([ probeGraphReachable(), probeSharePointHost(CONFIG.spSiteHostname) ]);
      setDiag((d: any) => ({ ...d, graphBaseUrl: toStr(CONFIG.graphBaseUrl), graphVersion: toStr(CONFIG.graphVersion), graphReachable: gProbe, sharePointReachable: spProbe }));

      addLog("[Diagnostics] Acquiring token (silent→popup)...");
      let tokenOk = false; try { const silent = await msal.acquireTokenSilent({ account: account || msal.getAllAccounts()[0], scopes: GRAPH_SCOPES }); tokenOk = !!silent?.accessToken; }
      catch { const interactive = await msal.acquireTokenPopup({ scopes: GRAPH_SCOPES }); tokenOk = !!interactive?.accessToken; }
      setDiag((d: any) => ({ ...d, tokenReceived: tokenOk }));

      addLog("[Diagnostics] Resolve site/list/drive & columns...]");
      const siteId = (await (async () => { try { const direct = await sgraph!.get(`/sites/${toStr(CONFIG.spSiteHostname)}:${toStr(CONFIG.spSitePath)}`); return direct.id; } catch { const term = lastSegment || toStr(CONFIG.spSitePath) || toStr(CONFIG.spSiteHostname); const res = await sgraph!.get(`/sites?search=${encodeURIComponent(term)}`); return res?.value?.[0]?.id; } })());
      const listId = await resolveListId(siteId, CONFIG.spListIdOrName);
      const driveId = await resolveDriveId(siteId, CONFIG.spLibraryDriveIdOrName);
      const { byLower, typeByName, choicesByName, metaByName } = await getColumns(siteId, listId);
      const titleName = chooseField(byLower, metaByName, ["Title"], { requireWritable: true }) || "Title";
      const villageName = chooseField(byLower, metaByName, ["Village"], { requireWritable: true }) || null;
      const notesName = chooseField(byLower, metaByName, ["Notes","Note","Description"], { requireWritable: true }) || null;
      const capturedOnName = chooseField(byLower, metaByName, ["CapturedOn","Captured On","Captured_On"], { requireWritable: true }) || null;
      const photoFieldName = chooseField(byLower, metaByName, ["PhotoUrl","PhotoUrL","Photo URL","Photo_Url"], { requireWritable: true }) || null;
      const categoryOptionsName = chooseField(byLower, metaByName, ["CategoryOptions","Category Options","Category"]) || null;
      const diagVillageChoices = villageName ? (choicesByName[villageName] ?? []) : [];
      const diagTitleChoices = categoryOptionsName ? (choicesByName[categoryOptionsName] ?? []) : [];
      setVillageChoices(diagVillageChoices.length ? diagVillageChoices : null);
      setTitleChoices(diagTitleChoices.length ? diagTitleChoices : null);
      setDiag((d: any) => ({
        ...d,
        siteId,
        listId,
        driveId,
        resolvedFields: {
          titleName,
          villageName,
          notesName,
          capturedOnName,
          photoFieldName,
          photoType: photoFieldName ? typeByName[photoFieldName] : "<missing>",
          categoryOptionsName,
          titleChoices: diagTitleChoices,
          villageChoices: diagVillageChoices,
        },
      }));

      addLog("[Diagnostics] Graph client smoke test: /me?$select=id,displayName ...");
      const me = await sgraph!.get(`/me?$select=id,displayName`);
      setDiag((d: any) => ({ ...d, meSelect: { id: me?.id, displayName: me?.displayName } }));

      // Path builder self-tests (no network)
      const rootChildrenPath = childrenEndpoint(driveId, "");
      const nestedChildrenPath = childrenEndpoint(driveId, "a/b");
      setDiag((d: any) => ({ ...d, pathTests: { rootChildrenPath, nestedChildrenPath } }));

      // Sample payload preview for diagnostics
      const sampleUrls = ["https://contoso/img-a.jpg","https://contoso/img-b.jpg","https://contoso/img-c.jpg"];
      const sampleNotes = "Example note";
      const photoValuePreview = photoFieldName ? formatPhotoValue(typeByName, photoFieldName, sampleUrls) : undefined;
      setDiag((d: any) => ({ ...d, sampleFields: { [titleName]: "TEST – Sample Multi Photo", ...(villageName ? { [villageName]: "SampleVillage" } : {}), ...(notesName ? { [notesName]: sampleNotes } : {}), ...(capturedOnName ? { [capturedOnName]: formatDateTimeForSharePoint(formatDateTimeLocal(new Date(), NZ_TIME_ZONE), NZ_TIME_ZONE) } : {}), ...(photoFieldName ? { [photoFieldName]: photoValuePreview } : {}) } }));

      addLog("[Diagnostics] Completed successfully ✅");
    } catch (e: any) {
      console.error(e);
      addLog(`[Diagnostics] Failed: ${e.message}`);
      setDiag((d: any) => ({ ...d, error: e.message }));
    } finally { setDiagBusy(false); }
  };

  // ── RENDER ────────────────────────────────────────────────────────────────
  return (
    <div className="app-shell">
      <div className="app-layout">
        <header className="app-header">
          <div className="app-brand">
            <img src={metlifecareLogo} alt="Metlifecare logo" className="app-logo" />
            <div className="app-heading">
              <span className="app-tagline"></span>
              <h1 className="app-title">Village Recon Tool</h1>
              <p className="app-subtitle">Record key observations in the field and push them to SharePoint in seconds.</p>
            </div>
          </div>
          <div className="auth-block">
            {!msalReady && <span className="auth-status">Initializing authentication…</span>}
            {account ? (
              <>
                <span className="auth-identity">{account.username}</span>
                <Button variant="ghost" onClick={signOut} disabled={!msalReady}>Sign out</Button>
              </>
            ) : (
              <Button variant="primary" onClick={signIn} disabled={!msalReady}>Sign in with Entra</Button>
            )}
          </div>
        </header>

        <main className="app-main">
          <section className="panel panel-form">
            <div className="panel-header">
              <div>
                <h2 className="panel-title">Create a new report</h2>
                <p className="panel-subtitle">Add a short summary, choose the village, and drop in your latest photos.</p>
              </div>
              {files.length > 0 ? (
                <span className="panel-badge">{files.length} photo{files.length > 1 ? "s" : ""}</span>
              ) : null}
            </div>

            <div className="form-grid">
              <Field label="Title" controlId="field-title" hint="Give the upload a clear, searchable name.">
                {Array.isArray(titleChoices) && titleChoices.length > 0 ? (
                  <select className="form-control" value={title} onChange={(e) => setTitle(e.target.value)}>
                    <option value="">Select a category.</option>
                    {titleChoices.map((opt) => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                ) : (
                  <input type="text" className="form-control" value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Unknown hardware" />
                )}
              </Field>

              {Array.isArray(villageChoices) && villageChoices.length > 0 ? (
                <Field label="Village" controlId="field-village">
                  <select className="form-control" value={village} onChange={(e) => setVillage(e.target.value)}>
                    <option value="">Select a village…</option>
                    {villageChoices.map((opt) => (
                      <option key={opt} value={opt}>{opt}</option>
                    ))}
                  </select>
                </Field>
              ) : (
                <Field label="Village" controlId="field-village" hint="Please sign in first to populate the village list.">
                  <input type="text" className="form-control" value={village} onChange={(e) => setVillage(e.target.value)} placeholder="Please Authenticate" />
                </Field>
              )}

              <Field label="Captured on" controlId="field-captured" hint="Store the local date and time for this report.">
                <input type="datetime-local" className="form-control" value={capturedOn} onChange={(e) => setCapturedOn(e.target.value)} />
              </Field>

              <Field label="Notes" controlId="field-notes" hint="Add context, observations, or follow-up actions." full>
                <textarea className="form-control" rows={4} value={notes} onChange={(e) => setNotes(e.target.value)} placeholder="Add context for the photos." />
              </Field>

              <Field label="Photos" controlId="field-photos" hint="Upload one or more images or choose from your library — we will create the SharePoint folder for you." full>
                <input
                  ref={inputRef}
                  type="file"
                  multiple
                  accept="image/*"
                  onChange={onFilesChanged}
                  className="form-control file-input"
                />
                {previews.length > 0 && (
                  <div className="photo-grid">
                    {previews.map((preview, index) => (
                      <img key={preview.key} src={preview.url} alt={`preview-${index + 1}`} className="photo-thumb" />
                    ))}
                  </div>
                )}
              </Field>
            </div>

            <Button variant="primary" onClick={submit} disabled={!account || busy || files.length === 0 || !msalReady || !sgraph}>
              {busy ? "Submitting…" : "Submit report"}
            </Button>
          </section>

          <section className="panel">
            <div className="panel-header">
              <div>
                <h3 className="panel-title">Activity</h3>
                <p className="panel-subtitle">Live log of uploads, diagnostics, and Microsoft Graph calls.</p>
              </div>
            </div>
            <div className="activity-log">{log.length ? log.join("\n") : "No activity yet."}</div>
          </section>

          <section className={`panel diagnostics-panel${diagnosticsOpen ? " diagnostics-panel--open" : ""}`}>
            <div className="panel-header diagnostics-header">
              <button
                type="button"
                className="diagnostics-toggle"
                onClick={toggleDiagnostics}
                aria-expanded={diagnosticsOpen}
                aria-controls="diagnostics-content"
              >
                <span className="diagnostics-toggle-icon" aria-hidden="true" />
                <span className="diagnostics-text">
                  <span className="panel-title">Diagnostics</span>
                  <span className="panel-subtitle">Run this checklist if something looks off before escalating.</span>
                </span>
              </button>
              <Button variant="secondary" onClick={runDiagnostics} disabled={!account || diagBusy || !msalReady || !sgraph}>
                {diagBusy ? "Running…" : "Run diagnostics"}
              </Button>
            </div>
            <div
              id="diagnostics-content"
              className="diagnostics-content"
              hidden={!diagnosticsOpen}
              aria-hidden={!diagnosticsOpen}
            >
              <ul className="diagnostics-list">
                <li>Config present for tenant/site/list/library (with host/path warnings)</li>
                <li>MSAL initialises and enforces popup-only flows</li>
                <li>Redirect URI matches the current origin</li>
                <li>Network probes confirm Graph and SharePoint host reachability</li>
                <li>Token acquisition via MSAL (silent ⇢ popup fallback)</li>
                <li>Resolve site, list, and library drive identifiers</li>
                <li>Resolve internal field names (Title, Village, Notes, CapturedOn, PhotoUrl)</li>
                <li>Detect photo column type and format payload accordingly</li>
                <li>Append additional photo URLs into Notes for readability</li>
                <li>Create/validate library folder path <code>{toStr(CONFIG.libraryFolderPath) || "<library root>"}/YYYY/MM</code></li>
                <li>Graph base/version compose correctly and `$metadata` endpoint responds</li>
                <li>Graph client smoke test: `/me?$select=id,displayName` returns user identity</li>
                <li>Path builder tests ensure no `root::/children` usage</li>
              </ul>
              <div className="diagnostics-json">{Object.keys(diag).length === 0 ? "No diagnostics have been run yet." : JSON.stringify(diag, null, 2)}</div>
            </div>
          </section>

          <footer className="app-footer">
            Creates a single SharePoint list item per submission and uploads multiple photos to the library. If the Village column is configured as Choice, a dropdown appears automatically; otherwise, free-text is used.
          </footer>
        </main>
      </div>
    </div>
  );
}


