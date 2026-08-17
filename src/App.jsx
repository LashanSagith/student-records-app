import React from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { GraduationCap, ShieldCheck, Cloud, History } from "lucide-react";
import { loginRequest } from "./authConfig";
import StudentRecords from "./StudentRecords";

const THEME = {
  bg: "#F0F6FF",
  surface: "#FFFFFF",
  ink: "#144FB3",
  inkMuted: "#97B2E8",
  border: "#DADCE0",
  primary: "#3D7DFF",
  primaryDark: "#083BA1",
  primarySoft: "#A4BADE",
};

function LoginScreen() {
  const { instance } = useMsal();
  const [error, setError] = React.useState(null);
  const [busy, setBusy] = React.useState(false);

  async function signIn() {
    setError(null);
    setBusy(true);
    try {
      await instance.loginPopup(loginRequest);
    } catch (e) {
      setError(e?.errorMessage || e?.message || "Sign-in failed. Please try again.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div style={{ background: THEME.bg, fontFamily: "'Inter', sans-serif" }} className="min-h-screen flex items-center justify-center p-6">
      <div style={{ background: THEME.surface, border: `1px solid ${THEME.border}` }} className="w-full max-w-sm rounded-xl p-7 text-center">
        <div style={{ background: THEME.primarySoft }} className="w-12 h-12 rounded-full flex items-center justify-center mx-auto mb-4">
          <GraduationCap size={22} color={THEME.primaryDark} />
        </div>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", color: THEME.ink }} className="text-lg font-semibold">Students Master Database</div>
        <div style={{ color: THEME.inkMuted }} className="text-xs mt-1.5 mb-6">Sign in with your organisation Microsoft account to continue.</div>

        <button
          type="button"
          onClick={signIn}
          disabled={busy}
          style={{ background: THEME.primary, color: "#fff" }}
          className="w-full flex items-center justify-center gap-2 text-sm font-semibold px-4 py-2.5 rounded-lg hover:opacity-90 disabled:opacity-60"
        >
          {busy ? "Signing in…" : "Sign in with Microsoft"}
        </button>

        {error && (
          <div style={{ color: "#B23A48" }} className="text-xs mt-3 text-left">{error}</div>
        )}

        <div style={{ borderTop: `1px solid ${THEME.border}` }} className="mt-6 pt-5 flex flex-col gap-3 text-left">
          <div className="flex items-start gap-2.5">
            <ShieldCheck size={15} color={THEME.primary} className="mt-0.5 shrink-0" />
            <span style={{ color: THEME.inkMuted }} className="text-xs">Only people in your organisation can sign in.</span>
          </div>
          <div className="flex items-start gap-2.5">
            <Cloud size={15} color={THEME.primary} className="mt-0.5 shrink-0" />
            <span style={{ color: THEME.inkMuted }} className="text-xs">Data is backed up to your own OneDrive automatically.</span>
          </div>
          <div className="flex items-start gap-2.5">
            <History size={15} color={THEME.primary} className="mt-0.5 shrink-0" />
            <span style={{ color: THEME.inkMuted }} className="text-xs">Every add, edit and delete is written to an edit-history log.</span>
          </div>
        </div>
      </div>
    </div>
  );
}

export default function App() {
  const isAuthenticated = useIsAuthenticated();
  const { inProgress } = useMsal();

  if (inProgress !== "none") {
    return (
      <div style={{ background: THEME.bg }} className="min-h-screen flex items-center justify-center">
        <div style={{ color: THEME.inkMuted }} className="text-sm">Signing in…</div>
      </div>
    );
  }

  if (!isAuthenticated) return <LoginScreen />;

  return <StudentRecords />;
}
