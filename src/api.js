const API_BASE = (import.meta.env.VITE_API_BASE || "/api").replace(/\/$/, "");

function token() {
  try { return window.localStorage.getItem("rabet-token") || ""; } catch { return ""; }
}

async function request(path, options = {}) {
  const auth = token();
  const response = await fetch(`${API_BASE}${path}`, {
    ...options,
    headers: { "Content-Type": "application/json", ...(auth ? { Authorization: `Bearer ${auth}` } : {}), ...(options.headers || {}) },
  });
  if (!response.ok) throw new Error(`API ${response.status}: ${await response.text()}`);
  return response.status === 204 ? null : response.json();
}

export async function loginRemote(username, password, server) {
  const result = await request("/auth/login", { method: "POST", body: JSON.stringify({ username, password, server }) });
  try { window.localStorage.setItem("rabet-token", result.token); } catch { /* التخزين المحلي اختياري */ }
  return result.user;
}

export function logoutRemote() {
  try { window.localStorage.removeItem("rabet-token"); } catch { /* لا شيء */ }
}

export async function loadRemoteState() {
  if (!token()) return null;
  try {
    const result = await request("/state");
    return result?.state || null;
  } catch (error) {
    console.warn("Rabet: backend غير متاح، سيُستخدم التخزين المحلي مؤقتاً.", error);
    return null;
  }
}

export async function saveRemoteState(state) {
  if (!token()) return false;
  try {
    await request("/state", { method: "PUT", body: JSON.stringify({ state }) });
    return true;
  } catch (error) {
    console.warn("Rabet: تعذرت مزامنة البيانات مع الخادم.", error);
    return false;
  }
}

export async function healthCheck() {
  try { return await request("/health"); } catch { return null; }
}
