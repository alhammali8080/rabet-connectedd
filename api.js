const API_BASE = (import.meta.env.VITE_API_BASE || "/api").replace(/\/$/, "");

async function request(path, options = {}) {
  const response = await fetch(`${API_BASE}${path}`, {
    ...options,
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
  });
  if (!response.ok) throw new Error(`API ${response.status}: ${await response.text()}`);
  return response.status === 204 ? null : response.json();
}

export async function loadRemoteState() {
  try {
    const result = await request("/state");
    return result?.state || null;
  } catch (error) {
    console.warn("Rabet: backend غير متاح، سيُستخدم التخزين المحلي مؤقتاً.", error);
    return null;
  }
}

export async function saveRemoteState(state) {
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
