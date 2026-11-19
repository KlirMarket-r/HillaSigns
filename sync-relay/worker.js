// Hilla's Signs: sync relay (Cloudflare Worker)
//
// The app's סנכרון מחירים button POSTs here. This Worker asks GitHub to run the sync workflow
// (sync.yml, "workflow_dispatch") with a GitHub token that lives only in the Worker's secrets.
// The app itself never holds a token. Anything else (GET, other origins) gets a plain answer.
//
// Secret to set in the Worker: GH_TOKEN  (a fine-grained GitHub token, repo KlirMarket-r/HillaSigns,
//                                         permission "Actions: Read and write", nothing else)

const REPO = "KlirMarket-r/HillaSigns";
const WORKFLOW = "sync.yml";
const ALLOWED_ORIGIN = "https://klirmarket-r.github.io";

export default {
  async fetch(request, env) {
    const cors = {
      "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
      "Access-Control-Allow-Methods": "POST, OPTIONS",
      "Access-Control-Max-Age": "86400",
    };
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors });
    if (request.method !== "POST") {
      return new Response("Hilla's Signs sync relay. POST from the app starts a sync.", { headers: cors });
    }
    const r = await fetch(`https://api.github.com/repos/${REPO}/actions/workflows/${WORKFLOW}/dispatches`, {
      method: "POST",
      headers: {
        "Authorization": "Bearer " + env.GH_TOKEN,
        "Accept": "application/vnd.github+json",
        "User-Agent": "hillasigns-sync-relay",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ ref: "main" }),
    });
    const ok = r.status === 204;
    return new Response(JSON.stringify({ ok, github: r.status }), {
      status: ok ? 200 : 502,
      headers: { ...cors, "Content-Type": "application/json" },
    });
  },
};
