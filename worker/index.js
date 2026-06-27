/**
 * Cloudflare Worker — handles /api/trigger and /api/status
 * Triggers the GitHub Actions workflow and reports its run status back to the dashboard.
 *
 * Requires secrets (set via Cloudflare dashboard or `wrangler secret put`):
 *   GH_PAT          — GitHub PAT with `repo` + `workflow` scopes
 *   GH_REPO         — e.g. "thakkars/rotational_invest"
 *   GH_WORKFLOW_ID  — workflow file name e.g. "daily-refresh.yml"
 */

const REPO = "thakkars/rotational_invest";
const WORKFLOW_ID = "daily-refresh.yml";

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    // CORS + JSON helpers
    const json = (data, status = 200) =>
      new Response(JSON.stringify(data), {
        status,
        headers: {
          "Content-Type": "application/json",
          "Access-Control-Allow-Origin": "*",
          "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
          "Access-Control-Allow-Headers": "Content-Type",
        },
      });

    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204 });
    }

    if (!env.GH_PAT) {
      return json({ ok: false, error: "GH_PAT secret not configured" }, 500);
    }

    // --- /api/trigger : start the workflow ---
    if (url.pathname === "/api/trigger" && request.method === "POST") {
      try {
        const r = await fetch(
          `https://api.github.com/repos/${REPO}/actions/workflows/${WORKFLOW_ID}/dispatches`,
          {
            method: "POST",
            headers: {
              Authorization: `Bearer ${env.GH_PAT}`,
              Accept: "application/vnd.github+json",
              "X-GitHub-Api-Version": "2022-11-28",
              "User-Agent": "rotation-dashboard",
            },
            body: JSON.stringify({ ref: "main" }),
          }
        );
        if (!r.ok) {
          const t = await r.text();
          return json({ ok: false, error: `GitHub ${r.status}: ${t}` }, 502);
        }
        return json({ ok: true });
      } catch (e) {
        return json({ ok: false, error: String(e) }, 500);
      }
    }

    // --- /api/status : latest run status ---
    if (url.pathname === "/api/status" && request.method === "GET") {
      try {
        const r = await fetch(
          `https://api.github.com/repos/${REPO}/actions/runs?per_page=1`,
          {
            headers: {
              Authorization: `Bearer ${env.GH_PAT}`,
              Accept: "application/vnd.github+json",
              "X-GitHub-Api-Version": "2022-11-28",
              "User-Agent": "rotation-dashboard",
            },
          }
        );
        const d = await r.json();
        const run = d.workflow_runs && d.workflow_runs[0];
        if (!run) return json({ status: "unknown" });
        return json({
          status: run.status,            // queued | in_progress | completed
          conclusion: run.conclusion,    // success | failure | null
          created_at: run.created_at,
          html_url: run.html_url,
        });
      } catch (e) {
        return json({ status: "unknown", error: String(e) });
      }
    }

    // Fallback — let the assets handler serve static files (index.html, xlsx, etc.)
    return env.ASSETS.fetch(request);
  },
};
