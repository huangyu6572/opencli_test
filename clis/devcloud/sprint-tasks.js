import { cli, Strategy } from '@jackwener/opencli/registry';

cli({
  site: 'devcloud',
  name: 'sprint-tasks',
  description: 'List all tasks in a Huawei DevCloud sprint (title, status, assignee, priority)',
  strategy: Strategy.COOKIE,
  browser: true,
  access: 'read',
  args: [
    {
      name: 'url',
      required: true,
      help: 'Full DevCloud sprint URL, e.g. https://hn.devcloud.huaweicloud.com/projectman/scrum/<projectId>/task/sprint/<sprintId>/list',
    },
    {
      name: 'limit',
      default: 200,
      help: 'Maximum number of tasks to return',
    },
  ],
  columns: ['id', 'title', 'status', 'assignee', 'priority', 'type'],

  func: async (page, kwargs) => {
    const args = kwargs;
    const { url, limit } = args;

    // Parse projectId and sprintId from the URL
    const urlMatch = url.match(
      /\/scrum\/([a-f0-9]+)\/task\/sprint\/(\d+)/
    );
    if (!urlMatch) {
      throw new Error(
        `Invalid DevCloud sprint URL: expected /scrum/<projectId>/task/sprint/<sprintId>/ in "${url}"`
      );
    }
    const projectId = urlMatch[1];
    const sprintId = urlMatch[2];

    // Navigate to the sprint page to establish a valid authenticated session
    await page.goto(url);

    const apiBase = new URL(url).origin;

    // Attempt 1: v4 GET endpoint
    let tasks = null;
    const v4Url = `${apiBase}/v4/projects/${projectId}/sprints/${sprintId}/work-items?limit=${limit}&offset=0`;
    try {
      const resp = await page.evaluate(
        async ({ u, limit }) => {
          const r = await fetch(u, { credentials: "include" });
          if (!r.ok) return null;
          return r.json();
        },
        { u: v4Url, limit }
      );
      if (resp && (resp.work_items || resp.items || resp.data)) {
        const raw = resp.work_items || resp.items || resp.data || [];
        tasks = raw.slice(0, limit);
      }
    } catch (_) {
      // fall through
    }

    // Attempt 2: v2 GET endpoint
    if (!tasks) {
      const v2Url = `${apiBase}/v2/projects/${projectId}/sprints/${sprintId}/work-items?limit=${limit}&offset=0`;
      try {
        const resp = await page.evaluate(
          async ({ u }) => {
            const r = await fetch(u, { credentials: "include" });
            if (!r.ok) return null;
            return r.json();
          },
          { u: v2Url }
        );
        if (resp && (resp.work_items || resp.items || resp.data)) {
          const raw = resp.work_items || resp.items || resp.data || [];
          tasks = raw.slice(0, limit);
        }
      } catch (_) {
        // fall through
      }
    }

    // Attempt 3: POST filter endpoint
    if (!tasks) {
      const filterUrl = `${apiBase}/v4/projects/${projectId}/work-items/filter`;
      try {
        const resp = await page.evaluate(
          async ({ u, sprintId, limit }) => {
            const r = await fetch(u, {
              method: "POST",
              credentials: "include",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                sprint_id: sprintId,
                limit,
                offset: 0,
              }),
            });
            if (!r.ok) return null;
            return r.json();
          },
          { u: filterUrl, sprintId, limit }
        );
        if (resp && (resp.work_items || resp.items || resp.data)) {
          const raw = resp.work_items || resp.items || resp.data || [];
          tasks = raw.slice(0, limit);
        }
      } catch (_) {
        // fall through
      }
    }

    // Attempt 4: Intercept via network — wait for page to load and grab XHR
    if (!tasks) {
      const intercepted = await page.evaluate(() => {
        return window.__opencli_devcloud_tasks__ || null;
      });
      if (intercepted) tasks = intercepted.slice(0, limit);
    }

    // Attempt 5: DOM scraping fallback
    if (!tasks) {
      tasks = await page.evaluate(() => {
        const rows = Array.from(
          document.querySelectorAll(
            ".work-item-row, .task-item, [class*='work-item'], [class*='task-row'], tr[data-id]"
          )
        );
        return rows.map((row) => {
          const titleEl =
            row.querySelector("[class*='title'], [class*='name'], td:nth-child(2)") ||
            row.querySelector("a");
          const statusEl =
            row.querySelector("[class*='status'], [class*='state'], td:nth-child(3)");
          const assigneeEl =
            row.querySelector("[class*='assignee'], [class*='owner'], td:nth-child(4)");
          const priorityEl =
            row.querySelector("[class*='priority'], td:nth-child(5)");
          return {
            title: titleEl ? titleEl.textContent.trim() : "",
            status: statusEl ? statusEl.textContent.trim() : "",
            assignee: assigneeEl ? assigneeEl.textContent.trim() : "",
            priority: priorityEl ? priorityEl.textContent.trim() : "",
          };
        }).filter((t) => t.title);
      });
    }

    if (!tasks || tasks.length === 0) {
      throw new Error(
        "No tasks found. Make sure Chrome is open and you are logged into DevCloud, then try again."
      );
    }

    // Normalize field names across different API versions
    return tasks.map((t) => ({
      id: t.id || t.work_item_id || t.number || "",
      title: t.title || t.name || t.subject || "",
      status:
        t.status ||
        (t.status_detail && t.status_detail.name) ||
        (t.tracker && t.tracker.name) ||
        t.state ||
        "",
      assignee:
        t.assignee ||
        (t.assigned_user && t.assigned_user.name) ||
        (t.developer && t.developer.name) ||
        "",
      priority:
        t.priority ||
        (t.priority_detail && t.priority_detail.name) ||
        "",
      type:
        t.type ||
        (t.tracker && t.tracker.name) ||
        (t.work_item_type && t.work_item_type.name) ||
        "",
    }));
  },
});
