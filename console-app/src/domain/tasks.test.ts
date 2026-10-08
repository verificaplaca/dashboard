import test from "node:test";
import assert from "node:assert/strict";
import type { SupabaseClient } from "@supabase/supabase-js";
import {
  canManageTask,
  isTaskOverdue,
  taskInput,
  taskToday,
  validTaskDate,
  type LiveTask,
  type TaskMember,
} from "./tasks";
import { taskMutation } from "./task-api";
const members: TaskMember[] = [
  {
    user_id: "operator",
    display_name: "Operador",
    role: "manager",
    active: true,
  },
  { user_id: "member", display_name: "Membro", role: "member", active: true },
  {
    user_id: "inactive",
    display_name: "Inativo",
    role: "member",
    active: false,
  },
];
const task: LiveTask = {
  id: "task-id",
  title: "Revisar campanha",
  description: "Critérios",
  priority: "normal",
  status: "todo",
  assignee_id: "member",
  created_by: "operator",
  due_date: "2026-10-08",
  completed_at: null,
  archived_at: null,
  version: 3,
  created_at: "2026-10-07T12:00:00Z",
  updated_at: "2026-10-07T12:00:00Z",
};
test("tasks use São Paulo date boundaries and do not mark done or archived tasks overdue", () => {
  assert.equal(taskToday(new Date("2026-10-09T02:59:59Z")), "2026-10-08");
  assert.equal(taskToday(new Date("2026-10-09T03:00:00Z")), "2026-10-09");
  assert.equal(isTaskOverdue(task, "2026-10-08"), false);
  assert.equal(isTaskOverdue(task, "2026-10-09"), true);
  assert.equal(isTaskOverdue({ ...task, status: "done" }, "2026-10-09"), false);
  assert.equal(
    isTaskOverdue(
      { ...task, archived_at: "2026-10-07T12:00:00Z" },
      "2026-10-09",
    ),
    false,
  );
  assert.equal(isTaskOverdue({ ...task, due_date: null }, "2026-10-09"), false);
});
test("task input validates dates, active assignees, content boundaries and prototype keys", () => {
  const values = {
    title: "  Revisar campanha  ",
    description: " Contexto ",
    priority: "high",
    assignee_id: "member",
    due_date: "2026-10-08",
  };
  assert.deepEqual(taskInput(values, members, false), {
    ...values,
    title: "Revisar campanha",
    description: "Contexto",
  });
  for (const due_date of [
    "2026-02-30",
    "2026-99-01",
    "01/10/2026",
    "1999-12-31",
    "2201-01-01",
  ])
    assert.throws(() => taskInput({ ...values, due_date }, members, false));
  assert.equal(validTaskDate("2028-02-29"), true);
  assert.equal(validTaskDate("2026-02-29"), false);
  for (const priority of ["urgent", "constructor", "__proto__"])
    assert.throws(() => taskInput({ ...values, priority }, members, false));
  for (const assignee_id of ["inactive", "unknown", ""])
    assert.throws(() => taskInput({ ...values, assignee_id }, members, false));
  assert.equal(
    taskInput({ ...values, assignee_id: "", due_date: "" }, members, true)
      .assignee_id,
    null,
  );
  assert.throws(() => taskInput({ ...values, title: " " }, members, true));
  assert.throws(() =>
    taskInput({ ...values, description: "x".repeat(6001) }, members, true),
  );
});
test("detail management is restricted to the active creator or manager, not just assignee", () => {
  assert.equal(canManageTask(task, members[0]), true);
  assert.equal(canManageTask(task, members[1]), false);
  assert.equal(
    canManageTask({ ...task, created_by: "member" }, members[1]),
    true,
  );
  assert.equal(canManageTask(task, { ...members[0], active: false }), false);
});
function client(
  identity: string | null = "operator",
  result: unknown = { ...task, version: 4 },
  error: unknown = null,
) {
  const calls: { name: string; payload: unknown }[] = [];
  return {
    calls,
    api: {
      auth: {
        getUser: async () => ({
          data: { user: identity ? { id: identity } : null },
          error: null,
        }),
      },
      rpc: async (name: string, payload: unknown) => {
        calls.push({ name, payload });
        return { data: result, error };
      },
    } as unknown as SupabaseClient,
  };
}
test("no task RPC is sent for expired or changed identity", async () => {
  for (const identity of [null, "other"]) {
    const fake = client(identity);
    await assert.rejects(
      taskMutation(fake.api, "operator", {
        type: "update",
        task,
        changes: { status: "done" },
      }),
      /sessão mudou/,
    );
    assert.equal(fake.calls.length, 0);
  }
});
test("task updates carry the observed version and reject conflicts and false success", async () => {
  const fake = client();
  await taskMutation(fake.api, "operator", {
    type: "update",
    task,
    changes: { status: "done" },
  });
  assert.deepEqual(fake.calls, [
    {
      name: "console_update_task",
      payload: {
        p_id: "task-id",
        p_version: 3,
        p_changes: { status: "done" },
        p_actor: "operator",
      },
    },
  ]);
  for (const error of [{ code: "40001" }, { code: "42501" }]) {
    const rejected = client("operator", null, error);
    await assert.rejects(
      taskMutation(rejected.api, "operator", {
        type: "update",
        task,
        changes: { status: "done" },
      }),
    );
  }
  for (const row of [
    null,
    { ...task, version: 3 },
    { ...task, id: "different", version: 4 },
  ]) {
    const rejected = client("operator", row);
    await assert.rejects(
      taskMutation(rejected.api, "operator", {
        type: "update",
        task,
        changes: { status: "done" },
      }),
      /servidor não confirmou/,
    );
  }
});
test("create uses stable request ID and comments require server confirmation for the same task", async () => {
  const create = client("operator", { ...task, id: "request-id", version: 1 });
  const input = {
    id: "request-id",
    title: "Tarefa",
    description: "",
    priority: "normal" as const,
    assignee_id: "member",
    due_date: null,
  };
  await taskMutation(create.api, "operator", { type: "create", input });
  assert.deepEqual(create.calls, [
    {
      name: "console_create_task",
      payload: { p_task: input, p_actor: "operator" },
    },
  ]);
  const comment = client("operator", { id: "event", task_id: task.id });
  await taskMutation(comment.api, "operator", {
    type: "comment",
    taskId: task.id,
    body: "  Contexto  ",
  });
  assert.deepEqual(comment.calls, [
    {
      name: "console_add_task_comment",
      payload: { p_id: task.id, p_body: "Contexto", p_actor: "operator" },
    },
  ]);
  const wrong = client("operator", { id: "event", task_id: "different" });
  await assert.rejects(
    taskMutation(wrong.api, "operator", {
      type: "comment",
      taskId: task.id,
      body: "Contexto",
    }),
  );
});
