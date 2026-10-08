import type { SupabaseClient } from "@supabase/supabase-js";
import {
  taskError,
  type LiveTask,
  type TaskChanges,
  type TaskEvent,
  type TaskInput,
} from "./tasks";

export async function taskMutation(
  client: SupabaseClient,
  expectedUser: string,
  action:
    | { type: "create"; input: TaskInput & { id: string } }
    | { type: "update"; task: LiveTask; changes: TaskChanges }
    | { type: "comment"; taskId: string; body: string },
): Promise<LiveTask | TaskEvent> {
  const { data: identity, error } = await client.auth.getUser();
  if (error || !identity.user || identity.user.id !== expectedUser)
    throw new Error(
      "Sua sessão mudou. Entre novamente e atualize antes de salvar.",
    );
  const name =
    action.type === "create"
      ? "console_create_task"
      : action.type === "update"
        ? "console_update_task"
        : "console_add_task_comment";
  const payload =
    action.type === "create"
      ? { p_task: action.input }
      : action.type === "update"
        ? {
            p_id: action.task.id,
            p_version: action.task.version,
            p_changes: action.changes,
          }
        : { p_id: action.taskId, p_body: action.body.trim() };
  const result = await client.rpc(name, { ...payload, p_actor: expectedUser });
  if (result.error) throw new Error(taskError(result.error));
  const row = Array.isArray(result.data) ? result.data[0] : result.data;
  if (
    !row?.id ||
    (action.type === "update" &&
      (row.id !== action.task.id || row.version <= action.task.version)) ||
    (action.type === "create" && row.id !== action.input.id) ||
    (action.type === "comment" && row.task_id !== action.taskId)
  )
    throw new Error(
      "O servidor não confirmou a alteração. Atualize antes de tentar novamente.",
    );
  return row;
}
