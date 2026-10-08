export const taskStatuses = {
  todo: "A fazer",
  doing: "Em andamento",
  blocked: "Bloqueadas",
  done: "Concluídas",
} as const;
export const taskPriorities = {
  high: "Alta",
  normal: "Normal",
  low: "Baixa",
} as const;
export type TaskStatus = keyof typeof taskStatuses;
export type TaskPriority = keyof typeof taskPriorities;
export type TaskMember = {
  user_id: string;
  display_name: string;
  role: "manager" | "member";
  active: boolean;
};
export type LiveTask = {
  id: string;
  title: string;
  description: string;
  status: TaskStatus;
  priority: TaskPriority;
  assignee_id: string | null;
  created_by: string;
  due_date: string | null;
  completed_at: string | null;
  archived_at: string | null;
  version: number;
  created_at: string;
  updated_at: string;
};
export type TaskEvent = {
  id: string;
  task_id: string;
  actor_id: string;
  actor_name: string;
  kind: "created" | "updated" | "comment";
  changes: Record<string, { before: unknown; after: unknown }>;
  body: string;
  task_version: number;
  created_at: string;
};
export type TaskInput = {
  id?: string;
  title: string;
  description: string;
  priority: TaskPriority;
  assignee_id: string | null;
  due_date: string | null;
};
export type TaskChanges = Partial<Omit<TaskInput, "id">> & {
  status?: TaskStatus;
  archived?: boolean;
};
export type TaskSummary = {
  open: number;
  mine: number;
  overdue: number;
  today: number;
  done: number;
};
export const emptyTaskSummary: TaskSummary = {
  open: 0,
  mine: 0,
  overdue: 0,
  today: 0,
  done: 0,
};
export function taskToday(now = new Date()) {
  return new Intl.DateTimeFormat("sv-SE", {
    timeZone: "America/Sao_Paulo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(now);
}
export function validTaskDate(value: string) {
  const date = new Date(`${value}T12:00:00Z`);
  return (
    /^\d{4}-\d{2}-\d{2}$/.test(value) &&
    value >= "2000-01-01" &&
    value <= "2200-12-31" &&
    Number.isFinite(date.getTime()) &&
    date.toISOString().slice(0, 10) === value
  );
}
export function taskInput(
  values: Record<string, unknown>,
  members: TaskMember[],
  manager: boolean,
): TaskInput {
  const title = String(values.title || "").trim(),
    description = String(values.description || "").trim();
  const priority = String(values.priority || "normal"),
    assignee_id = values.assignee_id ? String(values.assignee_id) : null,
    due_date = values.due_date ? String(values.due_date) : null;
  if (title.length < 2 || title.length > 160)
    throw new Error("Use um título de 2 a 160 caracteres.");
  if (description.length > 6000)
    throw new Error("A descrição pode ter até 6.000 caracteres.");
  if (!Object.hasOwn(taskPriorities, priority))
    throw new Error("Escolha uma prioridade válida.");
  if (due_date && !validTaskDate(due_date))
    throw new Error("Informe uma data válida entre 2000 e 2200.");
  if (
    (!manager && !assignee_id) ||
    (assignee_id &&
      !members.some(
        (member) => member.user_id === assignee_id && member.active,
      ))
  )
    throw new Error("Escolha um responsável ativo.");
  return {
    title,
    description,
    priority: priority as TaskPriority,
    assignee_id,
    due_date,
  };
}
export function canManageTask(task: LiveTask, member: TaskMember) {
  return (
    member.active &&
    (member.role === "manager" || task.created_by === member.user_id)
  );
}
export function isTaskOverdue(task: LiveTask, today = taskToday()) {
  return (
    !task.archived_at &&
    task.status !== "done" &&
    !!task.due_date &&
    task.due_date < today
  );
}
export const taskDate = (value: string | null) =>
  value
    ? new Intl.DateTimeFormat("pt-BR", {
        day: "2-digit",
        month: "short",
        year: "numeric",
        timeZone: "America/Sao_Paulo",
      }).format(new Date(`${value}T12:00:00Z`))
    : "Sem prazo";
export function taskError(error: { code?: string; message?: string }) {
  if (error.code === "40001")
    return "Esta tarefa foi alterada por outra pessoa. Atualize e revise antes de tentar novamente.";
  if (error.code === "42501" || error.code === "PGRST301")
    return "Sua conta não tem acesso a esta ação ou a sessão expirou. Entre novamente e atualize.";
  if (error.code === "23505")
    return "Esta tarefa já foi registrada. Atualize a lista antes de tentar novamente.";
  if (error.code === "22023")
    return error.message || "Revise os campos da tarefa.";
  return "Não foi possível salvar a tarefa. Confira a conexão e tente novamente.";
}
