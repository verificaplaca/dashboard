import { useEffect, useRef, useState } from "react";
import {
  Archive,
  ArrowUpRight,
  CalendarDays,
  CheckCircle2,
  Clock3,
  LayoutGrid,
  List,
  LogIn,
  MessageSquare,
  Plus,
  RefreshCw,
  RotateCcw,
  Send,
} from "lucide-react";
import {
  Avatar,
  Badge,
  Card,
  Dialog,
  Drawer,
  Empty,
  PageHeading,
  Stat,
  Tabs,
  Toolbar,
} from "./components";
import { useAuth } from "./Auth";
import { getAuthClient } from "./api";
import { useUI } from "./ui";
import {
  canManageTask,
  emptyTaskSummary,
  isTaskOverdue,
  taskDate,
  taskInput,
  taskPriorities,
  taskStatuses,
  taskToday,
  type LiveTask,
  type TaskChanges,
  type TaskEvent,
  type TaskInput,
  type TaskMember,
  type TaskStatus,
  type TaskSummary,
} from "./domain/tasks";
import { taskMutation } from "./domain/task-api";
import "./tasks.css";

const PAGE = 50;
const time = (value: string) =>
  new Intl.DateTimeFormat("pt-BR", {
    timeZone: "America/Sao_Paulo",
    day: "2-digit",
    month: "short",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date(value));
type Filters = {
  scope: string;
  search: string;
  status: string;
  priority: string;
  assignee: string;
};
const initialFilters: Filters = {
  scope: "all",
  search: "",
  status: "",
  priority: "",
  assignee: "",
};

export default function LiveTasks() {
  const { user, openLogin } = useAuth();
  if (!user)
    return (
      <>
        <PageHeading
          eyebrow="EXECUÇÃO DA OPERAÇÃO"
          title="Tarefas"
          description="Organize prioridades, acompanhe prazos e registre o andamento da equipe."
        />
        <Card>
          <Empty
            title="Entre para acessar suas tarefas"
            description="Tarefas, comentários e responsáveis são privados. Use sua conta do Verifica Placa."
          >
            <button className="button primary" onClick={openLogin}>
              <LogIn size={16} />
              Entrar
            </button>
          </Empty>
        </Card>
      </>
    );
  return <TaskWorkspace key={user.id} userId={user.id} />;
}

function TaskWorkspace({ userId }: { userId: string }) {
  const { notify } = useUI();
  const [members, setMembers] = useState<TaskMember[]>([]),
    [membershipLoaded, setMembershipLoaded] = useState(false);
  const [rows, setRows] = useState<LiveTask[]>([]),
    [total, setTotal] = useState(0),
    [summary, setSummary] = useState<TaskSummary>(emptyTaskSummary);
  const [filters, setFilters] = useState<Filters>(initialFilters),
    [query, setQuery] = useState(""),
    [view, setView] = useState("board"),
    [limit, setLimit] = useState(PAGE);
  const [loading, setLoading] = useState(true),
    [error, setError] = useState(""),
    [busy, setBusy] = useState(false),
    [updated, setUpdated] = useState("");
  const [editor, setEditor] = useState<{ id: string; task?: LiveTask } | null>(
      null,
    ),
    [detail, setDetail] = useState<LiveTask | null>(null),
    [archive, setArchive] = useState<LiveTask | null>(null);
  const generation = useRef(0),
    running = useRef(false),
    active = useRef(true);
  const detailRef = useRef<LiveTask | null>(null);
  detailRef.current = detail;
  const me = members.find(
    (member) => member.user_id === userId && member.active,
  );
  useEffect(() => {
    active.current = true;
    return () => {
      active.current = false;
      generation.current++;
    };
  }, []);
  useEffect(() => {
    const timer = setTimeout(
      () => setFilters((previous) => ({ ...previous, search: query.trim() })),
      350,
    );
    return () => clearTimeout(timer);
  }, [query]);
  useEffect(() => {
    setLimit(PAGE);
  }, [filters]);
  const load = async (quiet = false) => {
    const request = ++generation.current;
    if (!quiet) setLoading(true);
    try {
      const client = await getAuthClient();
      const directory = await client
        .from("console_task_members")
        .select("user_id,display_name,role,active")
        .order("display_name");
      if (directory.error)
        throw new Error(
          "Não foi possível carregar as permissões. Tente atualizar.",
        );
      const people = directory.data as TaskMember[];
      if (!active.current || request !== generation.current) return;
      setMembers(people);
      setMembershipLoaded(true);
      if (
        !people.some((person) => person.user_id === userId && person.active)
      ) {
        setRows([]);
        setSummary(emptyTaskSummary);
        setDetail(null);
        setEditor(null);
        return;
      }
      let selection = client
        .from("console_tasks")
        .select("*", { count: "exact" });
      selection =
        filters.scope === "archived"
          ? selection.not("archived_at", "is", null)
          : selection.is("archived_at", null);
      if (filters.scope === "mine")
        selection = selection.eq("assignee_id", userId).neq("status", "done");
      if (filters.scope === "overdue")
        selection = selection.lt("due_date", taskToday()).neq("status", "done");
      if (filters.scope === "today")
        selection = selection.eq("due_date", taskToday()).neq("status", "done");
      if (filters.scope === "open") selection = selection.neq("status", "done");
      if (filters.status) selection = selection.eq("status", filters.status);
      if (filters.priority)
        selection = selection.eq("priority", filters.priority);
      if (filters.assignee)
        selection =
          filters.assignee === "unassigned"
            ? selection.is("assignee_id", null)
            : selection.eq("assignee_id", filters.assignee);
      if (filters.search)
        selection = selection.ilike(
          "title",
          `%${filters.search.replace(/[\\%_]/g, "\\$&")}%`,
        );
      const detailId = detailRef.current?.id;
      const [tasks, totals, selected] = await Promise.all([
        selection
          .order("created_at", { ascending: false })
          .order("id")
          .range(0, limit - 1),
        client.rpc("console_task_summary"),
        detailId
          ? client
              .from("console_tasks")
              .select("*")
              .eq("id", detailId)
              .maybeSingle()
          : Promise.resolve({ data: null, error: null }),
      ]);
      if (tasks.error || totals.error || selected.error)
        throw new Error(
          "Não foi possível atualizar as tarefas. Os dados exibidos podem estar desatualizados.",
        );
      if (!active.current || request !== generation.current) return;
      const loaded = tasks.data as LiveTask[];
      setRows(loaded);
      setTotal(tasks.count || 0);
      setSummary(totals.data as TaskSummary);
      setUpdated(time(new Date().toISOString()));
      setError("");
      if (detailId)
        setDetail((previous) =>
          previous?.id === detailId
            ? (selected.data as LiveTask | null)
            : previous,
        );
    } catch (reason) {
      if (active.current && request === generation.current)
        setError(
          reason instanceof Error
            ? reason.message
            : "Não foi possível carregar tarefas.",
        );
    } finally {
      if (active.current && request === generation.current) setLoading(false);
    }
  };
  useEffect(() => {
    void load();
    const refresh = () => {
      if (document.visibilityState === "visible" && !running.current)
        void load(true);
    };
    const timer = setInterval(refresh, 30_000);
    window.addEventListener("focus", refresh);
    return () => {
      generation.current++;
      clearInterval(timer);
      window.removeEventListener("focus", refresh);
    };
  }, [filters, limit]);
  const setFilter = (changes: Partial<Filters>) =>
    setFilters((previous) => ({ ...previous, ...changes }));
  const mutate = async (
    action: Parameters<typeof taskMutation>[2],
    message: string,
  ) => {
    if (running.current) return false;
    running.current = true;
    setBusy(true);
    setError("");
    try {
      const row = await taskMutation(await getAuthClient(), userId, action);
      if (!active.current) return true;
      if (action.type !== "comment") {
        const task = row as LiveTask;
        setDetail((previous) => (previous?.id === task.id ? task : previous));
      }
      notify(message);
      void load(true);
      return true;
    } catch (reason) {
      if (active.current) {
        const message =
          reason instanceof Error ? reason.message : "Não foi possível salvar.";
        setError(message);
        throw new Error(message);
      }
      return false;
    } finally {
      running.current = false;
      if (active.current) setBusy(false);
    }
  };
  const move = (task: LiveTask, status: TaskStatus) => {
    if (task.status === status || task.archived_at) return;
    void mutate(
      { type: "update", task, changes: { status } },
      status === "done" ? "Tarefa concluída." : "Andamento atualizado.",
    ).catch(() => {});
  };
  const name = (id: string | null) =>
    id
      ? members.find((member) => member.user_id === id)?.display_name ||
        "Conta indisponível"
      : "Sem responsável";
  if (membershipLoaded && !me)
    return (
      <>
        <PageHeading
          title="Tarefas"
          description="Gestão da execução da equipe."
        />
        <Card>
          <Empty
            title="Acesso ainda não habilitado"
            description="Sua conta precisa ser adicionada à equipe de tarefas por um administrador. As permissões são verificadas no servidor."
          >
            <button className="button secondary" onClick={() => void load()}>
              Verificar acesso
            </button>
          </Empty>
        </Card>
      </>
    );
  const taskCard = (task: LiveTask) => (
    <article
      className={`task-card priority-${task.priority} ${task.archived_at ? "task-archived" : ""}`}
      key={task.id}
      draggable={!busy && !task.archived_at}
      onDragStart={(event) => event.dataTransfer.setData("text/plain", task.id)}
    >
      <div className="task-card-top">
        <span className={`priority-tag ${task.priority}`}>
          {taskPriorities[task.priority]} prioridade
        </span>
        <button
          className="icon-button"
          aria-label={`Concluir ${task.title}`}
          disabled={busy || task.status === "done" || !!task.archived_at}
          onClick={() => move(task, "done")}
        >
          <CheckCircle2 size={17} />
        </button>
      </div>
      <button className="task-title-button" onClick={() => setDetail(task)}>
        {task.title}
        <ArrowUpRight size={14} />
      </button>
      {task.description && <p className="task-excerpt">{task.description}</p>}
      <div className="task-card-footer">
        <span className={isTaskOverdue(task) ? "warning" : ""}>
          <Clock3 size={14} />
          {taskDate(task.due_date)}
          {isTaskOverdue(task) ? " · Atrasada" : ""}
        </span>
        <span title={name(task.assignee_id)}>
          <Avatar size="small" name={name(task.assignee_id)} />
          <span className="task-assignee">{name(task.assignee_id)}</span>
        </span>
      </div>
      <select
        aria-label={`Estado de ${task.title}`}
        value={task.status}
        disabled={busy || !!task.archived_at}
        onChange={(event) => move(task, event.target.value as TaskStatus)}
      >
        {Object.entries(taskStatuses).map(([value, label]) => (
          <option value={value} key={value}>
            {label}
          </option>
        ))}
      </select>
    </article>
  );
  return (
    <div className="live-tasks">
      <PageHeading
        eyebrow="EXECUÇÃO DA OPERAÇÃO"
        title="Tarefas"
        description="Prioridades claras, responsáveis definidos e histórico de cada próximo passo."
      >
        <button
          className="button secondary"
          disabled={loading || busy}
          onClick={() => void load()}
        >
          <RefreshCw size={16} />
          Atualizar
        </button>
        <button
          className="button primary"
          disabled={!me || busy}
          onClick={() => setEditor({ id: crypto.randomUUID() })}
        >
          <Plus size={16} />
          Nova tarefa
        </button>
      </PageHeading>
      {error && (
        <p role="alert" className="task-error">
          {error}
          <button
            className="text-button"
            disabled={busy || loading}
            onClick={() => void load()}
          >
            Atualizar dados
          </button>
        </p>
      )}
      <div className="stats-grid four">
        <Stat
          label="Em aberto"
          value={loading && !updated ? "—" : String(summary.open)}
          hint="Tarefas ativas ainda não concluídas"
          icon={<List size={17} />}
          onClick={() => setFilter({ scope: "open", status: "" })}
        />
        <Stat
          label="Minhas pendências"
          value={loading && !updated ? "—" : String(summary.mine)}
          hint="Você é o responsável pela execução"
          icon={<CheckCircle2 size={17} />}
          onClick={() => setFilter({ scope: "mine", status: "" })}
        />
        <Stat
          label="Atrasadas"
          value={loading && !updated ? "—" : String(summary.overdue)}
          hint="Prazo vencido · horário de São Paulo"
          icon={<Clock3 size={17} />}
          onClick={() => setFilter({ scope: "overdue", status: "" })}
        />
        <Stat
          label="Para hoje"
          value={loading && !updated ? "—" : String(summary.today)}
          hint="Pendências com vencimento hoje"
          icon={<CalendarDays size={17} />}
          onClick={() => setFilter({ scope: "today", status: "" })}
        />
      </div>
      <div className="task-scope-bar">
        <Tabs
          values={[
            { id: "all", label: "Todas" },
            { id: "mine", label: "Minhas" },
            { id: "overdue", label: "Atrasadas" },
            { id: "today", label: "Hoje" },
            { id: "archived", label: "Arquivadas" },
            ...(filters.scope === "open"
              ? [{ id: "open", label: "Em aberto" }]
              : []),
          ]}
          selected={filters.scope}
          onSelect={(scope) => setFilter({ scope, status: "" })}
        />
        <span className="task-sync" role="status">
          {loading
            ? "Atualizando…"
            : updated
              ? `Atualizado ${updated}`
              : "Carregando…"}
        </span>
      </div>
      <Toolbar
        value={query}
        onChange={setQuery}
        placeholder="Buscar pelo título da tarefa…"
      >
        <select
          aria-label="Estado das tarefas"
          value={filters.status}
          onChange={(event) => setFilter({ status: event.target.value })}
        >
          <option value="">Todos os estados</option>
          {Object.entries(taskStatuses).map(([value, label]) => (
            <option value={value} key={value}>
              {label}
            </option>
          ))}
        </select>
        <select
          aria-label="Prioridade das tarefas"
          value={filters.priority}
          onChange={(event) => setFilter({ priority: event.target.value })}
        >
          <option value="">Todas as prioridades</option>
          {Object.entries(taskPriorities).map(([value, label]) => (
            <option value={value} key={value}>
              {label}
            </option>
          ))}
        </select>
        <select
          aria-label="Responsável das tarefas"
          value={filters.assignee}
          onChange={(event) => setFilter({ assignee: event.target.value })}
        >
          <option value="">Todos os responsáveis</option>
          <option value="unassigned">Sem responsável</option>
          {members
            .filter((member) => member.active)
            .map((member) => (
              <option key={member.user_id} value={member.user_id}>
                {member.display_name}
              </option>
            ))}
        </select>
        <div className="segment-control">
          <button
            aria-label="Quadro de tarefas"
            aria-pressed={view === "board"}
            className={view === "board" ? "selected" : ""}
            onClick={() => setView("board")}
          >
            <LayoutGrid size={17} />
          </button>
          <button
            aria-label="Lista de tarefas"
            aria-pressed={view === "list"}
            className={view === "list" ? "selected" : ""}
            onClick={() => setView("list")}
          >
            <List size={17} />
          </button>
        </div>
      </Toolbar>
      <div className="task-result-context">
        <span>
          {me?.role === "manager"
            ? "Tarefas da operação"
            : "Tarefas criadas por você ou atribuídas a você"}{" "}
          · {total} {total === 1 ? "resultado" : "resultados"}
        </span>
        {(query ||
          filters.status ||
          filters.priority ||
          filters.assignee ||
          filters.scope !== "all") && (
          <button
            className="text-button"
            onClick={() => {
              setQuery("");
              setFilters(initialFilters);
            }}
          >
            Limpar filtros
          </button>
        )}
      </div>
      {loading && !updated ? (
        <Card>
          <p role="status">Carregando tarefas da operação…</p>
        </Card>
      ) : view === "board" ? (
        <div className="task-board live-task-board">
          {Object.entries(taskStatuses).map(([status, label]) => (
            <section
              key={status}
              className={`task-column status-${status}`}
              aria-label={label}
              onDragOver={(event) => {
                if (!busy && filters.scope !== "archived")
                  event.preventDefault();
              }}
              onDrop={(event) => {
                event.preventDefault();
                const task = rows.find(
                  (row) => row.id === event.dataTransfer.getData("text/plain"),
                );
                if (task && !busy) move(task, status as TaskStatus);
              }}
            >
              <div className="kanban-heading">
                <span>
                  <i />
                  {label}
                  <em>
                    {rows.filter((task) => task.status === status).length}
                  </em>
                </span>
              </div>
              {rows.filter((task) => task.status === status).map(taskCard)}
              {!rows.some((task) => task.status === status) && (
                <div className="kanban-empty">Nenhuma tarefa nesta etapa.</div>
              )}
            </section>
          ))}
        </div>
      ) : (
        <Card>
          <div className="table-scroll">
            <table>
              <caption className="sr-only">Tarefas da operação</caption>
              <thead>
                <tr>
                  <th>Tarefa</th>
                  <th>Responsável</th>
                  <th>Prazo</th>
                  <th>Prioridade</th>
                  <th>Estado</th>
                  <th>Detalhes</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((task) => (
                  <tr key={task.id}>
                    <td>
                      <button
                        className="text-button task-table-title"
                        onClick={() => setDetail(task)}
                      >
                        {task.title}
                      </button>
                      <small>
                        {task.archived_at
                          ? "Arquivada"
                          : task.description.slice(0, 70)}
                      </small>
                    </td>
                    <td>{name(task.assignee_id)}</td>
                    <td className={isTaskOverdue(task) ? "warning" : ""}>
                      {taskDate(task.due_date)}
                      {isTaskOverdue(task) && <small>Atrasada</small>}
                    </td>
                    <td>
                      <span className={`priority-tag ${task.priority}`}>
                        {taskPriorities[task.priority]}
                      </span>
                    </td>
                    <td>
                      <select
                        aria-label={`Estado de ${task.title}`}
                        disabled={busy || !!task.archived_at}
                        value={task.status}
                        onChange={(event) =>
                          move(task, event.target.value as TaskStatus)
                        }
                      >
                        {Object.entries(taskStatuses).map(([value, label]) => (
                          <option key={value} value={value}>
                            {label}
                          </option>
                        ))}
                      </select>
                    </td>
                    <td>
                      <button
                        className="text-button"
                        onClick={() => setDetail(task)}
                      >
                        Abrir
                        <ArrowUpRight size={14} />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}
      {!loading && !rows.length && (
        <Empty
          title={
            filters.scope === "archived"
              ? "Nenhuma tarefa arquivada"
              : total
                ? "Nenhuma tarefa carregada"
                : "Nenhuma tarefa encontrada"
          }
          description={
            query ||
            filters.scope !== "all" ||
            filters.status ||
            filters.priority ||
            filters.assignee
              ? "Ajuste os filtros para consultar outras tarefas."
              : "Crie a primeira tarefa e defina o próximo passo da operação."
          }
        >
          {me && filters.scope === "all" && !query && (
            <button
              className="button secondary"
              onClick={() => setEditor({ id: crypto.randomUUID() })}
            >
              <Plus size={16} />
              Criar primeira tarefa
            </button>
          )}
        </Empty>
      )}
      {rows.length > 0 && (
        <div className="live-pager">
          <span>
            {rows.length} de {total} tarefas · contagem das colunas corresponde
            aos registros carregados
          </span>
          {rows.length < total && (
            <button
              className="button secondary"
              disabled={loading}
              onClick={() => setLimit((previous) => previous + PAGE)}
            >
              Carregar mais
            </button>
          )}
        </div>
      )}
      <p className="task-scope-note">
        Os indicadores consideram todas as tarefas que sua conta pode acessar.
        Filtros afetam o quadro e a lista. Prazos seguem o calendário de São
        Paulo.
      </p>
      {editor && me && (
        <TaskEditor
          key={editor.id}
          task={editor.task}
          members={members}
          me={me}
          close={() => {
            if (!busy) setEditor(null);
          }}
          busy={busy}
          save={async (input, baseline) => {
            const action = baseline
              ? { type: "update" as const, task: baseline, changes: input }
              : { type: "create" as const, input: { ...input, id: editor.id } };
            if (
              await mutate(
                action,
                editor.task ? "Tarefa atualizada." : "Tarefa criada.",
              )
            )
              setEditor(null);
          }}
        />
      )}
      {detail && me && (
        <TaskDetail
          task={detail}
          me={me}
          members={members}
          busy={busy}
          close={() => setDetail(null)}
          edit={() => {
            setEditor({ id: detail.id, task: detail });
            setDetail(null);
          }}
          archive={() => setArchive(detail)}
          update={async (changes) => {
            await mutate(
              { type: "update", task: detail, changes },
              "Tarefa atualizada.",
            );
          }}
          comment={async (body) => {
            return await mutate(
              { type: "comment", taskId: detail.id, body },
              "Comentário registrado.",
            );
          }}
          refresh={async () => {
            const result = await (await getAuthClient())
              .from("console_tasks")
              .select("*")
              .eq("id", detail.id)
              .maybeSingle();
            if (result.error)
              throw new Error("Não foi possível atualizar esta tarefa.");
            if (!result.data) {
              setDetail((previous) =>
                previous?.id === detail.id ? null : previous,
              );
              notify("Esta tarefa não está mais disponível para sua conta.");
            } else {
              setDetail((previous) =>
                previous?.id === detail.id
                  ? (result.data as LiveTask)
                  : previous,
              );
              void load(true);
            }
          }}
        />
      )}
      {archive && (
        <Dialog
          title="Arquivar tarefa"
          close={() => {
            if (!busy) setArchive(null);
          }}
        >
          <div className="form-fields">
            {error && (
              <p role="alert" className="task-error">
                {error}
              </p>
            )}
            <p>
              A tarefa “{archive.title}” sairá do quadro ativo. O histórico será
              preservado e você poderá restaurá-la em Arquivadas.
            </p>
          </div>
          <div className="dialog-footer">
            <button
              className="button secondary"
              disabled={busy}
              onClick={() => setArchive(null)}
            >
              Cancelar
            </button>
            <button
              className="button primary"
              disabled={busy}
              onClick={() => {
                void mutate(
                  {
                    type: "update",
                    task: archive,
                    changes: { archived: true },
                  },
                  "Tarefa arquivada.",
                )
                  .then((success) => {
                    if (success) {
                      setArchive(null);
                      setDetail(null);
                    }
                  })
                  .catch(() => {});
              }}
            >
              <Archive size={16} />
              {busy ? "Arquivando…" : "Arquivar tarefa"}
            </button>
          </div>
        </Dialog>
      )}
    </div>
  );
}

function TaskEditor({
  task,
  members,
  me,
  busy,
  close,
  save,
}: {
  task?: LiveTask;
  members: TaskMember[];
  me: TaskMember;
  busy: boolean;
  close: () => void;
  save: (input: TaskInput, baseline?: LiveTask) => Promise<void>;
}) {
  const [error, setError] = useState("");
  const [baseline, setBaseline] = useState(task),
    [latest, setLatest] = useState<LiveTask | null>(null),
    [reviewing, setReviewing] = useState(false);
  const review = async () => {
    if (!task) return;
    setReviewing(true);
    try {
      const result = await (await getAuthClient())
        .from("console_tasks")
        .select("*")
        .eq("id", task.id)
        .maybeSingle();
      if (result.error || !result.data)
        throw new Error(
          "Esta tarefa não está disponível. Feche e atualize a lista.",
        );
      if (result.data.archived_at)
        throw new Error("Esta tarefa foi arquivada. Restaure antes de editar.");
      setLatest(result.data as LiveTask);
    } catch (reason) {
      setError(
        reason instanceof Error
          ? reason.message
          : "Não foi possível atualizar.",
      );
    } finally {
      setReviewing(false);
    }
  };
  return (
    <Dialog title={task ? "Editar tarefa" : "Nova tarefa"} close={close}>
      <form
        onSubmit={async (event) => {
          event.preventDefault();
          if (busy) return;
          setError("");
          try {
            const input = taskInput(
              Object.fromEntries(new FormData(event.currentTarget)),
              members,
              me.role === "manager",
            );
            await save(input, baseline);
          } catch (reason) {
            setError(
              reason instanceof Error
                ? reason.message
                : "Não foi possível salvar.",
            );
          }
        }}
      >
        <fieldset disabled={busy} className="task-fieldset">
          <div className="form-fields task-editor">
            <label>
              O que precisa ser feito?
              <input
                name="title"
                autoFocus
                required
                minLength={2}
                maxLength={160}
                defaultValue={task?.title}
                placeholder="Ex.: Revisar campanha de recuperação"
              />
            </label>
            <label>
              Descrição
              <textarea
                name="description"
                maxLength={6000}
                rows={4}
                defaultValue={task?.description}
                placeholder="Contexto, critérios de conclusão e próximos passos."
              />
            </label>
            <div className="task-form-grid">
              <label>
                Responsável
                <select
                  name="assignee_id"
                  defaultValue={task?.assignee_id ?? (task ? "" : me.user_id)}
                  required={me.role !== "manager"}
                >
                  {me.role === "manager" && (
                    <option value="">Sem responsável</option>
                  )}
                  {members
                    .filter(
                      (member) =>
                        member.active || member.user_id === task?.assignee_id,
                    )
                    .map((member) => (
                      <option
                        key={member.user_id}
                        value={member.user_id}
                        disabled={!member.active}
                      >
                        {member.display_name}
                        {member.user_id === me.user_id ? " (você)" : ""}
                      </option>
                    ))}
                </select>
              </label>
              <label>
                Prioridade
                <select
                  name="priority"
                  defaultValue={task?.priority || "normal"}
                >
                  {Object.entries(taskPriorities).map(([value, label]) => (
                    <option key={value} value={value}>
                      {label}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Prazo
                <input
                  name="due_date"
                  type="date"
                  min="2000-01-01"
                  max="2200-12-31"
                  defaultValue={task?.due_date || ""}
                />
                <small>Opcional · até o fim do dia em São Paulo</small>
              </label>
            </div>
            {error && (
              <div role="alert" className="task-error">
                <p>{error}</p>
                {task && (
                  <button
                    type="button"
                    className="button secondary"
                    disabled={reviewing}
                    onClick={() => void review()}
                  >
                    {reviewing ? "Consultando…" : "Revisar versão do servidor"}
                  </button>
                )}
              </div>
            )}
            {latest && (
              <section className="task-conflict-review">
                <h3>Versão atual no servidor</h3>
                <dl>
                  <dt>Título</dt>
                  <dd>{latest.title}</dd>
                  <dt>Descrição</dt>
                  <dd>{latest.description || "Sem descrição"}</dd>
                  <dt>Estado / prioridade</dt>
                  <dd>
                    {taskStatuses[latest.status]} ·{" "}
                    {taskPriorities[latest.priority]}
                  </dd>
                  <dt>Responsável / prazo</dt>
                  <dd>
                    {members.find(
                      (member) => member.user_id === latest.assignee_id,
                    )?.display_name || "Sem responsável"}{" "}
                    · {taskDate(latest.due_date)}
                  </dd>
                </dl>
                <p>
                  Seu rascunho foi mantido nos campos acima. Ao salvar, os
                  detalhes serão substituídos pelos valores do seu rascunho.
                </p>
                <button
                  type="button"
                  className="button secondary"
                  onClick={() => {
                    setBaseline(latest);
                    setLatest(null);
                    setError("");
                  }}
                >
                  Manter rascunho e usar esta versão
                </button>
              </section>
            )}
          </div>
          <div className="dialog-footer">
            <button type="button" className="button secondary" onClick={close}>
              Cancelar
            </button>
            <button className="button primary" type="submit">
              {busy ? "Salvando…" : task ? "Salvar alterações" : "Criar tarefa"}
            </button>
          </div>
        </fieldset>
      </form>
    </Dialog>
  );
}

function TaskDetail({
  task,
  me,
  members,
  busy,
  close,
  edit,
  archive,
  update,
  comment,
  refresh,
}: {
  task: LiveTask;
  me: TaskMember;
  members: TaskMember[];
  busy: boolean;
  close: () => void;
  edit: () => void;
  archive: () => void;
  update: (changes: TaskChanges) => Promise<void>;
  comment: (body: string) => Promise<boolean>;
  refresh: () => Promise<void>;
}) {
  const [events, setEvents] = useState<TaskEvent[]>([]),
    [count, setCount] = useState(0),
    [limit, setLimit] = useState(50),
    [loading, setLoading] = useState(true),
    [error, setError] = useState(""),
    [body, setBody] = useState(""),
    [revision, setRevision] = useState(0);
  useEffect(() => {
    let active = true;
    setLoading(true);
    getAuthClient()
      .then((client) =>
        client
          .from("console_task_events")
          .select("*", { count: "exact" })
          .eq("task_id", task.id)
          .order("created_at", { ascending: false })
          .order("id")
          .range(0, limit - 1),
      )
      .then((result) => {
        if (!active) return;
        if (result.error) setError("Não foi possível carregar o histórico.");
        else {
          setEvents(result.data as TaskEvent[]);
          setCount(result.count || 0);
          setError("");
        }
        setLoading(false);
      })
      .catch(() => {
        if (active) {
          setError("Não foi possível carregar o histórico.");
          setLoading(false);
        }
      });
    return () => {
      active = false;
    };
  }, [task.id, task.version, revision, limit]);
  useEffect(() => {
    const timer = setInterval(() => {
      if (document.visibilityState === "visible")
        setRevision((previous) => previous + 1);
    }, 30_000);
    return () => clearInterval(timer);
  }, []);
  const person = (id: string | null) =>
    id
      ? members.find((member) => member.user_id === id)?.display_name ||
        "Conta indisponível"
      : "Sem responsável";
  const manage = canManageTask(task, me);
  const describe = (key: string, value: unknown) =>
    key === "assignee_id"
      ? person(value as string | null)
      : key === "due_date"
        ? taskDate(value as string | null)
        : key === "status"
          ? taskStatuses[value as TaskStatus]
          : key === "priority"
            ? taskPriorities[value as keyof typeof taskPriorities]
            : key === "archived_at"
              ? value
                ? "Arquivada"
                : "Ativa"
              : String(value || "Sem conteúdo");
  const fieldNames: Record<string, string> = {
    title: "Título",
    description: "Descrição",
    priority: "Prioridade",
    assignee_id: "Responsável",
    due_date: "Prazo",
    status: "Estado",
    archived_at: "Arquivamento",
  };
  return (
    <Drawer title="Detalhes da tarefa" close={close}>
      <div className="task-detail">
        <div className="task-detail-header">
          <span className={`priority-tag ${task.priority}`}>
            {taskPriorities[task.priority]} prioridade
          </span>
          <Badge status={task.status}>{taskStatuses[task.status]}</Badge>
          {task.archived_at && <Badge>Arquivada</Badge>}
        </div>
        <h2>{task.title}</h2>
        <p className="task-description">
          {task.description || "Sem descrição adicional."}
        </p>
        <dl className="task-facts">
          <div>
            <dt>Responsável</dt>
            <dd>{person(task.assignee_id)}</dd>
          </div>
          <div>
            <dt>Prazo</dt>
            <dd className={isTaskOverdue(task) ? "warning" : ""}>
              {taskDate(task.due_date)}
              {isTaskOverdue(task) ? " · Atrasada" : ""}
            </dd>
          </div>
          <div>
            <dt>Criada por</dt>
            <dd>{person(task.created_by)}</dd>
          </div>
          <div>
            <dt>Atualizada</dt>
            <dd>{time(task.updated_at)}</dd>
          </div>
          {task.completed_at && (
            <div>
              <dt>Concluída</dt>
              <dd>{time(task.completed_at)}</dd>
            </div>
          )}
        </dl>
        <div className="task-detail-actions">
          {!task.archived_at && (
            <label>
              Estado
              <select
                aria-label="Estado da tarefa"
                disabled={busy}
                value={task.status}
                onChange={(event) => {
                  void update({
                    status: event.target.value as TaskStatus,
                  }).catch((reason) => setError(reason.message));
                }}
              >
                {Object.entries(taskStatuses).map(([value, label]) => (
                  <option key={value} value={value}>
                    {label}
                  </option>
                ))}
              </select>
            </label>
          )}
          {manage &&
            (!task.archived_at ? (
              <>
                <button
                  className="button secondary"
                  disabled={busy}
                  onClick={edit}
                >
                  Editar tarefa
                </button>
                <button
                  className="button secondary"
                  disabled={busy}
                  onClick={archive}
                >
                  <Archive size={15} />
                  Arquivar
                </button>
              </>
            ) : (
              <button
                className="button primary"
                disabled={busy}
                onClick={() => {
                  void update({ archived: false }).catch((reason) =>
                    setError(reason.message),
                  );
                }}
              >
                <RotateCcw size={16} />
                Restaurar tarefa
              </button>
            ))}
        </div>
        <div className="task-history-heading">
          <h3>
            <MessageSquare size={17} />
            Comentários e histórico
          </h3>
          <button
            className="icon-button"
            aria-label="Atualizar histórico"
            disabled={busy || loading}
            onClick={() => {
              void refresh()
                .then(() => setRevision((previous) => previous + 1))
                .catch((reason) => setError(reason.message));
            }}
          >
            <RefreshCw size={16} />
          </button>
        </div>
        {error && (
          <p role="alert" className="task-error">
            {error}
          </p>
        )}
        {!task.archived_at && (
          <form
            className="task-comment-form"
            onSubmit={async (event) => {
              event.preventDefault();
              if (busy || !body.trim()) return;
              setError("");
              try {
                if (await comment(body)) {
                  setBody("");
                  setRevision((previous) => previous + 1);
                }
              } catch (reason) {
                setError(
                  reason instanceof Error
                    ? reason.message
                    : "Não foi possível comentar.",
                );
              }
            }}
          >
            <label>
              Comentário
              <textarea
                value={body}
                onChange={(event) => setBody(event.target.value)}
                maxLength={2000}
                required
                rows={3}
                placeholder="Registre o andamento ou esclareça o próximo passo."
                disabled={busy}
              />
            </label>
            <div>
              <small>{body.length}/2.000</small>
              <button
                className="button primary"
                disabled={busy || !body.trim()}
              >
                <Send size={15} />
                {busy ? "Enviando…" : "Comentar"}
              </button>
            </div>
          </form>
        )}
        {loading && (
          <p role="status" className="muted">
            Atualizando histórico…
          </p>
        )}
        <ol className="task-history">
          {events.map((event) => (
            <li key={event.id}>
              <div>
                <Avatar name={event.actor_name} size="small" />
                <strong>{event.actor_name}</strong>
                <time dateTime={event.created_at}>
                  {time(event.created_at)}
                </time>
              </div>
              {event.kind === "comment" ? (
                <p className="task-event-body">{event.body}</p>
              ) : event.kind === "created" ? (
                <p>Criou a tarefa.</p>
              ) : Object.keys(event.changes).length ? (
                <ul>
                  {Object.entries(event.changes).map(([key, change]) => (
                    <li key={key}>
                      <strong>{fieldNames[key] || key}:</strong>{" "}
                      {describe(key, change.before)} →{" "}
                      {describe(key, change.after)}
                    </li>
                  ))}
                </ul>
              ) : (
                <p>Revisou a tarefa.</p>
              )}
            </li>
          ))}
        </ol>
        {!loading && !events.length && (
          <p className="muted">Nenhuma atividade disponível.</p>
        )}
        {events.length < count && (
          <button
            className="button secondary"
            disabled={loading}
            onClick={() => setLimit((previous) => previous + 50)}
          >
            Carregar histórico anterior
          </button>
        )}
      </div>
    </Drawer>
  );
}
