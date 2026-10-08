import { useEffect, useState } from "react";
import { ArrowUpRight, CheckCircle2, Clock3 } from "lucide-react";
import { Card } from "./components";
import { useAuth } from "./Auth";
import { getAuthClient } from "./api";
import { go } from "./format";
import {
  emptyTaskSummary,
  isTaskOverdue,
  taskDate,
  type LiveTask,
  type TaskSummary,
} from "./domain/tasks";
import "./tasks.css";
export default function TaskHomeSummary() {
  const { user, openLogin } = useAuth();
  const [summary, setSummary] = useState<TaskSummary>(emptyTaskSummary),
    [tasks, setTasks] = useState<LiveTask[]>([]),
    [loaded, setLoaded] = useState(false),
    [error, setError] = useState("");
  useEffect(() => {
    let active = true,
      running = false;
    setLoaded(false);
    setTasks([]);
    setError("");
    if (!user) return;
    const load = async () => {
      if (running) return;
      running = true;
      try {
        const client = await getAuthClient();
        const [totals, list] = await Promise.all([
          client.rpc("console_task_summary"),
          client
            .from("console_tasks")
            .select("*")
            .eq("assignee_id", user.id)
            .is("archived_at", null)
            .neq("status", "done")
            .order("due_date", { ascending: true, nullsFirst: false })
            .order("created_at")
            .limit(3),
        ]);
        if (!active) return;
        if (totals.error || list.error)
          throw new Error("Não foi possível atualizar o resumo de tarefas.");
        setSummary(totals.data as TaskSummary);
        setTasks(list.data as LiveTask[]);
        setLoaded(true);
        setError("");
      } catch (reason) {
        if (active)
          setError(
            reason instanceof Error
              ? reason.message
              : "Falha ao consultar tarefas.",
          );
      } finally {
        running = false;
      }
    };
    void load();
    const timer = setInterval(() => {
      if (document.visibilityState === "visible") void load();
    }, 30_000);
    return () => {
      active = false;
      clearInterval(timer);
    };
  }, [user?.id]);
  return (
    <Card
      title="Sua execução, em foco"
      subtitle="Prioridades e prazos da operação."
      action={
        <button className="text-button" onClick={() => go("/tarefas")}>
          Abrir tarefas
          <ArrowUpRight size={15} />
        </button>
      }
    >
      {!user ? (
        <div className="task-home-login">
          <p>
            Entre para acompanhar suas pendências e os próximos passos da
            equipe.
          </p>
          <button className="button secondary" onClick={openLogin}>
            Entrar
          </button>
        </div>
      ) : error ? (
        <p role="alert" className="info-note">
          {error}
        </p>
      ) : !loaded ? (
        <p role="status" className="muted">
          Carregando suas prioridades…
        </p>
      ) : (
        <>
          <div className="task-home-stats">
            <div>
              <CheckCircle2 size={18} />
              <span>
                <strong>{summary.mine}</strong>Minhas pendências
              </span>
            </div>
            <div>
              <Clock3 size={18} />
              <span>
                <strong>{summary.overdue}</strong>Atrasadas da operação visível
              </span>
            </div>
            <div>
              <Clock3 size={18} />
              <span>
                <strong>{summary.today}</strong>Para hoje
              </span>
            </div>
          </div>
          {tasks.length ? (
            <ul className="task-home-list">
              {tasks.map((task) => (
                <li key={task.id}>
                  <button onClick={() => go("/tarefas")}>
                    <span>{task.title}</span>
                    <small className={isTaskOverdue(task) ? "warning" : ""}>
                      {taskDate(task.due_date)}
                      {isTaskOverdue(task) ? " · Atrasada" : ""}
                    </small>
                    <ArrowUpRight size={15} />
                  </button>
                </li>
              ))}
            </ul>
          ) : (
            <p className="muted">
              Você não tem tarefas pendentes atribuídas. Consulte o quadro para
              organizar o próximo passo.
            </p>
          )}
        </>
      )}
    </Card>
  );
}
