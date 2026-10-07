import { CRM, Clients, Sales, Consultations } from './Commerce';
import { Marketing, Tracking, WhatsApp, Reports } from './Operations';
import { Tasks, Evolution, Support, Admin, Profile } from './Team';
import { Empty } from './components';
export default function Workspaces({ page, id, query }: { page: string; id?: string; query: URLSearchParams }) {
  switch (page) {
    case 'crm': return <CRM id={id} />;
    case 'clientes': return <Clients id={id} />;
    case 'vendas': return <Sales id={id} query={query} />;
    case 'consultas': return <Consultations id={id} query={query} />;
    case 'marketing': return <Marketing id={id} />;
    case 'tracking': return <Tracking />;
    case 'whatsapp': return <WhatsApp />;
    case 'relatorios': return <Reports />;
    case 'tarefas': return <Tasks />;
    case 'evolucao': return <Evolution key={query.get('user') || 'overview'} personId={query.get('user') || undefined} />;
    case 'suporte': return <Support id={id} />;
    case 'administracao': return <Admin />;
    case 'perfil': return <Profile />;
    default: return <Empty title="Página não encontrada" />;
  }
}
