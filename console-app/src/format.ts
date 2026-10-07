import { useEffect, useState } from 'react';
export const money = (cents: number) => (cents / 100).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
export const number = (value: number) => value.toLocaleString('pt-BR');
export const percent = (value: number) => `${(value * 100).toLocaleString('pt-BR', { maximumFractionDigits: 1 })}%`;
export const date = (value?: string | null) => value ? new Date(`${value.slice(0, 10)}T12:00:00`).toLocaleDateString('pt-BR', { day: '2-digit', month: 'short' }) : '—';
export const roleNames: Record<string, string> = { owner: 'Proprietário', admin: 'Administrador', manager: 'Gestor', sales: 'Comercial', support: 'Atendimento', marketing: 'Marketing', finance: 'Financeiro' };
export const stageNames: Record<string, string> = { new: 'Novo lead', qualified: 'Qualificado', proposal: 'Oferta enviada', payment: 'Aguardando pagamento', won: 'Ganho', lost: 'Perdido' };
export const statusNames: Record<string, string> = { paid: 'Pago', pending: 'Pendente', refunded: 'Estornado', delivered: 'Entregue', processing: 'Processando', failed: 'Falhou', cancelled: 'Cancelada', waiting: 'Aguardando', open: 'Aberta', internal: 'Ação interna', resolved: 'Resolvida', connected: 'Conectada', disconnected: 'Desconectada', todo: 'A fazer', doing: 'Em andamento', blocked: 'Bloqueada', done: 'Concluída', active: 'Ativa', paused: 'Pausada', confirmed: 'Confirmado', reverted: 'Revertido', accepted: 'Aceito', expired: 'Expirado', revoked: 'Revogado' };

export function go(path: string) { window.location.hash = path; }
export function useRoute() {
  const [path, setPath] = useState(window.location.hash.slice(1) || '/inicio');
  useEffect(() => { const onHash = () => setPath(window.location.hash.slice(1) || '/inicio'); window.addEventListener('hashchange', onHash); return () => window.removeEventListener('hashchange', onHash); }, []);
  return path;
}
