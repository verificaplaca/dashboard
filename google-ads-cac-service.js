/**
 * GoogleAdsOptimizationService
 * ─────────────────────────────────────────────────────────────────────────
 * Service de análise de CAC do Google Ads (search terms + keywords).
 * Módulo JS puro, sem dependência de framework de backend — consumido
 * diretamente pela página google-ads-cac.html via <script>.
 *
 * V1: apenas classifica e recomenda. NÃO escreve nada no Google Ads.
 *
 * Responsabilidades:
 *  - normalizar métricas (cost em micros → R$, cpa, conversion rate)
 *  - aplicar lista de termos bloqueados
 *  - verificar se search term já existe como keyword ativa
 *  - aplicar regras de recomendação (search terms e keywords)
 *  - montar payload final para cards e tabelas
 *  - registrar revisão manual local (accepted/ignored/reviewed)
 * ─────────────────────────────────────────────────────────────────────────
 */

// ─── METAS (ajustáveis conforme estratégia de CAC) ─────────────────────────
const GADS_CAC_TARGETS = {
  PROMOTE_HIGH_PRIORITY: { minPurchase: 10, minCost: 50, maxCpa: 8 },
  PROMOTE:               { minPurchase: 5,  minCost: 20, maxCpa: 9 },
  NEGATIVE_REVIEW:       { minCost: 20 },
  WATCH_SEARCH_TERM:     { minPurchase: 2, maxPurchase: 4, maxCpa: 9, minCost: 10 },
  PAUSE:                 { minPurchase: 5, minCost: 30, minCpaExclusive: 10 },
  PAUSE_NO_PURCHASE:     { minCost: 30 },
  KEEP_OR_SCALE:         { minPurchase: 5, maxCpa: 9, minCost: 20 },
  WATCH_KEYWORD:         { minCost: 20, minPurchase: 1, maxPurchase: 4, minCpaExclusive: 10 },
};

const GoogleAdsOptimizationService = (() => {

  // ── normalizeMetrics ───────────────────────────────────────────────────
  // Converte cost de micros para R$ (se vier em micros) e garante números.
  function normalizeMetrics(row) {
    const clicks      = Number(row.clicks ?? 0);
    const impressions = Number(row.impressions ?? 0);
    const purchases   = Number(row.purchases ?? 0);

    // Heurística: se vier `cost_micros`, converte. Se vier `cost` já em R$, usa direto.
    let cost;
    if (row.cost_micros !== undefined && row.cost_micros !== null) {
      cost = Number(row.cost_micros) / 1e6;
    } else {
      cost = Number(row.cost ?? 0);
    }

    return {
      ...row,
      clicks,
      impressions,
      purchases,
      cost: round2(cost),
      cpa_purchase: calculateCpa(cost, purchases),
      conversion_rate: calculateConversionRate(purchases, clicks),
    };
  }

  function round2(n) {
    return Math.round((Number(n) || 0) * 100) / 100;
  }

  // ── calculateCpa ───────────────────────────────────────────────────────
  function calculateCpa(cost, purchases) {
    if (!purchases || purchases <= 0) return null; // exibir como "—"
    return round2(cost / purchases);
  }

  // ── calculateConversionRate ───────────────────────────────────────────
  function calculateConversionRate(purchases, clicks) {
    if (!clicks || clicks <= 0) return null;
    return round2((purchases / clicks) * 100);
  }

  // ── containsBlockedTerm ────────────────────────────────────────────────
  // blockedTerms: array de { term, match_type, active }
  // match_type: 'exact' (texto inteiro igual) | 'contains' (substring) | 'word' (palavra isolada, com boundary)
  function containsBlockedTerm(text, blockedTerms) {
    if (!text) return null;
    const lower = String(text).toLowerCase();
    for (const bt of blockedTerms) {
      if (!bt.active) continue;
      const term = String(bt.term).toLowerCase();
      if (bt.match_type === 'exact') {
        if (lower === term) return bt;
      } else if (bt.match_type === 'word') {
        const escaped = term.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        const re = new RegExp(`(^|[^a-z0-9À-ÿ])${escaped}([^a-z0-9À-ÿ]|$)`, 'i');
        if (re.test(lower)) return bt;
      } else {
        if (lower.includes(term)) return bt;
      }
    }
    return null;
  }

  // ── makeEntityKey ───────────────────────────────────────────────────────
  // Chave composta de revisão: entity_type::entity_key::campaign_id::ad_group_id
  // Evita colisão entre o mesmo termo/keyword em campanhas/ad groups diferentes.
  function makeEntityKey(entityType, entityKey, campaignId, adGroupId) {
    return [entityType, entityKey, campaignId ?? '', adGroupId ?? ''].join('::');
  }

  // ── classifySearchTerm ─────────────────────────────────────────────────
  // row já deve estar normalizado (normalizeMetrics). blockedTerms: lista ativa.
  function classifySearchTerm(row, blockedTerms) {
    const blocked = containsBlockedTerm(row.search_term, blockedTerms);
    const { purchases, cost, cpa_purchase } = row;
    const isExistingKeyword = !!row.is_existing_keyword;

    if (blocked) {
      return {
        recommendation: 'NEGATIVE',
        reason: 'Termo contém expressão bloqueada pela estratégia.',
        matched_blocked_term: blocked.term,
      };
    }

    const t = GADS_CAC_TARGETS;

    if (
      purchases >= t.PROMOTE_HIGH_PRIORITY.minPurchase &&
      cost >= t.PROMOTE_HIGH_PRIORITY.minCost &&
      cpa_purchase !== null && cpa_purchase <= t.PROMOTE_HIGH_PRIORITY.maxCpa &&
      !isExistingKeyword
    ) {
      return {
        recommendation: 'PROMOTE_HIGH_PRIORITY',
        reason: 'Termo com alto volume de compras e CAC muito eficiente.',
      };
    }

    if (
      purchases >= t.PROMOTE.minPurchase &&
      cost >= t.PROMOTE.minCost &&
      cpa_purchase !== null && cpa_purchase <= t.PROMOTE.maxCpa &&
      !isExistingKeyword
    ) {
      return {
        recommendation: 'PROMOTE',
        reason: 'Termo com CAC abaixo da meta, volume mínimo de compras e custo suficiente.',
      };
    }

    if (purchases === 0 && cost >= t.NEGATIVE_REVIEW.minCost) {
      return {
        recommendation: 'NEGATIVE_REVIEW',
        reason: 'Termo consumiu verba sem gerar compra.',
      };
    }

    if (
      purchases >= t.WATCH_SEARCH_TERM.minPurchase &&
      purchases <= t.WATCH_SEARCH_TERM.maxPurchase &&
      cpa_purchase !== null && cpa_purchase <= t.WATCH_SEARCH_TERM.maxCpa &&
      cost >= t.WATCH_SEARCH_TERM.minCost
    ) {
      return {
        recommendation: 'WATCH',
        reason: 'Termo promissor, mas ainda com amostra pequena.',
      };
    }

    return {
      recommendation: 'KEEP',
      reason: 'Sem recomendação crítica no período analisado.',
    };
  }

  // ── classifyKeyword ─────────────────────────────────────────────────────
  function classifyKeyword(row) {
    const { purchases, cost, cpa_purchase } = row;
    const t = GADS_CAC_TARGETS;

    if (
      purchases >= t.PAUSE.minPurchase &&
      cost >= t.PAUSE.minCost &&
      cpa_purchase !== null && cpa_purchase > t.PAUSE.minCpaExclusive
    ) {
      return {
        recommendation: 'PAUSE',
        reason: 'Keyword com volume mínimo e CAC acima da meta.',
      };
    }

    if (purchases === 0 && cost >= t.PAUSE_NO_PURCHASE.minCost) {
      return {
        recommendation: 'PAUSE_NO_PURCHASE',
        reason: 'Keyword consumiu verba sem gerar compra.',
      };
    }

    if (
      purchases >= t.KEEP_OR_SCALE.minPurchase &&
      cpa_purchase !== null && cpa_purchase <= t.KEEP_OR_SCALE.maxCpa &&
      cost >= t.KEEP_OR_SCALE.minCost
    ) {
      return {
        recommendation: 'KEEP_OR_SCALE',
        reason: 'Keyword com CAC saudável.',
      };
    }

    if (
      cost >= t.WATCH_KEYWORD.minCost &&
      purchases >= t.WATCH_KEYWORD.minPurchase && purchases <= t.WATCH_KEYWORD.maxPurchase &&
      cpa_purchase !== null && cpa_purchase > t.WATCH_KEYWORD.minCpaExclusive
    ) {
      return {
        recommendation: 'WATCH',
        reason: 'CPA acima da meta, mas ainda com amostra moderada.',
      };
    }

    return {
      recommendation: 'KEEP',
      reason: 'Sem recomendação crítica no período analisado.',
    };
  }

  // ── buildSummary ───────────────────────────────────────────────────────
  // Monta os números dos cards do topo. Fonte financeira única = keywords
  // (search terms e keywords são visões diferentes do mesmo tráfego do Google Ads;
  // somar as duas duplicaria custo e purchases). Search terms alimentam apenas
  // contagens de recomendação (promover/negativar/observar); keywords alimentam
  // contagens de pausar/manter-escalar/observar e os números financeiros.
  function buildSummary(searchTerms, keywords) {
    const totalCost = round2(keywords.reduce((s, r) => s + r.cost, 0));
    const totalPurchases = keywords.reduce((s, r) => s + r.purchases, 0);

    const avgCac = totalPurchases > 0 ? round2(totalCost / totalPurchases) : null;

    const countBy = (list, rec) => list.filter(r => r.recommendation === rec).length;

    return {
      investimento: totalCost,
      purchases: totalPurchases,
      cac_medio: avgCac,
      termos_para_promover:
        countBy(searchTerms, 'PROMOTE_HIGH_PRIORITY') + countBy(searchTerms, 'PROMOTE'),
      keywords_para_pausar:
        countBy(keywords, 'PAUSE') + countBy(keywords, 'PAUSE_NO_PURCHASE'),
      termos_para_negativar:
        countBy(searchTerms, 'NEGATIVE') + countBy(searchTerms, 'NEGATIVE_REVIEW'),
    };
  }

  // ── markAsReviewed ─────────────────────────────────────────────────────
  // Monta o payload de revisão local (persistido via POST /api/google-ads-cac/review).
  function markAsReviewed({ entityType, entityKey, campaignId, campaignName, adGroupId, adGroupName, recommendation, actionTaken, notes }) {
    if (!['search_term', 'keyword'].includes(entityType)) {
      throw new Error('entityType inválido: deve ser "search_term" ou "keyword".');
    }
    if (!['accepted', 'ignored', 'reviewed'].includes(actionTaken)) {
      throw new Error('actionTaken inválido: deve ser "accepted", "ignored" ou "reviewed".');
    }
    return {
      entity_type: entityType,
      // campaign_id/ad_group_id são NOT NULL DEFAULT '' no banco (unique index não
      // bloqueia duplicatas com NULL) — normaliza ausência para string vazia.
      entity_key: entityKey,
      campaign_id: campaignId || '',
      campaign_name: campaignName ?? null,
      ad_group_id: adGroupId || '',
      ad_group_name: adGroupName ?? null,
      recommendation,
      action_taken: actionTaken,
      notes: notes ?? null,
    };
  }

  // ── reviewsToMap ───────────────────────────────────────────────────────
  // Agrega reviews (já filtradas por entity_type) em Map por chave composta.
  // Assume status atual = uma linha por chave (upsert no banco); se vier
  // histórico com duplicatas, usa a de updated_at mais recente.
  function reviewsToMap(reviews) {
    const map = new Map();
    for (const r of reviews) {
      const key = makeEntityKey(r.entity_type, r.entity_key, r.campaign_id, r.ad_group_id);
      const existing = map.get(key);
      if (!existing || new Date(r.updated_at) > new Date(existing.updated_at)) {
        map.set(key, r);
      }
    }
    return map;
  }

  // ── aggregateByEntity ──────────────────────────────────────────────────
  // rawRows do Supabase vêm uma linha por dia (date + entidade + campanha +
  // ad group). Para classificar e exibir corretamente um período (ex: 30d),
  // soma clicks/impressions/cost_micros/purchases de todas as linhas da
  // mesma entidade+campanha+ad group, gerando 1 linha por entidade no período.
  // entityField: 'keyword' ou 'search_term'.
  function aggregateByEntity(rawRows, entityField) {
    const byKey = new Map();
    for (const raw of rawRows) {
      const entityValue = raw[entityField];
      const key = [entityValue, raw.campaign_id ?? '', raw.ad_group_id ?? ''].join('::');
      const existing = byKey.get(key);
      if (!existing) {
        byKey.set(key, {
          ...raw,
          clicks: Number(raw.clicks ?? 0),
          impressions: Number(raw.impressions ?? 0),
          cost_micros: Number(raw.cost_micros ?? 0),
          purchases: Number(raw.purchases ?? 0),
          is_existing_keyword: !!raw.is_existing_keyword,
        });
      } else {
        existing.clicks += Number(raw.clicks ?? 0);
        existing.impressions += Number(raw.impressions ?? 0);
        existing.cost_micros += Number(raw.cost_micros ?? 0);
        existing.purchases += Number(raw.purchases ?? 0);
        existing.is_existing_keyword = existing.is_existing_keyword || !!raw.is_existing_keyword;
        // Mantém date/status_google_ads/match_type da linha mais recente (último dia visto).
        if (raw.date > existing.date) {
          existing.date = raw.date;
          existing.status_google_ads = raw.status_google_ads;
          existing.match_type = raw.match_type;
        }
      }
    }
    return Array.from(byKey.values());
  }

  // ── Pipeline completo ─────────────────────────────────────────────────
  // Recebe dados brutos + blockedTerms + reviews já salvas, devolve tudo
  // já normalizado, classificado e com status de revisão anexado.
  // Agrega por entidade+campanha+ad group antes de classificar, pois rawRows
  // vem 1 linha por dia do Supabase (ver aggregateByEntity).
  function analyzeSearchTerms(rawRows, blockedTerms, reviewsByKey) {
    const aggregated = aggregateByEntity(rawRows, 'search_term');
    return aggregated.map(raw => {
      const norm = normalizeMetrics(raw);
      const { recommendation, reason, matched_blocked_term } = classifySearchTerm(norm, blockedTerms);
      const key = makeEntityKey('search_term', norm.search_term, norm.campaign_id, norm.ad_group_id);
      const review = reviewsByKey?.get(key) ?? null;
      return {
        ...norm,
        recommendation,
        reason,
        matched_blocked_term: matched_blocked_term ?? null,
        review_status: review?.action_taken ?? null,
        review_key: key,
      };
    });
  }

  // Agrega por entidade+campanha+ad group antes de classificar, pois rawRows
  // vem 1 linha por dia do Supabase (ver aggregateByEntity).
  function analyzeKeywords(rawRows, reviewsByKey) {
    const aggregated = aggregateByEntity(rawRows, 'keyword');
    return aggregated.map(raw => {
      const norm = normalizeMetrics(raw);
      const { recommendation, reason } = classifyKeyword(norm);
      const key = makeEntityKey('keyword', norm.keyword, norm.campaign_id, norm.ad_group_id);
      const review = reviewsByKey?.get(key) ?? null;
      return {
        ...norm,
        recommendation,
        reason,
        review_status: review?.action_taken ?? null,
        review_key: key,
      };
    });
  }

  return {
    normalizeMetrics,
    calculateCpa,
    calculateConversionRate,
    containsBlockedTerm,
    makeEntityKey,
    classifySearchTerm,
    classifyKeyword,
    buildSummary,
    markAsReviewed,
    reviewsToMap,
    aggregateByEntity,
    analyzeSearchTerms,
    analyzeKeywords,
  };
})();

// Exporta para uso em <script> simples (browser) e, se aplicável, CommonJS (testes).
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { GoogleAdsOptimizationService, GADS_CAC_TARGETS };
}
