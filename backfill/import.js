/**
 * import.js — Backfill manual de 90 dias (Keywords + Search terms) para o
 * módulo Google Ads CAC, a partir dos CSVs exportados manualmente do Report
 * Editor do Google Ads (sync-google-ads-cac via Edge Function não é usado
 * aqui por causa do WORKER_RESOURCE_LIMIT do Supabase Free no backfill).
 *
 * Lê:
 *   backfill/gads_keywords_90d.csv
 *   backfill/gads_search_terms_90d.csv
 *
 * Escreve (upsert idempotente, mesmas chaves da Edge Function):
 *   google_ads_keywords_daily      onConflict: date,keyword,campaign_id,ad_group_id
 *   google_ads_search_terms_daily  onConflict: date,search_term,campaign_id,ad_group_id
 *
 * CAVEATS (decisões confirmadas com o usuário em 2026-06-17):
 *   - "purchase" é uma custom column da conta, usada aqui como proxy de
 *     metrics.conversions filtrado por PURCHASE category na Edge Function.
 *     NÃO há garantia de equivalência exata — não verificado.
 *   - status_google_ads: Eligible→ENABLED, Paused→PAUSED, Not eligible→PAUSED.
 *   - is_existing_keyword (search terms): calculado cruzando com o CSV de
 *     Keywords (match por texto normalizado + mesma campaign_id/ad_group_id/date,
 *     status Eligible no dia). Aproximação — Edge Function usa keyword_view
 *     real da API, aqui é csv-vs-csv.
 *
 * Uso:
 *   node import.js --dry-run        // só mostra contagens/erros, não grava
 *   node import.js                  // grava de fato
 */

const fs = require('fs')
const path = require('path')
const { parse } = require('csv-parse/sync')
const { createClient } = require('@supabase/supabase-js')
require('dotenv').config({ path: path.join(__dirname, '.env') })

const DRY_RUN = process.argv.includes('--dry-run')
const BATCH_SIZE = 1000

const SUPABASE_URL = process.env.SUPABASE_URL
const SUPABASE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY
if (!SUPABASE_URL || !SUPABASE_KEY) {
  console.error('Faltam SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY no .env')
  process.exit(1)
}

const supabase = createClient(SUPABASE_URL, SUPABASE_KEY)

const KEYWORDS_CSV = path.join(__dirname, 'gads_keywords_90d.csv')
const SEARCH_TERMS_CSV = path.join(__dirname, 'gads_search_terms_90d.csv')

// ── Helpers ──────────────────────────────────────────────────────────────

function readReportCsv(filePath) {
  const raw = fs.readFileSync(filePath, 'utf8')
  // Primeiras 2 linhas do export do Report Editor são título + range de datas,
  // não fazem parte do CSV tabular. A 3ª linha é o header real.
  const lines = raw.split('\n')
  const tabular = lines.slice(2).join('\n')
  return parse(tabular, {
    columns: true,
    skip_empty_lines: true,
    trim: true,
  })
}

// "1,576" -> 1576 ; "0" -> 0
function parseIntLoose(v) {
  if (v == null) return 0
  const n = parseInt(String(v).replace(/,/g, ''), 10)
  return Number.isFinite(n) ? n : 0
}

// "201.72" (reais) -> 201720000 (micros)
function costToMicros(v) {
  const n = parseFloat(String(v).replace(/,/g, ''))
  if (!Number.isFinite(n)) return 0
  return Math.round(n * 1_000_000)
}

function parsePurchases(v) {
  const n = parseFloat(String(v).replace(/,/g, ''))
  return Number.isFinite(n) ? n : 0
}

const MATCH_TYPE_MAP = {
  'Broad match': 'broad',
  'Phrase match': 'phrase',
  'Exact match': 'exact',
}

const STATUS_MAP = {
  'Eligible': 'ENABLED',
  'Paused': 'PAUSED',
  'Not eligible': 'PAUSED', // decisão confirmada: trata como inativo
}

function normTerm(s) {
  return String(s ?? '').trim().toLowerCase()
}

async function upsertBatches(table, rows, onConflict) {
  let total = 0
  for (let i = 0; i < rows.length; i += BATCH_SIZE) {
    const batch = rows.slice(i, i + BATCH_SIZE)
    if (DRY_RUN) {
      total += batch.length
      continue
    }
    const { error } = await supabase.from(table).upsert(batch, { onConflict })
    if (error) {
      throw new Error(`${table} upsert falhou no batch ${i}-${i + batch.length}: ${error.message}`)
    }
    total += batch.length
    process.stdout.write(`\r${table}: ${total}/${rows.length} gravados`)
  }
  console.log()
  return total
}

// ── Keywords ─────────────────────────────────────────────────────────────

async function processKeywords() {
  console.log('Lendo', KEYWORDS_CSV)
  const records = readReportCsv(KEYWORDS_CSV)
  console.log(`  ${records.length} linhas`)

  const byKey = new Map() // date|campaign_id|ad_group_id|keyword -> row
  const unmappedMatchType = new Set()
  const unmappedStatus = new Set()

  for (const r of records) {
    const date = r['Day']
    const keyword = r['Search keyword']
    const campaignId = r['Campaign ID']
    const adGroupId = r['Ad group ID']
    if (!date || !keyword) continue

    const key = [date, campaignId, adGroupId, keyword].join('|')

    const rawMatchType = r['Search keyword match type']
    const rawStatus = r['Search keyword status']
    if (rawMatchType && !MATCH_TYPE_MAP[rawMatchType]) unmappedMatchType.add(rawMatchType)
    if (rawStatus && !STATUS_MAP[rawStatus]) unmappedStatus.add(rawStatus)

    const clicks = parseIntLoose(r['Clicks'])
    const impressions = parseIntLoose(r['Impr.'])
    const costMicros = costToMicros(r['Cost'])
    const purchases = parsePurchases(r['purchase'])

    const existing = byKey.get(key)
    if (existing) {
      existing.clicks += clicks
      existing.impressions += impressions
      existing.cost_micros += costMicros
      existing.purchases += purchases
    } else {
      byKey.set(key, {
        date,
        keyword,
        match_type: MATCH_TYPE_MAP[rawMatchType] ?? 'broad',
        campaign_id: campaignId || '',
        campaign_name: r['Campaign'] || '',
        ad_group_id: adGroupId || '',
        ad_group_name: r['Ad group'] || '',
        status_google_ads: STATUS_MAP[rawStatus] ?? 'ENABLED',
        clicks,
        impressions,
        cost_micros: costMicros,
        purchases,
        ingested_at: new Date().toISOString(),
      })
    }
  }

  if (unmappedMatchType.size) console.warn('AVISO match_type não mapeado:', [...unmappedMatchType])
  if (unmappedStatus.size) console.warn('AVISO status não mapeado:', [...unmappedStatus])

  const rows = Array.from(byKey.values())
  console.log(`  ${rows.length} linhas agregadas (chave date+campaign+ad_group+keyword)`)
  return rows
}

// ── Search terms ─────────────────────────────────────────────────────────

async function processSearchTerms(keywordRows) {
  console.log('Lendo', SEARCH_TERMS_CSV)
  const records = readReportCsv(SEARCH_TERMS_CSV)
  console.log(`  ${records.length} linhas`)

  // Index de keywords ENABLED por date|campaign_id|ad_group_id|texto normalizado,
  // para aproximar is_existing_keyword (ver caveat no topo do arquivo).
  const enabledKeywordKeys = new Set(
    keywordRows
      .filter(k => k.status_google_ads === 'ENABLED')
      .map(k => [k.date, k.campaign_id, k.ad_group_id, normTerm(k.keyword)].join('|'))
  )

  const byKey = new Map()

  for (const r of records) {
    const date = r['Day']
    const searchTerm = r['Search term']
    const campaignId = r['Campaign ID']
    const adGroupId = r['Ad group ID']
    if (!date || !searchTerm) continue

    const key = [date, campaignId, adGroupId, searchTerm].join('|')
    const clicks = parseIntLoose(r['Clicks'])
    const impressions = parseIntLoose(r['Impr.'])
    const costMicros = costToMicros(r['Cost'])
    const purchases = parsePurchases(r['purchase'])

    const existing = byKey.get(key)
    if (existing) {
      existing.clicks += clicks
      existing.impressions += impressions
      existing.cost_micros += costMicros
      existing.purchases += purchases
    } else {
      const isExisting = enabledKeywordKeys.has(
        [date, campaignId, adGroupId, normTerm(searchTerm)].join('|')
      )
      byKey.set(key, {
        date,
        search_term: searchTerm,
        campaign_id: campaignId || '',
        campaign_name: r['Campaign'] || '',
        ad_group_id: adGroupId || '',
        ad_group_name: r['Ad group'] || '',
        clicks,
        impressions,
        cost_micros: costMicros,
        purchases,
        is_existing_keyword: isExisting,
        ingested_at: new Date().toISOString(),
      })
    }
  }

  const rows = Array.from(byKey.values())
  console.log(`  ${rows.length} linhas agregadas (chave date+campaign+ad_group+search_term)`)
  return rows
}

// ── Main ─────────────────────────────────────────────────────────────────
//
// Dividido em subcomandos porque o ambiente de execução mata processos em
// background entre chamadas (sem persistência de PID) — então "prepare"
// roda rápido e grava JSON intermediário em disco; "upload" lê esse JSON e
// grava em Supabase em uma janela de linhas por vez (cabe no timeout de cada
// chamada). Tudo idempotente via upsert, então pode re-rodar upload com
// overlap sem duplicar dados.

const KEYWORDS_JSON = path.join(__dirname, '_keywords_rows.json')
const SEARCH_TERMS_JSON = path.join(__dirname, '_search_terms_rows.json')

async function cmdPrepare() {
  const keywordRows = await processKeywords()
  const searchTermRows = await processSearchTerms(keywordRows)
  fs.writeFileSync(KEYWORDS_JSON, JSON.stringify(keywordRows))
  fs.writeFileSync(SEARCH_TERMS_JSON, JSON.stringify(searchTermRows))
  console.log('\nSalvos:')
  console.log(' ', KEYWORDS_JSON, `(${keywordRows.length} linhas)`)
  console.log(' ', SEARCH_TERMS_JSON, `(${searchTermRows.length} linhas)`)
}

function getArgInt(flag, def) {
  const i = process.argv.indexOf(flag)
  if (i === -1) return def
  return parseInt(process.argv[i + 1], 10)
}

async function cmdUpload(table, jsonFile, onConflict) {
  if (!fs.existsSync(jsonFile)) {
    console.error(`${jsonFile} não existe — rode "node import.js prepare" primeiro.`)
    process.exit(1)
  }
  const rows = JSON.parse(fs.readFileSync(jsonFile, 'utf8'))
  const offset = getArgInt('--offset', 0)
  const limit = getArgInt('--limit', rows.length)
  const slice = rows.slice(offset, offset + limit)
  console.log(`${table}: gravando linhas ${offset}..${offset + slice.length} de ${rows.length}`)
  if (DRY_RUN) {
    console.log('  (dry-run, nada gravado)')
    return
  }
  await upsertBatches(table, slice, onConflict)
  console.log(`${table}: done (offset ${offset}, ${slice.length} linhas)`)
}

async function main() {
  const cmd = process.argv[2]
  console.log(DRY_RUN ? '=== DRY RUN ===' : '=== RUN ===')

  if (cmd === 'prepare') {
    await cmdPrepare()
  } else if (cmd === 'upload-keywords') {
    await cmdUpload('google_ads_keywords_daily', KEYWORDS_JSON, 'date,keyword,campaign_id,ad_group_id')
  } else if (cmd === 'upload-search-terms') {
    await cmdUpload('google_ads_search_terms_daily', SEARCH_TERMS_JSON, 'date,search_term,campaign_id,ad_group_id')
  } else {
    console.error('Uso: node import.js <prepare|upload-keywords|upload-search-terms> [--offset N] [--limit N] [--dry-run]')
    process.exit(1)
  }
}

main().catch(err => {
  console.error('ERRO:', err)
  process.exit(1)
})
