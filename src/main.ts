import './style.css'
import * as XLSX from 'xlsx'

type LeadStatus =
  | 'New Lead'
  | 'Contacted'
  | 'Follow Up'
  | 'Quotation Sent'
  | 'Negotiation'
  | 'Sample Stage'
  | 'Awaiting Order'
  | 'Order Received'
  | 'Lost Opportunity'
type ContactMethod = 'Email' | 'LinkedIn' | 'WhatsApp' | 'Phone'
type ResponseStatus = 'Pending' | 'Responded' | 'No Response' | 'Follow-up Required'
type ViewKey = 'dashboard' | 'leads' | 'contacts' | 'quotations' | 'discovery' | 'assistant'
type ReportWindow = 'daily' | 'weekly' | 'monthly'
type AssistantType =
  | 'Introduction Email'
  | 'Follow-up Email'
  | 'Quotation Email'
  | 'Complaint Response'
  | 'Meeting Summary'

interface ContactHistory {
  id: string
  dateContacted: string
  method: ContactMethod
  subject: string
  messageSent: string
  followUpDate: string
  responseStatus: ResponseStatus
}

interface Quotation {
  id: string
  quotationNumber: string
  dateSent: string
  productType: string
  quotationValue: number
  validityDate: string
  customerFeedback: string
}

interface Lead {
  id: string
  companyName: string
  country: string
  city: string
  website: string
  contactPerson: string
  email: string
  phoneNumber: string
  industryCategory: string
  sourceOfLead: string
  dateAdded: string
  status: LeadStatus
  customerType: string
  notes: string
  contacts: ContactHistory[]
  quotations: Quotation[]
}

interface DiscoveryLead {
  id: string
  companyName: string
  country: string
  city: string
  website: string
  contactPerson: string
  email: string
  phoneNumber: string
  industryCategory: string
  sourceOfLead: string
  customerType: string
  keywords: string[]
}

interface AssistantDraft {
  id: string
  type: AssistantType
  leadId: string
  createdAt: string
  subject: string
  content: string
}

interface Reminder {
  leadId: string
  leadName: string
  dueDate: string
  type: string
  summary: string
  ageDays: number
  priority: 'high' | 'medium' | 'low'
}

interface AppData {
  leads: Lead[]
  discovery: DiscoveryLead[]
  assistantDrafts: AssistantDraft[]
}

interface AppState {
  data: AppData
  view: ViewKey
  timeframe: ReportWindow
  leadSearch: string
  countryFilter: string
  statusFilter: string
  discoveryCountry: string
  discoveryCategory: string
  discoveryType: string
  discoverySearch: string
  activeLeadId: string | null
  assistantType: AssistantType
  assistantLeadId: string | null
  generatedSubject: string
  generatedOutput: string
  notice: { tone: 'success' | 'error' | 'info'; text: string } | null
}

const storageKey = 'crm-export-marketing-v1'
const pipelineStatuses: LeadStatus[] = [
  'New Lead',
  'Contacted',
  'Follow Up',
  'Quotation Sent',
  'Negotiation',
  'Sample Stage',
  'Awaiting Order',
  'Order Received',
  'Lost Opportunity',
]
const contactMethods: ContactMethod[] = ['Email', 'LinkedIn', 'WhatsApp', 'Phone']
const responseStatuses: ResponseStatus[] = ['Pending', 'Responded', 'No Response', 'Follow-up Required']
const assistantTypes: AssistantType[] = [
  'Introduction Email',
  'Follow-up Email',
  'Quotation Email',
  'Complaint Response',
  'Meeting Summary',
]

const appRoot = document.querySelector<HTMLDivElement>('#app')
if (!appRoot) throw new Error('App container was not found.')
const app: HTMLDivElement = appRoot

document.title = 'CRM Export Marketing'

const state: AppState = {
  data: loadData(),
  view: 'dashboard',
  timeframe: 'weekly',
  leadSearch: '',
  countryFilter: 'All',
  statusFilter: 'All',
  discoveryCountry: 'All',
  discoveryCategory: 'All',
  discoveryType: 'All',
  discoverySearch: '',
  activeLeadId: null,
  assistantType: 'Introduction Email',
  assistantLeadId: null,
  generatedSubject: '',
  generatedOutput: '',
  notice: null,
}

state.activeLeadId = state.data.leads[0]?.id ?? null
state.assistantLeadId = state.activeLeadId

renderApp()
bindEvents()

function loadData(): AppData {
  const stored = localStorage.getItem(storageKey)
  if (!stored) {
    const seeded = createSampleData()
    localStorage.setItem(storageKey, JSON.stringify(seeded))
    return seeded
  }
  try {
    const parsed = JSON.parse(stored) as Partial<AppData>
    return {
      leads: Array.isArray(parsed.leads) ? parsed.leads : [],
      discovery: Array.isArray(parsed.discovery) ? parsed.discovery : createSampleData().discovery,
      assistantDrafts: Array.isArray(parsed.assistantDrafts) ? parsed.assistantDrafts : [],
    }
  } catch {
    const fallback = createSampleData()
    localStorage.setItem(storageKey, JSON.stringify(fallback))
    return fallback
  }
}

function createSampleData(): AppData {
  const iso = (daysAgo: number) => {
    const date = new Date()
    date.setDate(date.getDate() - daysAgo)
    return date.toISOString().slice(0, 10)
  }

  const leads: Lead[] = [
    {
      id: createId(),
      companyName: 'PetroFlow Arabia',
      country: 'Saudi Arabia',
      city: 'Dammam',
      website: 'https://petroflowarabia.com',
      contactPerson: 'Ahmed Al Rashid',
      email: 'procurement@petroflowarabia.com',
      phoneNumber: '+966-13-555-1040',
      industryCategory: 'Oil & Gas Supply',
      sourceOfLead: 'ADIPEC UAE',
      dateAdded: iso(18),
      status: 'Follow Up',
      customerType: 'Distributor',
      notes: 'Interested in forged flanges untuk refinery maintenance project.',
      contacts: [
        {
          id: createId(),
          dateContacted: iso(15),
          method: 'Email',
          subject: 'Introduction for forged flanges and fittings',
          messageSent: 'Sent company profile, product catalog, dan key references.',
          followUpDate: iso(8),
          responseStatus: 'No Response',
        },
        {
          id: createId(),
          dateContacted: iso(7),
          method: 'WhatsApp',
          subject: 'First follow up after intro',
          messageSent: 'Short reminder and asked if they have active inquiry.',
          followUpDate: iso(1),
          responseStatus: 'Follow-up Required',
        },
      ],
      quotations: [],
    },
    {
      id: createId(),
      companyName: 'Andes Industrial Piping',
      country: 'Chile',
      city: 'Santiago',
      website: 'https://andespiping.cl',
      contactPerson: 'Camila Torres',
      email: 'camila.torres@andespiping.cl',
      phoneNumber: '+56-2-2840-1102',
      industryCategory: 'Industrial Piping',
      sourceOfLead: 'LinkedIn Search',
      dateAdded: iso(30),
      status: 'Quotation Sent',
      customerType: 'Contractor',
      notes: 'Buys carbon steel and stainless fittings for mining project.',
      contacts: [
        {
          id: createId(),
          dateContacted: iso(22),
          method: 'Email',
          subject: 'Requested technical requirement',
          messageSent: 'Asked for size range, schedule, and grade requirement.',
          followUpDate: iso(15),
          responseStatus: 'Responded',
        },
      ],
      quotations: [
        {
          id: createId(),
          quotationNumber: 'Q-2026-0418',
          dateSent: iso(14),
          productType: 'ASTM A234 WPB fittings',
          quotationValue: 18250,
          validityDate: iso(-16),
          customerFeedback: 'Buyer comparing with local stockist.',
        },
      ],
    },
    {
      id: createId(),
      companyName: 'NordValve Solutions',
      country: 'Germany',
      city: 'Hamburg',
      website: 'https://nordvalve.de',
      contactPerson: 'Julia Meyer',
      email: 'j.meyer@nordvalve.de',
      phoneNumber: '+49-40-1188-730',
      industryCategory: 'Valve Distribution',
      sourceOfLead: 'Google',
      dateAdded: iso(42),
      status: 'Order Received',
      customerType: 'Stockist',
      notes: 'Placed trial order for stainless steel ball valves.',
      contacts: [
        {
          id: createId(),
          dateContacted: iso(34),
          method: 'Email',
          subject: 'Valve portfolio introduction',
          messageSent: 'Sent approval references, certificates, and catalog.',
          followUpDate: iso(28),
          responseStatus: 'Responded',
        },
      ],
      quotations: [
        {
          id: createId(),
          quotationNumber: 'Q-2026-0332',
          dateSent: iso(31),
          productType: 'Stainless steel ball valves',
          quotationValue: 26400,
          validityDate: iso(-7),
          customerFeedback: 'Trial order confirmed.',
        },
      ],
    },
    {
      id: createId(),
      companyName: 'Delta Marine Procurement',
      country: 'Indonesia',
      city: 'Surabaya',
      website: 'https://deltamarine.id',
      contactPerson: 'Rizky Saputra',
      email: 'sales@deltamarine.id',
      phoneNumber: '+62-31-7710-882',
      industryCategory: 'Marine Engineering',
      sourceOfLead: 'TradeMap',
      dateAdded: iso(3),
      status: 'New Lead',
      customerType: 'Importer',
      notes: 'Potential buyer for forged fittings and seamless pipes.',
      contacts: [],
      quotations: [],
    },
  ]

  const discovery: DiscoveryLead[] = [
    {
      id: createId(),
      companyName: 'Atlas Piping Mexico',
      country: 'Mexico',
      city: 'Monterrey',
      website: 'https://atlaspiping.mx',
      contactPerson: 'Luis Navarro',
      email: 'imports@atlaspiping.mx',
      phoneNumber: '+52-81-5100-3388',
      industryCategory: 'Industrial Piping',
      sourceOfLead: 'Lead Discovery',
      customerType: 'Distributor',
      keywords: ['flanges', 'pipes', 'industrial supplier'],
    },
    {
      id: createId(),
      companyName: 'Borealis Energy Supplies',
      country: 'Norway',
      city: 'Stavanger',
      website: 'https://borealis-energy.no',
      contactPerson: 'Sofie Larsen',
      email: 'projects@borealis-energy.no',
      phoneNumber: '+47-51-809-440',
      industryCategory: 'Oil & Gas Supply',
      sourceOfLead: 'Lead Discovery',
      customerType: 'Contractor',
      keywords: ['valves', 'refinery', 'approved vendor'],
    },
    {
      id: createId(),
      companyName: 'Qutbi Trading WLL',
      country: 'Qatar',
      city: 'Doha',
      website: 'https://qutbitrading.qa',
      contactPerson: 'Hassan Qutbi',
      email: 'purchase@qutbitrading.qa',
      phoneNumber: '+974-4001-8890',
      industryCategory: 'Industrial Equipment Supply',
      sourceOfLead: 'G4WB',
      customerType: 'Distributor',
      keywords: ['duplicate example'],
    },
  ]

  return { leads, discovery, assistantDrafts: [] }
}

function createId(): string {
  return crypto.randomUUID()
}

function persist(): void {
  localStorage.setItem(storageKey, JSON.stringify(state.data))
}

function setNotice(text: string, tone: 'success' | 'error' | 'info' = 'info'): void {
  state.notice = { text, tone }
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
}

function formatDate(date: string): string {
  if (!date) return 'Belum ada'
  return new Intl.DateTimeFormat('id-ID', { day: '2-digit', month: 'short', year: 'numeric' }).format(new Date(date))
}

function formatCurrency(value: number): string {
  return new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD', maximumFractionDigits: 0 }).format(value)
}

function normalizeText(value: string): string {
  return value
    .toLowerCase()
    .replace(/https?:\/\//g, '')
    .replace(/^www\./, '')
    .replace(/[^\p{L}\p{N}@]+/gu, '')
}

function getLeadById(leadId: string | null): Lead | undefined {
  return state.data.leads.find((lead) => lead.id === leadId)
}

function flattenContacts(leads: Lead[]) {
  return leads.flatMap((lead) =>
    lead.contacts.map((contact) => ({
      ...contact,
      leadId: lead.id,
      companyName: lead.companyName,
      country: lead.country,
    })),
  )
}

function flattenQuotations(leads: Lead[]) {
  return leads.flatMap((lead) =>
    lead.quotations.map((quotation) => ({
      ...quotation,
      leadId: lead.id,
      companyName: lead.companyName,
      country: lead.country,
    })),
  )
}

function getCountryCounts(leads: Lead[]) {
  return leads.reduce<Record<string, number>>((acc, lead) => {
    acc[lead.country] = (acc[lead.country] ?? 0) + 1
    return acc
  }, {})
}

function getSourceCounts(leads: Lead[]) {
  return leads.reduce<Record<string, number>>((acc, lead) => {
    acc[lead.sourceOfLead] = (acc[lead.sourceOfLead] ?? 0) + 1
    return acc
  }, {})
}

function getPipelineCounts(leads: Lead[]) {
  return pipelineStatuses.map((status) => ({
    status,
    count: leads.filter((lead) => lead.status === status).length,
  }))
}

function getLastActivityDate(lead: Lead): string {
  const dates = [lead.dateAdded, ...lead.contacts.map((item) => item.dateContacted), ...lead.quotations.map((item) => item.dateSent)].filter(Boolean)
  return dates.sort().at(-1) ?? lead.dateAdded
}

function getLastFeedback(lead: Lead): string {
  return lead.quotations[0]?.customerFeedback || lead.contacts.find((item) => item.responseStatus === 'Responded')?.subject || 'Belum ada'
}

function findDuplicateLead(companyName: string, website: string, email: string, excludeId?: string): Lead | undefined {
  const companyKey = normalizeText(companyName)
  const siteKey = normalizeText(website)
  const emailKey = normalizeText(email)
  return state.data.leads.find((lead) => {
    if (lead.id === excludeId) return false
    const reasons = [
      companyKey && normalizeText(lead.companyName) === companyKey,
      siteKey && normalizeText(lead.website) === siteKey,
      emailKey && normalizeText(lead.email) === emailKey,
    ]
    return reasons.some(Boolean)
  })
}

function getReminders(leads: Lead[]): Reminder[] {
  const reminders: Reminder[] = []
  leads.forEach((lead) => {
    const sortedContacts = [...lead.contacts].sort((a, b) => a.dateContacted.localeCompare(b.dateContacted))
    const firstEmail = sortedContacts.find((item) => item.method === 'Email')
    const lastActivity = getLastActivityDate(lead)

    if (firstEmail) {
      const diff = daysFrom(firstEmail.dateContacted)
      if (diff >= 7 && !sortedContacts.some((item) => item.responseStatus === 'Responded')) {
        reminders.push({
          leadId: lead.id,
          leadName: lead.companyName,
          dueDate: firstEmail.followUpDate || firstEmail.dateContacted,
          type: 'Follow Up',
          summary: `${lead.companyName} belum reply sejak intro pertama.`,
          ageDays: diff,
          priority: 'high',
        })
      }
    }

    const latestFollowUp = sortedContacts.filter((item) => item.followUpDate).sort((a, b) => b.followUpDate.localeCompare(a.followUpDate)).at(0)
    if (latestFollowUp && daysFrom(latestFollowUp.followUpDate) >= 14) {
      reminders.push({
        leadId: lead.id,
        leadName: lead.companyName,
        dueDate: latestFollowUp.followUpDate,
        type: 'Overdue Follow Up',
        summary: `${lead.companyName} sudah lewat jadwal follow up.`,
        ageDays: daysFrom(latestFollowUp.followUpDate),
        priority: 'medium',
      })
    }

    if (daysFrom(lastActivity) >= 30 && lead.status !== 'Order Received' && lead.status !== 'Lost Opportunity') {
      reminders.push({
        leadId: lead.id,
        leadName: lead.companyName,
        dueDate: lastActivity,
        type: 'Inactive',
        summary: `${lead.companyName} tidak ada aktivitas 30+ hari.`,
        ageDays: daysFrom(lastActivity),
        priority: 'low',
      })
    }
  })
  return reminders.sort((a, b) => b.ageDays - a.ageDays)
}

function daysFrom(date: string): number {
  const from = new Date(date)
  const now = new Date()
  return Math.floor((now.getTime() - from.getTime()) / 86400000)
}

function getCutoff(window: ReportWindow): Date {
  const date = new Date()
  date.setDate(date.getDate() - (window === 'daily' ? 1 : window === 'weekly' ? 7 : 30))
  return date
}

function isOnOrAfter(dateValue: string, cutoff: Date): boolean {
  return new Date(dateValue) >= cutoff
}

function getWindowMetrics() {
  const cutoff = getCutoff(state.timeframe)
  const contacts = flattenContacts(state.data.leads)
  const quotations = flattenQuotations(state.data.leads)
  return {
    newLeads: state.data.leads.filter((lead) => isOnOrAfter(lead.dateAdded, cutoff)).length,
    emailsSent: contacts.filter((contact) => contact.method === 'Email' && isOnOrAfter(contact.dateContacted, cutoff)).length,
    responsesReceived: contacts.filter((contact) => contact.responseStatus === 'Responded' && isOnOrAfter(contact.dateContacted, cutoff)).length,
    quotationsSent: quotations.filter((quotation) => isOnOrAfter(quotation.dateSent, cutoff)).length,
    ordersWon: state.data.leads.filter((lead) => lead.status === 'Order Received' && isOnOrAfter(getLastActivityDate(lead), cutoff)).length,
    ordersLost: state.data.leads.filter((lead) => lead.status === 'Lost Opportunity' && isOnOrAfter(getLastActivityDate(lead), cutoff)).length,
  }
}

function getRecentActivity() {
  const contactEntries = flattenContacts(state.data.leads).map((entry) => ({
    date: entry.dateContacted,
    label: `${entry.companyName}: ${entry.method}`,
    detail: `${entry.subject} • ${entry.responseStatus}`,
  }))
  const quotationEntries = flattenQuotations(state.data.leads).map((entry) => ({
    date: entry.dateSent,
    label: `${entry.companyName}: ${entry.quotationNumber}`,
    detail: `${entry.productType} • ${formatCurrency(entry.quotationValue)}`,
  }))
  return [...contactEntries, ...quotationEntries].sort((a, b) => b.date.localeCompare(a.date)).slice(0, 8)
}

function getFilteredLeads(): Lead[] {
  const term = state.leadSearch.trim().toLowerCase()
  return state.data.leads
    .filter((lead) => {
      const matchesSearch =
        !term ||
        [lead.companyName, lead.country, lead.city, lead.website, lead.contactPerson, lead.email, lead.industryCategory, lead.sourceOfLead]
          .join(' ')
          .toLowerCase()
          .includes(term)
      const matchesCountry = state.countryFilter === 'All' || lead.country === state.countryFilter
      const matchesStatus = state.statusFilter === 'All' || lead.status === state.statusFilter
      return matchesSearch && matchesCountry && matchesStatus
    })
    .sort((a, b) => b.dateAdded.localeCompare(a.dateAdded))
}

function getFilteredDiscovery(): Array<DiscoveryLead & { duplicate?: Lead }> {
  return state.data.discovery
    .map((candidate) => ({
      ...candidate,
      duplicate: findDuplicateLead(candidate.companyName, candidate.website, candidate.email),
    }))
    .filter((candidate) => {
      const term = state.discoverySearch.trim().toLowerCase()
      const matchesSearch =
        !term ||
        [candidate.companyName, candidate.country, candidate.city, candidate.website, candidate.contactPerson, candidate.email, candidate.industryCategory, candidate.customerType, candidate.keywords.join(' ')]
          .join(' ')
          .toLowerCase()
          .includes(term)
      const matchesCountry = state.discoveryCountry === 'All' || candidate.country === state.discoveryCountry
      const matchesCategory = state.discoveryCategory === 'All' || candidate.industryCategory === state.discoveryCategory
      const matchesType = state.discoveryType === 'All' || candidate.customerType === state.discoveryType
      return matchesSearch && matchesCountry && matchesCategory && matchesType
    })
}

function renderApp(): void {
  const filteredLeads = getFilteredLeads()
  const activeLead = getLeadById(state.activeLeadId) ?? filteredLeads[0] ?? state.data.leads[0]
  const reminders = getReminders(state.data.leads)
  const metrics = getWindowMetrics()
  const countries = ['All', ...Object.keys(getCountryCounts(state.data.leads)).sort()]
  const discoveryCountries = ['All', ...new Set(state.data.discovery.map((item) => item.country)).values()]
  const discoveryCategories = ['All', ...new Set(state.data.discovery.map((item) => item.industryCategory)).values()]
  const discoveryTypes = ['All', ...new Set(state.data.discovery.map((item) => item.customerType)).values()]

  if (activeLead) {
    state.activeLeadId = activeLead.id
    if (!state.assistantLeadId) state.assistantLeadId = activeLead.id
  }

  app.innerHTML = `
    <div class="shell">
      <header class="hero-shell">
        <div>
          <p class="eyebrow">Export sales control</p>
          <h1>CRM Export Marketing</h1>
          <p class="hero-copy">
            Simpan company yang sedang kamu tawarkan, kontrol status, lihat histori kontak dan quotation, lalu pakai reminder untuk tahu siapa yang harus kamu follow up hari ini.
          </p>
        </div>
        <div class="hero-actions">
          <button class="btn btn-primary" data-action="export-excel">Export to Excel</button>
          <button class="btn" data-action="load-sample">Reload sample</button>
          <button class="btn btn-danger" data-action="reset-data">Reset data</button>
        </div>
      </header>

      <section class="highlight-grid">
        <article class="highlight-card"><span>Total company</span><strong>${state.data.leads.length}</strong><small>${Object.keys(getCountryCounts(state.data.leads)).length} country coverage</small></article>
        <article class="highlight-card"><span>Due follow up</span><strong>${reminders.length}</strong><small>${reminders.filter((item) => item.priority === 'high').length} high priority</small></article>
        <article class="highlight-card"><span>Quotation sent</span><strong>${flattenQuotations(state.data.leads).length}</strong><small>${formatCurrency(flattenQuotations(state.data.leads).reduce((sum, item) => sum + item.quotationValue, 0))}</small></article>
        <article class="highlight-card"><span>Untouched discovery</span><strong>${getFilteredDiscovery().filter((item) => !item.duplicate).length}</strong><small>potential new customer</small></article>
      </section>

      <nav class="tabs">${renderTabs()}</nav>

      ${
        state.notice
          ? `<div class="notice ${state.notice.tone}">
              <span>${escapeHtml(state.notice.text)}</span>
              <button class="icon-btn" data-action="dismiss-notice">×</button>
            </div>`
          : ''
      }

      <main class="main-content">
        ${
          state.view === 'dashboard'
            ? renderDashboard(metrics, reminders)
            : state.view === 'leads'
              ? renderLeadDatabase(filteredLeads, activeLead, countries)
              : state.view === 'contacts'
                ? renderContacts(activeLead)
                : state.view === 'quotations'
                  ? renderQuotations(activeLead)
                  : state.view === 'discovery'
                    ? renderDiscovery(discoveryCountries, discoveryCategories, discoveryTypes)
                    : renderAssistant()
        }
      </main>
    </div>
  `
}

function renderTabs(): string {
  const tabs: Array<{ key: ViewKey; label: string }> = [
    { key: 'dashboard', label: 'Dashboard' },
    { key: 'leads', label: 'Lead Database' },
    { key: 'contacts', label: 'Contact History' },
    { key: 'quotations', label: 'Quotation Management' },
    { key: 'discovery', label: 'Lead Discovery' },
    { key: 'assistant', label: 'AI Assistant' },
  ]
  return tabs.map((tab) => `<button class="tab ${state.view === tab.key ? 'active' : ''}" data-view="${tab.key}">${tab.label}</button>`).join('')
}

function renderDashboard(metrics: ReturnType<typeof getWindowMetrics>, reminders: Reminder[]): string {
  const countryCounts = Object.entries(getCountryCounts(state.data.leads)).sort((a, b) => b[1] - a[1]).slice(0, 6)
  const sourceCounts = Object.entries(getSourceCounts(state.data.leads)).sort((a, b) => b[1] - a[1]).slice(0, 6)
  const pipelineCounts = getPipelineCounts(state.data.leads)
  const recent = getRecentActivity()
  return `
    <section class="section-stack">
      <div class="section-head">
        <div>
          <h2>Overview</h2>
          <p>Lihat company mana saja yang ada di database, statusnya, dan siapa yang harus difollow up.</p>
        </div>
        <div class="segment-control">
          <button data-timeframe="daily" class="${state.timeframe === 'daily' ? 'active' : ''}">Daily</button>
          <button data-timeframe="weekly" class="${state.timeframe === 'weekly' ? 'active' : ''}">Weekly</button>
          <button data-timeframe="monthly" class="${state.timeframe === 'monthly' ? 'active' : ''}">Monthly</button>
        </div>
      </div>

      <div class="metric-grid">
        <article class="metric-card"><span>New leads</span><strong>${metrics.newLeads}</strong></article>
        <article class="metric-card"><span>Emails sent</span><strong>${metrics.emailsSent}</strong></article>
        <article class="metric-card"><span>Responses</span><strong>${metrics.responsesReceived}</strong></article>
        <article class="metric-card"><span>Quotations</span><strong>${metrics.quotationsSent}</strong></article>
        <article class="metric-card"><span>Orders won</span><strong>${metrics.ordersWon}</strong></article>
        <article class="metric-card"><span>Orders lost</span><strong>${metrics.ordersLost}</strong></article>
      </div>

      <div class="dashboard-grid">
        <article class="panel">
          <div class="panel-head"><h3>Leads by country</h3><span>Market coverage</span></div>
          <div class="stack-list">
            ${countryCounts.map(([country, count]) => `<div class="row-line"><span>${escapeHtml(country)}</span><strong>${count}</strong></div>`).join('')}
          </div>
        </article>

        <article class="panel">
          <div class="panel-head"><h3>Leads by source</h3><span>Where your companies come from</span></div>
          <div class="stack-list">
            ${sourceCounts.map(([source, count]) => `<div class="row-line"><span>${escapeHtml(source)}</span><strong>${count}</strong></div>`).join('')}
          </div>
        </article>

        <article class="panel">
          <div class="panel-head"><h3>Status pipeline</h3><span>Current company position</span></div>
          <div class="stack-list">
            ${pipelineCounts
              .filter((item) => item.count > 0)
              .map((item) => `<div class="row-line"><span>${escapeHtml(item.status)}</span><strong>${item.count}</strong></div>`)
              .join('')}
          </div>
        </article>

        <article class="panel">
          <div class="panel-head"><h3>Reminder system</h3><span>Who needs attention today</span></div>
          <div class="stack-list">
            ${
              reminders.length
                ? reminders
                    .slice(0, 6)
                    .map(
                      (reminder) => `
                        <button class="reminder-card priority-${reminder.priority}" data-action="select-lead" data-id="${reminder.leadId}">
                          <div>
                            <strong>${escapeHtml(reminder.leadName)}</strong>
                            <p>${escapeHtml(reminder.summary)}</p>
                          </div>
                          <small>${formatDate(reminder.dueDate)}</small>
                        </button>
                      `,
                    )
                    .join('')
                : '<p class="empty">No reminder at the moment.</p>'
            }
          </div>
        </article>

        <article class="panel full-span">
          <div class="panel-head"><h3>Recent movements</h3><span>Latest contact and quotation activity</span></div>
          <div class="activity-list">
            ${
              recent.length
                ? recent
                    .map(
                      (item) => `
                        <div class="activity-row">
                          <div>
                            <strong>${escapeHtml(item.label)}</strong>
                            <p>${escapeHtml(item.detail)}</p>
                          </div>
                          <small>${formatDate(item.date)}</small>
                        </div>
                      `,
                    )
                    .join('')
                : '<p class="empty">No activity recorded yet.</p>'
            }
          </div>
        </article>
      </div>
    </section>
  `
}

function renderLeadDatabase(filteredLeads: Lead[], activeLead: Lead | undefined, countries: string[]): string {
  const duplicate = activeLead ? findDuplicateLead(activeLead.companyName, activeLead.website, activeLead.email, activeLead.id) : undefined
  return `
    <section class="section-stack">
      <div class="two-column">
        <article class="panel">
          <div class="panel-head"><h2>Lead database</h2><span>Control company data better than Excel</span></div>
          <div class="filter-grid">
            <label><span>Search</span><input id="lead-search" type="search" value="${escapeHtml(state.leadSearch)}" placeholder="Company, country, contact, website" /></label>
            <label><span>Country</span><select id="country-filter">${countries
              .map((country) => `<option value="${escapeHtml(country)}" ${state.countryFilter === country ? 'selected' : ''}>${escapeHtml(country)}</option>`)
              .join('')}</select></label>
            <label><span>Status</span><select id="status-filter"><option value="All">All</option>${pipelineStatuses
              .map((status) => `<option value="${escapeHtml(status)}" ${state.statusFilter === status ? 'selected' : ''}>${escapeHtml(status)}</option>`)
              .join('')}</select></label>
          </div>
          <div class="table-shell">
            <table>
              <thead>
                <tr>
                  <th>Company</th>
                  <th>Country</th>
                  <th>Status</th>
                  <th>Last contact</th>
                  <th>Feedback</th>
                </tr>
              </thead>
              <tbody>
                ${
                  filteredLeads.length
                    ? filteredLeads
                        .map(
                          (lead) => `
                            <tr class="${activeLead?.id === lead.id ? 'selected-row' : ''}" data-action="select-lead" data-id="${lead.id}">
                              <td><strong>${escapeHtml(lead.companyName)}</strong><div class="muted">${escapeHtml(lead.sourceOfLead)}</div></td>
                              <td>${escapeHtml(lead.country)}</td>
                              <td><span class="status-pill">${escapeHtml(lead.status)}</span></td>
                              <td>${formatDate(getLastActivityDate(lead))}</td>
                              <td>${escapeHtml(getLastFeedback(lead))}</td>
                            </tr>
                          `,
                        )
                        .join('')
                    : '<tr><td colspan="5"><div class="empty">No leads match the current filter.</div></td></tr>'
                }
              </tbody>
            </table>
          </div>
        </article>

        <article class="panel">
          <div class="panel-head"><h2>Create new customer</h2><span>Stop double sent with duplicate check</span></div>
          <form id="lead-form" class="form-grid">
            <label><span>Company Name</span><input name="companyName" required /></label>
            <label><span>Country</span><input name="country" required /></label>
            <label><span>City</span><input name="city" /></label>
            <label><span>Website</span><input name="website" type="url" /></label>
            <label><span>Contact Person</span><input name="contactPerson" /></label>
            <label><span>Email</span><input name="email" type="email" /></label>
            <label><span>Phone Number</span><input name="phoneNumber" /></label>
            <label><span>Industry Category</span><input name="industryCategory" required /></label>
            <label><span>Source of Lead</span><input name="sourceOfLead" required /></label>
            <label><span>Customer Type</span><select name="customerType"><option>Distributor</option><option>Importer</option><option>Contractor</option><option>Stockist</option><option>Industrial Supplier</option></select></label>
            <label><span>Status</span><select name="status">${pipelineStatuses.map((status) => `<option value="${status}">${status}</option>`).join('')}</select></label>
            <label class="wide"><span>Notes</span><textarea name="notes" rows="3" placeholder="Initial customer note"></textarea></label>
            <div class="form-actions wide"><button class="btn btn-primary" type="submit">Save customer</button></div>
          </form>
        </article>
      </div>

      <article class="panel">
        <div class="panel-head"><h2>Selected customer</h2><span>One place to see company, contact history, quotation, and follow up</span></div>
        ${
          activeLead
            ? `
              <div class="detail-grid">
                <div class="detail-card">
                  <div class="detail-head">
                    <h3>${escapeHtml(activeLead.companyName)}</h3>
                    <span class="status-pill">${escapeHtml(activeLead.status)}</span>
                  </div>
                  <p>${escapeHtml(activeLead.country)} • ${escapeHtml(activeLead.city)} • ${escapeHtml(activeLead.customerType)}</p>
                  <p>${activeLead.website ? `<a href="${escapeHtml(activeLead.website)}" target="_blank">${escapeHtml(activeLead.website)}</a>` : 'No website'}</p>
                  <p>${escapeHtml(activeLead.contactPerson || 'No contact person')} • ${escapeHtml(activeLead.email || 'No email')}</p>
                  <p>${escapeHtml(activeLead.phoneNumber || 'No phone number')}</p>
                  <p>${escapeHtml(activeLead.notes || 'No notes')}</p>
                  <label class="inline-select"><span>Update status</span><select data-role="lead-status" data-id="${activeLead.id}">${pipelineStatuses
                    .map((status) => `<option value="${status}" ${status === activeLead.status ? 'selected' : ''}>${status}</option>`)
                    .join('')}</select></label>
                  ${
                    duplicate
                      ? `<div class="warning-box">Possible duplicate with ${escapeHtml(duplicate.companyName)} • last contact ${formatDate(getLastActivityDate(duplicate))} • status ${escapeHtml(duplicate.status)}</div>`
                      : ''
                  }
                </div>
                <div class="detail-card">
                  <h3>Contact history</h3>
                  ${
                    activeLead.contacts.length
                      ? activeLead.contacts
                          .slice()
                          .sort((a, b) => b.dateContacted.localeCompare(a.dateContacted))
                          .map(
                            (contact) => `
                              <div class="mini-item">
                                <strong>${formatDate(contact.dateContacted)} • ${escapeHtml(contact.method)}</strong>
                                <p>${escapeHtml(contact.subject)}</p>
                                <small>${escapeHtml(contact.responseStatus)}${contact.followUpDate ? ` • Follow up ${formatDate(contact.followUpDate)}` : ''}</small>
                              </div>
                            `,
                          )
                          .join('')
                      : '<p class="empty">No contact history yet.</p>'
                  }
                </div>
                <div class="detail-card">
                  <h3>Quotation history</h3>
                  ${
                    activeLead.quotations.length
                      ? activeLead.quotations
                          .slice()
                          .sort((a, b) => b.dateSent.localeCompare(a.dateSent))
                          .map(
                            (quotation) => `
                              <div class="mini-item">
                                <strong>${escapeHtml(quotation.quotationNumber)} • ${formatCurrency(quotation.quotationValue)}</strong>
                                <p>${escapeHtml(quotation.productType)}</p>
                                <small>Sent ${formatDate(quotation.dateSent)} • Valid ${formatDate(quotation.validityDate)}</small>
                              </div>
                            `,
                          )
                          .join('')
                      : '<p class="empty">No quotation history yet.</p>'
                  }
                </div>
              </div>
            `
            : '<p class="empty">Select a customer from the table.</p>'
        }
      </article>
    </section>
  `
}

function renderContacts(activeLead: Lead | undefined): string {
  const allContacts = flattenContacts(state.data.leads).sort((a, b) => b.dateContacted.localeCompare(a.dateContacted))
  return `
    <section class="section-stack">
      <div class="two-column">
        <article class="panel">
          <div class="panel-head"><h2>Update follow up</h2><span>After creating customer, mostly you only update this section</span></div>
          <form id="contact-form" class="form-grid">
            <label><span>Lead</span><select name="leadId" required><option value="">Select company</option>${state.data.leads
              .map((lead) => `<option value="${lead.id}" ${activeLead?.id === lead.id ? 'selected' : ''}>${escapeHtml(lead.companyName)} • ${escapeHtml(lead.country)}</option>`)
              .join('')}</select></label>
            <label><span>Date Contacted</span><input name="dateContacted" type="date" value="${new Date().toISOString().slice(0, 10)}" required /></label>
            <label><span>Method</span><select name="method">${contactMethods.map((method) => `<option value="${method}">${method}</option>`).join('')}</select></label>
            <label><span>Follow-up Date</span><input name="followUpDate" type="date" /></label>
            <label class="wide"><span>Subject</span><input name="subject" required placeholder="First follow up / quotation reminder / call result" /></label>
            <label class="wide"><span>Message Sent</span><textarea name="messageSent" rows="4" placeholder="Short summary of what you sent or discussed"></textarea></label>
            <label><span>Response Status</span><select name="responseStatus">${responseStatuses.map((item) => `<option value="${item}">${item}</option>`).join('')}</select></label>
            <div class="form-actions wide"><button class="btn btn-primary" type="submit">Save update</button></div>
          </form>
        </article>

        <article class="panel">
          <div class="panel-head"><h2>Quick control</h2><span>Simple daily indicators</span></div>
          <div class="stack-list">
            <div class="summary-card"><strong>Total contact entries</strong><p>${allContacts.length}</p><small>all recorded contact attempts</small></div>
            <div class="summary-card"><strong>Response rate</strong><p>${allContacts.length ? Math.round((allContacts.filter((item) => item.responseStatus === 'Responded').length / allContacts.length) * 100) : 0}%</p><small>based on all contact records</small></div>
            <div class="summary-card"><strong>Email touchpoints</strong><p>${allContacts.filter((item) => item.method === 'Email').length}</p><small>helps track your offered customers</small></div>
          </div>
        </article>
      </div>

      <article class="panel">
        <div class="panel-head"><h2>Contact history</h2><span>All company updates in one place</span></div>
        <div class="table-shell">
          <table>
            <thead>
              <tr>
                <th>Date</th>
                <th>Company</th>
                <th>Method</th>
                <th>Subject</th>
                <th>Follow Up</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              ${
                allContacts.length
                  ? allContacts
                      .map(
                        (contact) => `
                          <tr>
                            <td>${formatDate(contact.dateContacted)}</td>
                            <td><button class="link-btn" data-action="select-lead" data-id="${contact.leadId}">${escapeHtml(contact.companyName)}</button></td>
                            <td>${escapeHtml(contact.method)}</td>
                            <td>${escapeHtml(contact.subject)}</td>
                            <td>${formatDate(contact.followUpDate)}</td>
                            <td>${escapeHtml(contact.responseStatus)}</td>
                          </tr>
                        `,
                      )
                      .join('')
                  : '<tr><td colspan="6"><div class="empty">No contact history yet.</div></td></tr>'
              }
            </tbody>
          </table>
        </div>
      </article>
    </section>
  `
}

function renderQuotations(activeLead: Lead | undefined): string {
  const allQuotations = flattenQuotations(state.data.leads).sort((a, b) => b.dateSent.localeCompare(a.dateSent))
  return `
    <section class="section-stack">
      <div class="two-column">
        <article class="panel">
          <div class="panel-head"><h2>Quotation management</h2><span>Store quotation history so you know which customer was already offered</span></div>
          <form id="quotation-form" class="form-grid">
            <label><span>Lead</span><select name="leadId" required><option value="">Select company</option>${state.data.leads
              .map((lead) => `<option value="${lead.id}" ${activeLead?.id === lead.id ? 'selected' : ''}>${escapeHtml(lead.companyName)} • ${escapeHtml(lead.country)}</option>`)
              .join('')}</select></label>
            <label><span>Quotation Number</span><input name="quotationNumber" required /></label>
            <label><span>Date Sent</span><input name="dateSent" type="date" value="${new Date().toISOString().slice(0, 10)}" required /></label>
            <label><span>Product Type</span><input name="productType" required /></label>
            <label><span>Quotation Value (USD)</span><input name="quotationValue" type="number" min="0" step="0.01" required /></label>
            <label><span>Validity Date</span><input name="validityDate" type="date" required /></label>
            <label class="wide"><span>Customer Feedback</span><textarea name="customerFeedback" rows="3" placeholder="Target price / buyer comment / evaluation status"></textarea></label>
            <div class="form-actions wide"><button class="btn btn-primary" type="submit">Save quotation</button></div>
          </form>
        </article>

        <article class="panel">
          <div class="panel-head"><h2>Quotation snapshot</h2><span>Simple overview</span></div>
          <div class="stack-list">
            <div class="summary-card"><strong>Total quotations</strong><p>${allQuotations.length}</p><small>all quotation records</small></div>
            <div class="summary-card"><strong>Total quoted value</strong><p>${formatCurrency(allQuotations.reduce((sum, item) => sum + item.quotationValue, 0))}</p><small>for quick control</small></div>
            <div class="summary-card"><strong>Quoted customers</strong><p>${state.data.leads.filter((lead) => lead.quotations.length > 0).length}</p><small>who already received quotation</small></div>
          </div>
        </article>
      </div>

      <article class="panel">
        <div class="panel-head"><h2>Quotation history</h2><span>Know which customer already got an offer</span></div>
        <div class="table-shell">
          <table>
            <thead>
              <tr>
                <th>Quotation</th>
                <th>Company</th>
                <th>Date</th>
                <th>Product</th>
                <th>Value</th>
                <th>Feedback</th>
              </tr>
            </thead>
            <tbody>
              ${
                allQuotations.length
                  ? allQuotations
                      .map(
                        (quotation) => `
                          <tr>
                            <td>${escapeHtml(quotation.quotationNumber)}</td>
                            <td><button class="link-btn" data-action="select-lead" data-id="${quotation.leadId}">${escapeHtml(quotation.companyName)}</button></td>
                            <td>${formatDate(quotation.dateSent)}</td>
                            <td>${escapeHtml(quotation.productType)}</td>
                            <td>${formatCurrency(quotation.quotationValue)}</td>
                            <td>${escapeHtml(quotation.customerFeedback || 'No feedback')}</td>
                          </tr>
                        `,
                      )
                      .join('')
                  : '<tr><td colspan="6"><div class="empty">No quotations stored yet.</div></td></tr>'
              }
            </tbody>
          </table>
        </div>
      </article>
    </section>
  `
}

function renderDiscovery(countries: string[], categories: string[], customerTypes: string[]): string {
  const candidates = getFilteredDiscovery()
  return `
    <section class="section-stack">
      <article class="panel">
        <div class="panel-head"><h2>Lead discovery</h2><span>Search new company and block duplicate before importing</span></div>
        <div class="filter-grid">
          <label><span>Search</span><input id="discovery-search" type="search" value="${escapeHtml(state.discoverySearch)}" placeholder="Country, company, keyword" /></label>
          <label><span>Country</span><select id="discovery-country">${countries
            .map((country) => `<option value="${escapeHtml(country)}" ${state.discoveryCountry === country ? 'selected' : ''}>${escapeHtml(country)}</option>`)
            .join('')}</select></label>
          <label><span>Industry</span><select id="discovery-category">${categories
            .map((category) => `<option value="${escapeHtml(category)}" ${state.discoveryCategory === category ? 'selected' : ''}>${escapeHtml(category)}</option>`)
            .join('')}</select></label>
          <label><span>Customer Type</span><select id="discovery-type">${customerTypes
            .map((type) => `<option value="${escapeHtml(type)}" ${state.discoveryType === type ? 'selected' : ''}>${escapeHtml(type)}</option>`)
            .join('')}</select></label>
        </div>
        <div class="card-grid">
          ${
            candidates.length
              ? candidates
                  .map(
                    (candidate) => `
                      <article class="lead-card">
                        <div class="detail-head">
                          <div>
                            <h3>${escapeHtml(candidate.companyName)}</h3>
                            <p>${escapeHtml(candidate.country)} • ${escapeHtml(candidate.customerType)}</p>
                          </div>
                          <span class="status-pill ${candidate.duplicate ? 'warn' : ''}">${candidate.duplicate ? 'Already exists' : 'New prospect'}</span>
                        </div>
                        <p>${escapeHtml(candidate.industryCategory)}</p>
                        <p>${escapeHtml(candidate.email || 'No email')} • ${escapeHtml(candidate.contactPerson || 'No PIC')}</p>
                        <div class="tag-row">${candidate.keywords.map((keyword) => `<span class="tag">${escapeHtml(keyword)}</span>`).join('')}</div>
                        ${
                          candidate.duplicate
                            ? `<div class="warning-box">${escapeHtml(candidate.duplicate.companyName)} already exists • last contact ${formatDate(getLastActivityDate(candidate.duplicate))} • status ${escapeHtml(candidate.duplicate.status)}</div>`
                            : ''
                        }
                        <div class="form-actions">
                          <button class="btn btn-primary" data-action="import-candidate" data-id="${candidate.id}" ${candidate.duplicate ? 'disabled' : ''}>Add to CRM</button>
                        </div>
                      </article>
                    `,
                  )
                  .join('')
              : '<p class="empty">No discovery candidates match the current filters.</p>'
          }
        </div>
      </article>
    </section>
  `
}

function renderAssistant(): string {
  const currentLead = getLeadById(state.assistantLeadId)
  return `
    <section class="section-stack">
      <div class="two-column">
        <article class="panel">
          <div class="panel-head"><h2>AI Assistant</h2><span>Generate quick sales text from selected company data</span></div>
          <form id="assistant-form" class="form-grid">
            <label><span>Lead</span><select name="leadId" required>${state.data.leads
              .map((lead) => `<option value="${lead.id}" ${state.assistantLeadId === lead.id ? 'selected' : ''}>${escapeHtml(lead.companyName)} • ${escapeHtml(lead.country)}</option>`)
              .join('')}</select></label>
            <label><span>Content Type</span><select name="assistantType">${assistantTypes
              .map((type) => `<option value="${type}" ${state.assistantType === type ? 'selected' : ''}>${type}</option>`)
              .join('')}</select></label>
            <label class="wide"><span>Additional Context</span><textarea name="context" rows="5" placeholder="Product focus, pricing note, meeting result, atau complaint detail"></textarea></label>
            <div class="form-actions wide">
              <button class="btn btn-primary" type="submit">Generate draft</button>
              <button class="btn" type="button" data-action="copy-output">Copy output</button>
            </div>
          </form>
        </article>

        <article class="panel">
          <div class="panel-head"><h2>Draft output</h2><span>${escapeHtml(currentLead?.companyName || 'Select a lead')}</span></div>
          <div class="draft-box">
            <label><span>Subject</span><input readonly value="${escapeHtml(state.generatedSubject)}" /></label>
            <label><span>Content</span><textarea readonly rows="16">${escapeHtml(state.generatedOutput)}</textarea></label>
          </div>
        </article>
      </div>
    </section>
  `
}

function bindEvents(): void {
  app.addEventListener('click', (event) => {
    const target = event.target as HTMLElement
    const viewButton = target.closest<HTMLElement>('[data-view]')
    const actionButton = target.closest<HTMLElement>('[data-action]')
    const timeframeButton = target.closest<HTMLElement>('[data-timeframe]')

    if (viewButton?.dataset.view) {
      state.view = viewButton.dataset.view as ViewKey
      state.notice = null
      renderApp()
      return
    }
    if (timeframeButton?.dataset.timeframe) {
      state.timeframe = timeframeButton.dataset.timeframe as ReportWindow
      renderApp()
      return
    }
    if (!actionButton?.dataset.action) return

    const action = actionButton.dataset.action
    const id = actionButton.dataset.id ?? null

    if (action === 'dismiss-notice') {
      state.notice = null
      renderApp()
      return
    }
    if (action === 'select-lead' && id) {
      state.activeLeadId = id
      state.assistantLeadId = id
      state.view = 'leads'
      renderApp()
      return
    }
    if (action === 'load-sample') {
      state.data = createSampleData()
      state.activeLeadId = state.data.leads[0]?.id ?? null
      state.assistantLeadId = state.activeLeadId
      state.generatedSubject = ''
      state.generatedOutput = ''
      persist()
      setNotice('Sample data loaded.', 'success')
      renderApp()
      return
    }
    if (action === 'reset-data') {
      state.data = { leads: [], discovery: createSampleData().discovery, assistantDrafts: [] }
      state.activeLeadId = null
      state.assistantLeadId = null
      state.generatedSubject = ''
      state.generatedOutput = ''
      persist()
      setNotice('CRM data reset.', 'info')
      renderApp()
      return
    }
    if (action === 'import-candidate' && id) {
      importDiscoveryCandidate(id)
      return
    }
    if (action === 'export-excel') {
      exportWorkbook()
      return
    }
    if (action === 'copy-output') {
      void copyGeneratedOutput()
    }
  })

  app.addEventListener('change', (event) => {
    const target = event.target as HTMLInputElement | HTMLSelectElement
    if (target.id === 'country-filter') {
      state.countryFilter = target.value
      renderApp()
      return
    }
    if (target.id === 'status-filter') {
      state.statusFilter = target.value
      renderApp()
      return
    }
    if (target.id === 'discovery-country') {
      state.discoveryCountry = target.value
      renderApp()
      return
    }
    if (target.id === 'discovery-category') {
      state.discoveryCategory = target.value
      renderApp()
      return
    }
    if (target.id === 'discovery-type') {
      state.discoveryType = target.value
      renderApp()
      return
    }
    if (target.dataset.role === 'lead-status' && target instanceof HTMLSelectElement) {
      const lead = getLeadById(target.dataset.id ?? null)
      if (!lead) return
      lead.status = target.value as LeadStatus
      persist()
      setNotice(`Status updated for ${lead.companyName}.`, 'success')
      renderApp()
    }
  })

  app.addEventListener('input', (event) => {
    const target = event.target as HTMLInputElement
    if (target.id === 'lead-search') {
      state.leadSearch = target.value
      renderApp()
      return
    }
    if (target.id === 'discovery-search') {
      state.discoverySearch = target.value
      renderApp()
    }
  })

  app.addEventListener('submit', (event) => {
    event.preventDefault()
    const form = event.target
    if (!(form instanceof HTMLFormElement)) return
    if (form.id === 'lead-form') {
      handleLeadSubmit(form)
      return
    }
    if (form.id === 'contact-form') {
      handleContactSubmit(form)
      return
    }
    if (form.id === 'quotation-form') {
      handleQuotationSubmit(form)
      return
    }
    if (form.id === 'assistant-form') {
      handleAssistantSubmit(form)
    }
  })
}

function handleLeadSubmit(form: HTMLFormElement): void {
  const formData = new FormData(form)
  const companyName = String(formData.get('companyName') || '').trim()
  const website = String(formData.get('website') || '').trim()
  const email = String(formData.get('email') || '').trim()
  const duplicate = findDuplicateLead(companyName, website, email)
  if (duplicate) {
    setNotice(`Stop double sent: ${duplicate.companyName} already exists • last contact ${formatDate(getLastActivityDate(duplicate))} • status ${duplicate.status}`, 'error')
    state.activeLeadId = duplicate.id
    renderApp()
    return
  }

  const lead: Lead = {
    id: createId(),
    companyName,
    country: String(formData.get('country') || '').trim(),
    city: String(formData.get('city') || '').trim(),
    website,
    contactPerson: String(formData.get('contactPerson') || '').trim(),
    email,
    phoneNumber: String(formData.get('phoneNumber') || '').trim(),
    industryCategory: String(formData.get('industryCategory') || '').trim(),
    sourceOfLead: String(formData.get('sourceOfLead') || '').trim(),
    dateAdded: new Date().toISOString().slice(0, 10),
    status: String(formData.get('status') || 'New Lead') as LeadStatus,
    customerType: String(formData.get('customerType') || '').trim(),
    notes: String(formData.get('notes') || '').trim(),
    contacts: [],
    quotations: [],
  }

  state.data.leads.unshift(lead)
  state.activeLeadId = lead.id
  state.assistantLeadId = lead.id
  persist()
  setNotice(`Customer ${lead.companyName} saved.`, 'success')
  renderApp()
}

function handleContactSubmit(form: HTMLFormElement): void {
  const formData = new FormData(form)
  const lead = getLeadById(String(formData.get('leadId') || ''))
  if (!lead) {
    setNotice('Please select a company.', 'error')
    renderApp()
    return
  }

  lead.contacts.unshift({
    id: createId(),
    dateContacted: String(formData.get('dateContacted') || ''),
    method: String(formData.get('method') || 'Email') as ContactMethod,
    subject: String(formData.get('subject') || '').trim(),
    messageSent: String(formData.get('messageSent') || '').trim(),
    followUpDate: String(formData.get('followUpDate') || ''),
    responseStatus: String(formData.get('responseStatus') || 'Pending') as ResponseStatus,
  })

  if (lead.status === 'New Lead') lead.status = 'Contacted'
  if (lead.contacts[0].responseStatus === 'Follow-up Required') lead.status = 'Follow Up'

  state.activeLeadId = lead.id
  state.assistantLeadId = lead.id
  persist()
  setNotice(`Follow up updated for ${lead.companyName}.`, 'success')
  renderApp()
}

function handleQuotationSubmit(form: HTMLFormElement): void {
  const formData = new FormData(form)
  const lead = getLeadById(String(formData.get('leadId') || ''))
  if (!lead) {
    setNotice('Please select a company.', 'error')
    renderApp()
    return
  }

  lead.quotations.unshift({
    id: createId(),
    quotationNumber: String(formData.get('quotationNumber') || '').trim(),
    dateSent: String(formData.get('dateSent') || ''),
    productType: String(formData.get('productType') || '').trim(),
    quotationValue: Number(formData.get('quotationValue') || 0),
    validityDate: String(formData.get('validityDate') || ''),
    customerFeedback: String(formData.get('customerFeedback') || '').trim(),
  })
  lead.status = 'Quotation Sent'
  state.activeLeadId = lead.id
  state.assistantLeadId = lead.id
  persist()
  setNotice(`Quotation saved for ${lead.companyName}.`, 'success')
  renderApp()
}

function handleAssistantSubmit(form: HTMLFormElement): void {
  const formData = new FormData(form)
  const lead = getLeadById(String(formData.get('leadId') || ''))
  if (!lead) {
    setNotice('Please choose a lead.', 'error')
    renderApp()
    return
  }
  const type = String(formData.get('assistantType') || 'Introduction Email') as AssistantType
  const context = String(formData.get('context') || '').trim()
  const generated = generateAssistantDraft(type, lead, context)
  state.generatedSubject = generated.subject
  state.generatedOutput = generated.content
  state.assistantType = type
  state.assistantLeadId = lead.id
  state.data.assistantDrafts.unshift({
    id: createId(),
    type,
    leadId: lead.id,
    createdAt: new Date().toISOString().slice(0, 10),
    subject: generated.subject,
    content: generated.content,
  })
  persist()
  setNotice(`${type} generated for ${lead.companyName}.`, 'success')
  renderApp()
}

function generateAssistantDraft(type: AssistantType, lead: Lead, context: string) {
  const latestQuote = lead.quotations[0]
  const productHint = latestQuote?.productType || lead.industryCategory || 'industrial products'
  const extra = context ? `\n\nAdditional context:\n${context}` : ''
  if (type === 'Introduction Email') {
    return {
      subject: `Introduction - ${lead.companyName}`,
      content: `Dear ${lead.contactPerson || 'Team'},\n\nWe would like to introduce our company as a reliable exporter of ${productHint}. We believe our range can support ${lead.companyName} in ${lead.country}.\n\nPlease let us know if you have any current inquiry or sourcing requirement.\n\nBest regards,\n[Your Name]${extra}`,
    }
  }
  if (type === 'Follow-up Email') {
    return {
      subject: `Follow up - ${lead.companyName}`,
      content: `Dear ${lead.contactPerson || 'Team'},\n\nI would like to follow up on our previous message regarding ${productHint}. Please let me know if there is any active requirement we can support.\n\nBest regards,\n[Your Name]${extra}`,
    }
  }
  if (type === 'Quotation Email') {
    return {
      subject: `Quotation ${latestQuote?.quotationNumber || ''} - ${lead.companyName}`,
      content: `Dear ${lead.contactPerson || 'Team'},\n\nPlease find our quotation for ${latestQuote?.productType || productHint}.\n\nQuoted value: ${latestQuote ? formatCurrency(latestQuote.quotationValue) : '[Quoted Value]'}\nValidity: ${latestQuote ? formatDate(latestQuote.validityDate) : '[Validity Date]'}\n\nBest regards,\n[Your Name]${extra}`,
    }
  }
  if (type === 'Complaint Response') {
    return {
      subject: `Response to your concern - ${lead.companyName}`,
      content: `Dear ${lead.contactPerson || 'Team'},\n\nThank you for your feedback. We are reviewing the issue and will respond with corrective action as soon as possible.\n\nBest regards,\n[Your Name]${extra}`,
    }
  }
  return {
    subject: `Meeting summary - ${lead.companyName}`,
    content: `Meeting summary for ${lead.companyName}\nCountry: ${lead.country}\nCurrent status: ${lead.status}\n\nNotes:\n${context || 'Add meeting notes here.'}`,
  }
}

function importDiscoveryCandidate(candidateId: string): void {
  const candidate = state.data.discovery.find((item) => item.id === candidateId)
  if (!candidate) return
  const duplicate = findDuplicateLead(candidate.companyName, candidate.website, candidate.email)
  if (duplicate) {
    state.activeLeadId = duplicate.id
    setNotice(`Stop double sent: ${duplicate.companyName} already exists in database.`, 'error')
    renderApp()
    return
  }
  const newLead: Lead = {
    id: createId(),
    companyName: candidate.companyName,
    country: candidate.country,
    city: candidate.city,
    website: candidate.website,
    contactPerson: candidate.contactPerson,
    email: candidate.email,
    phoneNumber: candidate.phoneNumber,
    industryCategory: candidate.industryCategory,
    sourceOfLead: candidate.sourceOfLead,
    dateAdded: new Date().toISOString().slice(0, 10),
    status: 'New Lead',
    customerType: candidate.customerType,
    notes: `Imported from discovery: ${candidate.keywords.join(', ')}`,
    contacts: [],
    quotations: [],
  }
  state.data.leads.unshift(newLead)
  state.activeLeadId = newLead.id
  persist()
  setNotice(`${candidate.companyName} added to CRM.`, 'success')
  renderApp()
}

function exportWorkbook(): void {
  const workbook = XLSX.utils.book_new()
  const reminders = getReminders(state.data.leads)
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(
      state.data.leads.map((lead) => ({
        Company: lead.companyName,
        Country: lead.country,
        City: lead.city,
        Website: lead.website,
        Contact: lead.contactPerson,
        Email: lead.email,
        Phone: lead.phoneNumber,
        Industry: lead.industryCategory,
        Source: lead.sourceOfLead,
        'Date Added': lead.dateAdded,
        Status: lead.status,
        'Last Contact': getLastActivityDate(lead),
        Feedback: getLastFeedback(lead),
      })),
    ),
    'Leads',
  )
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(
      flattenContacts(state.data.leads).map((contact) => ({
        Company: contact.companyName,
        Date: contact.dateContacted,
        Method: contact.method,
        Subject: contact.subject,
        'Follow Up': contact.followUpDate,
        Status: contact.responseStatus,
      })),
    ),
    'Contacts',
  )
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(
      flattenQuotations(state.data.leads).map((quotation) => ({
        Company: quotation.companyName,
        Number: quotation.quotationNumber,
        Date: quotation.dateSent,
        Product: quotation.productType,
        Value: quotation.quotationValue,
        Validity: quotation.validityDate,
        Feedback: quotation.customerFeedback,
      })),
    ),
    'Quotations',
  )
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.json_to_sheet(
      reminders.map((reminder) => ({
        Company: reminder.leadName,
        Type: reminder.type,
        Summary: reminder.summary,
        'Due Date': reminder.dueDate,
        Priority: reminder.priority,
      })),
    ),
    'Reminders',
  )
  XLSX.writeFile(workbook, `crm-export-marketing-${new Date().toISOString().slice(0, 10)}.xlsx`)
  setNotice('Excel exported successfully.', 'success')
  renderApp()
}

async function copyGeneratedOutput(): Promise<void> {
  const output = `${state.generatedSubject}\n\n${state.generatedOutput}`.trim()
  if (!output) {
    setNotice('Generate a draft first.', 'error')
    renderApp()
    return
  }
  await navigator.clipboard.writeText(output)
  setNotice('Draft copied to clipboard.', 'success')
  renderApp()
}
