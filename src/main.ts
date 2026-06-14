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
type UiLanguage = 'id' | 'en'
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
  uiLanguage: UiLanguage
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
const uiLanguageKey = 'crm-export-marketing-ui-language'
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
const translations: Record<UiLanguage, Record<string, string>> = {
  id: {
    all: 'Semua',
    exportWorkspace: 'Export workspace',
    sidebarCopy: 'Satu tempat untuk cek company, histori penawaran, duplicate risk, dan prioritas follow up harian.',
    quickPulse: 'Quick pulse',
    noUrgentReminder: 'Tidak ada reminder mendesak',
    needsAttention: 'perlu perhatian',
    databaseSize: 'Database size',
    countries: 'negara',
    sources: 'source',
    language: 'Bahasa',
    search: 'Cari',
    searchGlobalPlaceholder: 'Cari company / email / website',
    newCustomer: '+ Customer Baru',
    updateFollowUpShort: '+ Update Follow Up',
    export: 'Export',
    quickOverview: 'Overview cepat',
    heroTitle: 'UI baru difokuskan untuk scan cepat, update cepat, dan kontrol customer lebih kuat dari Excel.',
    heroCopy:
      'Semua informasi penting dibuat lebih mudah dibaca: company coverage, reminder, quotation, dan histori customer dalam satu workspace yang konsisten.',
    reloadSample: 'Muat ulang sample',
    resetData: 'Reset data',
    totalCompany: 'Total company',
    countryCoverage: 'cakupan negara',
    dueFollowUp: 'Jatuh tempo follow up',
    highPriority: 'prioritas tinggi',
    quotationSent: 'Quotation terkirim',
    untouchedDiscovery: 'Discovery belum dipakai',
    potentialNewCustomer: 'calon customer baru',
    dashboard: 'Dashboard',
    leadDatabase: 'Lead Database',
    contactHistory: 'Riwayat Kontak',
    quotationManagement: 'Manajemen Quotation',
    leadDiscovery: 'Lead Discovery',
    aiAssistant: 'AI Assistant',
    dashboardTitle: 'Dashboard kerja harian',
    dashboardDesc: 'Mulai hari dengan melihat overview company, prioritas follow up, dan pergerakan terbaru.',
    leadsTitle: 'Lead database',
    leadsDesc: 'Cari company dengan cepat, cek duplicate, dan buka workspace customer tanpa rasa seperti Excel.',
    contactsTitle: 'Riwayat kontak',
    contactsDesc: 'Update follow up dan pantau semua histori komunikasi dari satu tempat.',
    quotationsTitle: 'Manajemen quotation',
    quotationsDesc: 'Simpan semua penawaran yang pernah dikirim agar kamu tahu customer mana yang sudah pernah di-offer.',
    discoveryTitle: 'Lead discovery',
    discoveryDesc: 'Temukan prospect baru lalu blokir duplicate sebelum masuk ke database.',
    assistantTitle: 'AI assistant',
    assistantDesc: 'Buat draft email atau ringkasan dengan cepat dari data customer yang sudah ada.',
    dashboardEasyScan: 'Dashboard yang lebih mudah discan',
    dashboardEasyScanDesc: 'Informasi dipecah menjadi empat area: angka ringkas, market overview, fokus hari ini, dan movement terbaru.',
    activityPeriod: 'Aktivitas periode ini',
    activityPeriodDesc: 'Email, response, dan quotation di periode aktif',
    newLeads: 'Lead baru',
    newLeadsDesc: 'company baru yang masuk database',
    responses: 'Respons',
    responsesDesc: 'customer yang sudah memberi jawaban',
    ordersWon: 'Order masuk',
    ordersWonDesc: 'customer yang berhasil closing',
    databaseOverview: 'Overview database',
    databaseOverviewDesc: 'Negara, source, dan status',
    leadsByCountry: 'Lead per negara',
    leadsBySource: 'Lead per source',
    statusPipeline: 'Status pipeline',
    priorityCustomers: 'Prioritas customer',
    priorityCustomersDesc: 'Customer yang perlu kamu lihat lebih dulu',
    noPriorityCustomers: 'Belum ada customer prioritas.',
    reminderSystem: 'Sistem reminder',
    reminderSystemDesc: 'Due today, overdue, dan inactive',
    noReminder: 'Tidak ada reminder saat ini.',
    recentMovements: 'Pergerakan terbaru',
    recentMovementsDesc: 'Apa yang terakhir berubah di database customer',
    noActivity: 'Belum ada aktivitas tercatat.',
    leadDatabaseFeels: 'Lead database yang terasa seperti app, bukan spreadsheet',
    leadDatabaseFeelsDesc: 'Cari cepat, filter cepat, lalu buka workspace customer tanpa baca tabel terlalu lama.',
    createNewCustomer: 'Buat customer baru',
    createNewCustomerDesc: 'Form dibuat lebih ringkas agar cepat dipakai berulang kali',
    companyName: 'Nama Company',
    country: 'Negara',
    city: 'Kota',
    website: 'Website',
    contactPerson: 'Contact Person',
    email: 'Email',
    phoneNumber: 'Nomor Telepon',
    industryCategory: 'Kategori Industri',
    sourceOfLead: 'Sumber Lead',
    customerType: 'Tipe Customer',
    status: 'Status',
    notes: 'Catatan',
    initialCustomerNote: 'Catatan awal customer',
    saveCustomer: 'Simpan customer',
    quickUxIdea: 'Ide UX cepat',
    quickUxIdeaDesc: 'Yang bikin aplikasi terasa ringan',
    searchAlwaysVisible: 'Search selalu terlihat',
    fasterDuplicateCheck: 'lebih cepat cek duplicate',
    leadRowAsCard: 'Lead row jadi card',
    easierToScan: 'lebih mudah discan',
    customerWorkspace: 'Workspace customer',
    updateWithoutSwitching: 'update tanpa pindah-pindah',
    customerWorkspaceTitle: 'Customer workspace',
    customerWorkspaceDesc: 'Semua informasi customer penting dikumpulkan di area kerja yang lebih nyaman dilihat',
    selectedCustomer: 'Customer terpilih',
    source: 'Source',
    lastContact: 'Kontak terakhir',
    feedback: 'Feedback',
    noWebsite: 'Belum ada website',
    noContactPerson: 'Belum ada PIC',
    noEmail: 'Belum ada email',
    noPhoneNumber: 'Belum ada nomor telepon',
    noNotes: 'Belum ada catatan',
    updateStatus: 'Update status',
    possibleDuplicateWith: 'Kemungkinan duplicate dengan',
    contactTimeline: 'Timeline kontak',
    latestActivity: 'Aktivitas terbaru',
    quotationTimeline: 'Timeline quotation',
    offersSentBefore: 'Offer yang pernah dikirim',
    noContactHistoryYet: 'Belum ada histori kontak.',
    noQuotationHistoryYet: 'Belum ada histori quotation.',
    selectCustomerFromTable: 'Pilih customer dari daftar.',
    updateFollowUpTitle: 'Update follow up',
    updateFollowUpDesc: 'Setelah membuat customer, bagian ini yang paling sering kamu update',
    lead: 'Lead',
    selectCompany: 'Pilih company',
    dateContacted: 'Tanggal kontak',
    method: 'Metode',
    followUpDate: 'Tanggal follow up',
    subject: 'Subjek',
    subjectPlaceholder: 'Follow up pertama / reminder quotation / hasil call',
    messageSent: 'Pesan yang dikirim',
    messageSentPlaceholder: 'Ringkasan singkat apa yang kamu kirim atau diskusikan',
    responseStatus: 'Status respons',
    saveUpdate: 'Simpan update',
    quickControl: 'Kontrol cepat',
    quickControlDesc: 'Indikator harian sederhana',
    totalContactEntries: 'Total entry kontak',
    allRecordedContactAttempts: 'semua percobaan kontak yang terekam',
    responseRate: 'Rasio respons',
    basedOnAllContactRecords: 'berdasarkan semua catatan kontak',
    emailTouchpoints: 'Touchpoint email',
    helpsTrackOfferedCustomers: 'membantu melacak customer yang sudah di-offer',
    contactHistoryTitle: 'Riwayat kontak',
    contactHistoryDesc: 'Semua update company dalam satu tempat',
    date: 'Tanggal',
    company: 'Company',
    followUp: 'Follow Up',
    noContactHistory: 'Belum ada histori kontak.',
    quotationManagementTitle: 'Manajemen quotation',
    quotationManagementDesc: 'Simpan histori quotation supaya kamu tahu customer mana yang sudah pernah di-offer',
    quotationNumber: 'Nomor Quotation',
    dateSent: 'Tanggal kirim',
    productType: 'Tipe produk',
    quotationValueUsd: 'Nilai quotation (USD)',
    validityDate: 'Tanggal berlaku',
    customerFeedback: 'Feedback customer',
    customerFeedbackPlaceholder: 'Target harga / komentar buyer / status evaluasi',
    saveQuotation: 'Simpan quotation',
    quotationSnapshot: 'Snapshot quotation',
    quotationSnapshotDesc: 'Overview sederhana',
    totalQuotations: 'Total quotation',
    allQuotationRecords: 'semua catatan quotation',
    totalQuotedValue: 'Total nilai quotation',
    forQuickControl: 'untuk kontrol cepat',
    quotedCustomers: 'Customer yang di-quote',
    alreadyReceivedQuotation: 'yang sudah menerima quotation',
    quotationHistory: 'Riwayat quotation',
    quotationHistoryDesc: 'Ketahui customer yang sudah mendapat penawaran',
    quotation: 'Quotation',
    product: 'Produk',
    value: 'Nilai',
    noFeedback: 'Belum ada feedback',
    noQuotationsStored: 'Belum ada quotation tersimpan.',
    leadDiscoveryTitle: 'Lead discovery',
    leadDiscoveryDesc: 'Cari company baru dan blokir duplicate sebelum import',
    discoverySearchPlaceholder: 'Negara, company, keyword',
    industry: 'Industri',
    alreadyExists: 'Sudah ada',
    newProspect: 'Prospek baru',
    noPic: 'Belum ada PIC',
    addToCrm: 'Tambah ke CRM',
    noDiscoveryCandidates: 'Tidak ada kandidat discovery yang cocok dengan filter saat ini.',
    aiAssistantTitle: 'AI Assistant',
    aiAssistantDesc: 'Buat teks sales cepat dari data company yang dipilih',
    contentType: 'Tipe konten',
    additionalContext: 'Konteks tambahan',
    additionalContextPlaceholder: 'Fokus produk, catatan harga, hasil meeting, atau detail complaint',
    generateDraft: 'Buat draft',
    copyOutput: 'Salin output',
    draftOutput: 'Hasil draft',
    selectLead: 'Pilih lead',
    content: 'Konten',
    daily: 'Harian',
    weekly: 'Mingguan',
    monthly: 'Bulanan',
  },
  en: {
    all: 'All',
    exportWorkspace: 'Export workspace',
    sidebarCopy: 'One place to check companies, quotation history, duplicate risk, and daily follow-up priorities.',
    quickPulse: 'Quick pulse',
    noUrgentReminder: 'No urgent reminders',
    needsAttention: 'needs attention',
    databaseSize: 'Database size',
    countries: 'countries',
    sources: 'sources',
    language: 'Language',
    search: 'Search',
    searchGlobalPlaceholder: 'Search company / email / website',
    newCustomer: '+ New Customer',
    updateFollowUpShort: '+ Update Follow Up',
    export: 'Export',
    quickOverview: 'Quick overview',
    heroTitle: 'The new UI focuses on faster scanning, faster updates, and stronger customer control than Excel.',
    heroCopy:
      'Important information is easier to read: company coverage, reminders, quotations, and customer history are gathered in one consistent workspace.',
    reloadSample: 'Reload sample',
    resetData: 'Reset data',
    totalCompany: 'Total companies',
    countryCoverage: 'country coverage',
    dueFollowUp: 'Due follow-ups',
    highPriority: 'high priority',
    quotationSent: 'Quotations sent',
    untouchedDiscovery: 'Untouched discovery',
    potentialNewCustomer: 'potential new customer',
    dashboard: 'Dashboard',
    leadDatabase: 'Lead Database',
    contactHistory: 'Contact History',
    quotationManagement: 'Quotation Management',
    leadDiscovery: 'Lead Discovery',
    aiAssistant: 'AI Assistant',
    dashboardTitle: 'Daily work dashboard',
    dashboardDesc: 'Start the day with a company overview, follow-up priorities, and recent movements.',
    leadsTitle: 'Lead database',
    leadsDesc: 'Search companies quickly, check duplicates, and open the customer workspace without feeling like Excel.',
    contactsTitle: 'Contact history',
    contactsDesc: 'Update follow-ups and monitor all communication history from one place.',
    quotationsTitle: 'Quotation management',
    quotationsDesc: 'Store every quotation sent so you always know which customers were already offered.',
    discoveryTitle: 'Lead discovery',
    discoveryDesc: 'Find new prospects and block duplicates before adding them to the database.',
    assistantTitle: 'AI assistant',
    assistantDesc: 'Create quick drafts or summaries from existing customer data.',
    dashboardEasyScan: 'A dashboard that is easier to scan',
    dashboardEasyScanDesc: 'Information is split into four areas: summary numbers, market overview, today’s focus, and recent movement.',
    activityPeriod: 'Activity in this period',
    activityPeriodDesc: 'Emails, responses, and quotations in the active period',
    newLeads: 'New leads',
    newLeadsDesc: 'new companies added to the database',
    responses: 'Responses',
    responsesDesc: 'customers who have replied',
    ordersWon: 'Orders won',
    ordersWonDesc: 'customers successfully closed',
    databaseOverview: 'Database overview',
    databaseOverviewDesc: 'Countries, sources, and statuses',
    leadsByCountry: 'Leads by country',
    leadsBySource: 'Leads by source',
    statusPipeline: 'Pipeline status',
    priorityCustomers: 'Priority customers',
    priorityCustomersDesc: 'Customers you should review first',
    noPriorityCustomers: 'No priority customers yet.',
    reminderSystem: 'Reminder system',
    reminderSystemDesc: 'Due today, overdue, and inactive',
    noReminder: 'No reminders at the moment.',
    recentMovements: 'Recent movements',
    recentMovementsDesc: 'What changed most recently in the customer database',
    noActivity: 'No activity recorded yet.',
    leadDatabaseFeels: 'A lead database that feels like an app, not a spreadsheet',
    leadDatabaseFeelsDesc: 'Search quickly, filter quickly, then open the customer workspace without staring at tables too long.',
    createNewCustomer: 'Create new customer',
    createNewCustomerDesc: 'The form is more compact so it can be used repeatedly with less effort',
    companyName: 'Company Name',
    country: 'Country',
    city: 'City',
    website: 'Website',
    contactPerson: 'Contact Person',
    email: 'Email',
    phoneNumber: 'Phone Number',
    industryCategory: 'Industry Category',
    sourceOfLead: 'Lead Source',
    customerType: 'Customer Type',
    status: 'Status',
    notes: 'Notes',
    initialCustomerNote: 'Initial customer note',
    saveCustomer: 'Save customer',
    quickUxIdea: 'Quick UX idea',
    quickUxIdeaDesc: 'What keeps the app feeling light',
    searchAlwaysVisible: 'Search always visible',
    fasterDuplicateCheck: 'faster duplicate checking',
    leadRowAsCard: 'Lead row becomes a card',
    easierToScan: 'easier to scan',
    customerWorkspace: 'Customer workspace',
    updateWithoutSwitching: 'update without switching around',
    customerWorkspaceTitle: 'Customer workspace',
    customerWorkspaceDesc: 'All important customer information is collected in a workspace that is easier to review',
    selectedCustomer: 'Selected customer',
    source: 'Source',
    lastContact: 'Last contact',
    feedback: 'Feedback',
    noWebsite: 'No website',
    noContactPerson: 'No contact person',
    noEmail: 'No email',
    noPhoneNumber: 'No phone number',
    noNotes: 'No notes',
    updateStatus: 'Update status',
    possibleDuplicateWith: 'Possible duplicate with',
    contactTimeline: 'Contact timeline',
    latestActivity: 'Latest activity',
    quotationTimeline: 'Quotation timeline',
    offersSentBefore: 'Offers sent before',
    noContactHistoryYet: 'No contact history yet.',
    noQuotationHistoryYet: 'No quotation history yet.',
    selectCustomerFromTable: 'Select a customer from the list.',
    updateFollowUpTitle: 'Update follow-up',
    updateFollowUpDesc: 'After creating a customer, this is the section you will update most often',
    lead: 'Lead',
    selectCompany: 'Select company',
    dateContacted: 'Date Contacted',
    method: 'Method',
    followUpDate: 'Follow-up Date',
    subject: 'Subject',
    subjectPlaceholder: 'First follow-up / quotation reminder / call result',
    messageSent: 'Message Sent',
    messageSentPlaceholder: 'A short summary of what you sent or discussed',
    responseStatus: 'Response Status',
    saveUpdate: 'Save update',
    quickControl: 'Quick control',
    quickControlDesc: 'Simple daily indicators',
    totalContactEntries: 'Total contact entries',
    allRecordedContactAttempts: 'all recorded contact attempts',
    responseRate: 'Response rate',
    basedOnAllContactRecords: 'based on all contact records',
    emailTouchpoints: 'Email touchpoints',
    helpsTrackOfferedCustomers: 'helps track offered customers',
    contactHistoryTitle: 'Contact history',
    contactHistoryDesc: 'All company updates in one place',
    date: 'Date',
    company: 'Company',
    followUp: 'Follow-up',
    noContactHistory: 'No contact history yet.',
    quotationManagementTitle: 'Quotation management',
    quotationManagementDesc: 'Store quotation history so you know which customer was already offered',
    quotationNumber: 'Quotation Number',
    dateSent: 'Date Sent',
    productType: 'Product Type',
    quotationValueUsd: 'Quotation Value (USD)',
    validityDate: 'Validity Date',
    customerFeedback: 'Customer Feedback',
    customerFeedbackPlaceholder: 'Target price / buyer comment / evaluation status',
    saveQuotation: 'Save quotation',
    quotationSnapshot: 'Quotation snapshot',
    quotationSnapshotDesc: 'Simple overview',
    totalQuotations: 'Total quotations',
    allQuotationRecords: 'all quotation records',
    totalQuotedValue: 'Total quoted value',
    forQuickControl: 'for quick control',
    quotedCustomers: 'Quoted customers',
    alreadyReceivedQuotation: 'who already received quotations',
    quotationHistory: 'Quotation history',
    quotationHistoryDesc: 'Know which customers already received an offer',
    quotation: 'Quotation',
    product: 'Product',
    value: 'Value',
    noFeedback: 'No feedback',
    noQuotationsStored: 'No quotations stored yet.',
    leadDiscoveryTitle: 'Lead discovery',
    leadDiscoveryDesc: 'Search new companies and block duplicates before importing',
    discoverySearchPlaceholder: 'Country, company, keyword',
    industry: 'Industry',
    alreadyExists: 'Already exists',
    newProspect: 'New prospect',
    noPic: 'No PIC',
    addToCrm: 'Add to CRM',
    noDiscoveryCandidates: 'No discovery candidates match the current filters.',
    aiAssistantTitle: 'AI Assistant',
    aiAssistantDesc: 'Generate quick sales text from selected company data',
    contentType: 'Content Type',
    additionalContext: 'Additional Context',
    additionalContextPlaceholder: 'Product focus, pricing note, meeting result, or complaint detail',
    generateDraft: 'Generate draft',
    copyOutput: 'Copy output',
    draftOutput: 'Draft output',
    selectLead: 'Select lead',
    content: 'Content',
    daily: 'Daily',
    weekly: 'Weekly',
    monthly: 'Monthly',
  },
}

const appRoot = document.querySelector<HTMLDivElement>('#app')
if (!appRoot) throw new Error('App container was not found.')
const app: HTMLDivElement = appRoot

document.title = 'CRM Export Marketing'

const state: AppState = {
  data: loadData(),
  uiLanguage: loadUiLanguage(),
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
document.documentElement.lang = state.uiLanguage

renderApp()
bindEvents()

function loadUiLanguage(): UiLanguage {
  return localStorage.getItem(uiLanguageKey) === 'en' ? 'en' : 'id'
}

function persistUiLanguage(): void {
  localStorage.setItem(uiLanguageKey, state.uiLanguage)
  document.documentElement.lang = state.uiLanguage
}

function t(key: string): string {
  return translations[state.uiLanguage][key] ?? key
}

function translateLeadStatus(status: LeadStatus): string {
  if (state.uiLanguage === 'en') return status
  const map: Record<LeadStatus, string> = {
    'New Lead': 'Lead Baru',
    Contacted: 'Sudah Dihubungi',
    'Follow Up': 'Follow Up',
    'Quotation Sent': 'Quotation Terkirim',
    Negotiation: 'Negosiasi',
    'Sample Stage': 'Tahap Sample',
    'Awaiting Order': 'Menunggu Order',
    'Order Received': 'Order Diterima',
    'Lost Opportunity': 'Peluang Hilang',
  }
  return map[status]
}

function translateContactMethod(method: ContactMethod): string {
  if (state.uiLanguage === 'en') return method
  const map: Record<ContactMethod, string> = {
    Email: 'Email',
    LinkedIn: 'LinkedIn',
    WhatsApp: 'WhatsApp',
    Phone: 'Telepon',
  }
  return map[method]
}

function translateResponseStatus(status: ResponseStatus): string {
  if (state.uiLanguage === 'en') return status
  const map: Record<ResponseStatus, string> = {
    Pending: 'Menunggu',
    Responded: 'Merespons',
    'No Response': 'Belum Merespons',
    'Follow-up Required': 'Perlu Follow Up',
  }
  return map[status]
}

function translateAssistantType(type: AssistantType): string {
  if (state.uiLanguage === 'en') return type
  const map: Record<AssistantType, string> = {
    'Introduction Email': 'Email Perkenalan',
    'Follow-up Email': 'Email Follow Up',
    'Quotation Email': 'Email Quotation',
    'Complaint Response': 'Balasan Komplain',
    'Meeting Summary': 'Ringkasan Meeting',
  }
  return map[type]
}

function translateCustomerType(type: string): string {
  if (state.uiLanguage === 'en') return type
  const map: Record<string, string> = {
    Distributor: 'Distributor',
    Importer: 'Importir',
    Contractor: 'Kontraktor',
    Stockist: 'Stockist',
    'Industrial Supplier': 'Pemasok Industri',
  }
  return map[type] ?? type
}

function translateReminderSummary(reminder: Reminder): string {
  if (state.uiLanguage === 'id') return reminder.summary
  if (reminder.type === 'Follow Up') return `${reminder.leadName} has not replied since the first introduction.`
  if (reminder.type === 'Overdue Follow Up') return `${reminder.leadName} is already past the scheduled follow-up date.`
  if (reminder.type === 'Inactive') return `${reminder.leadName} has no activity for 30+ days.`
  return reminder.summary
}

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
  if (!date) return state.uiLanguage === 'id' ? 'Belum ada' : 'Not available'
  return new Intl.DateTimeFormat(state.uiLanguage === 'id' ? 'id-ID' : 'en-US', { day: '2-digit', month: 'short', year: 'numeric' }).format(new Date(date))
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
  return lead.quotations[0]?.customerFeedback || lead.contacts.find((item) => item.responseStatus === 'Responded')?.subject || (state.uiLanguage === 'id' ? 'Belum ada' : 'No feedback yet')
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

function getViewMeta(view: ViewKey): { title: string; description: string } {
  const map: Record<ViewKey, { title: string; description: string }> = {
    dashboard: {
      title: t('dashboardTitle'),
      description: t('dashboardDesc'),
    },
    leads: {
      title: t('leadsTitle'),
      description: t('leadsDesc'),
    },
    contacts: {
      title: t('contactsTitle'),
      description: t('contactsDesc'),
    },
    quotations: {
      title: t('quotationsTitle'),
      description: t('quotationsDesc'),
    },
    discovery: {
      title: t('discoveryTitle'),
      description: t('discoveryDesc'),
    },
    assistant: {
      title: t('assistantTitle'),
      description: t('assistantDesc'),
    },
  }
  return map[view]
}

function getPriorityLeads(): Lead[] {
  return [...state.data.leads]
    .sort((a, b) => {
      const aScore =
        (a.status === 'Quotation Sent' ? 4 : 0) +
        (a.status === 'Negotiation' ? 5 : 0) +
        (a.status === 'Follow Up' ? 3 : 0) +
        (a.contacts[0]?.responseStatus === 'Follow-up Required' ? 2 : 0)
      const bScore =
        (b.status === 'Quotation Sent' ? 4 : 0) +
        (b.status === 'Negotiation' ? 5 : 0) +
        (b.status === 'Follow Up' ? 3 : 0) +
        (b.contacts[0]?.responseStatus === 'Follow-up Required' ? 2 : 0)
      return bScore - aScore || getLastActivityDate(b).localeCompare(getLastActivityDate(a))
    })
    .slice(0, 5)
}

function renderApp(): void {
  const filteredLeads = getFilteredLeads()
  const activeLead = getLeadById(state.activeLeadId) ?? filteredLeads[0] ?? state.data.leads[0]
  const reminders = getReminders(state.data.leads)
  const metrics = getWindowMetrics()
  const viewMeta = getViewMeta(state.view)
  const topReminder = reminders[0]
  const countries = ['All', ...Object.keys(getCountryCounts(state.data.leads)).sort()]
  const discoveryCountries = ['All', ...new Set(state.data.discovery.map((item) => item.country)).values()]
  const discoveryCategories = ['All', ...new Set(state.data.discovery.map((item) => item.industryCategory)).values()]
  const discoveryTypes = ['All', ...new Set(state.data.discovery.map((item) => item.customerType)).values()]

  if (activeLead) {
    state.activeLeadId = activeLead.id
    if (!state.assistantLeadId) state.assistantLeadId = activeLead.id
  }

  app.innerHTML = `
    <div class="app-shell">
      <aside class="sidebar">
        <div class="sidebar-head">
          <div class="brand-mark">CE</div>
          <div>
            <p class="eyebrow">${t('exportWorkspace')}</p>
            <h1>CRM Export Marketing</h1>
          </div>
        </div>
        <p class="sidebar-copy">${t('sidebarCopy')}</p>
        <nav class="nav-stack">${renderTabs()}</nav>
        <section class="sidebar-card">
          <span>${t('quickPulse')}</span>
          <strong>${reminders.length}</strong>
          <small>${topReminder ? `${escapeHtml(topReminder.leadName)} ${t('needsAttention')}` : t('noUrgentReminder')}</small>
        </section>
        <section class="sidebar-card">
          <span>${t('databaseSize')}</span>
          <strong>${state.data.leads.length}</strong>
          <small>${Object.keys(getCountryCounts(state.data.leads)).length} ${t('countries')} • ${Object.keys(getSourceCounts(state.data.leads)).length} ${t('sources')}</small>
        </section>
      </aside>

      <div class="content-shell">
        <header class="topbar">
          <div class="topbar-copy">
            <p class="eyebrow">${escapeHtml(viewMeta.title)}</p>
            <h2>${escapeHtml(viewMeta.description)}</h2>
          </div>
          <div class="topbar-actions">
            <label class="topbar-search">
              <span>${t('search')}</span>
              <input id="global-search" type="search" value="${escapeHtml(state.leadSearch)}" placeholder="${t('searchGlobalPlaceholder')}" />
            </label>
            <label class="language-switch">
              <span>${t('language')}</span>
              <select id="language-switch">
                <option value="id" ${state.uiLanguage === 'id' ? 'selected' : ''}>ID</option>
                <option value="en" ${state.uiLanguage === 'en' ? 'selected' : ''}>EN</option>
              </select>
            </label>
            <button class="btn btn-primary" data-view="leads">${t('newCustomer')}</button>
            <button class="btn" data-view="contacts">${t('updateFollowUpShort')}</button>
            <button class="btn" data-action="export-excel">${t('export')}</button>
          </div>
        </header>

        <section class="hero-band">
          <div class="hero-intro">
            <p class="eyebrow">${t('quickOverview')}</p>
            <h3>${t('heroTitle')}</h3>
            <p class="hero-copy">${t('heroCopy')}</p>
          </div>
          <div class="hero-actions">
            <button class="btn" data-action="load-sample">${t('reloadSample')}</button>
            <button class="btn btn-danger" data-action="reset-data">${t('resetData')}</button>
          </div>
        </section>

        <section class="hero-metrics">
          <article class="highlight-card accent-blue"><span>${t('totalCompany')}</span><strong>${state.data.leads.length}</strong><small>${Object.keys(getCountryCounts(state.data.leads)).length} ${t('countryCoverage')}</small></article>
          <article class="highlight-card accent-orange"><span>${t('dueFollowUp')}</span><strong>${reminders.length}</strong><small>${reminders.filter((item) => item.priority === 'high').length} ${t('highPriority')}</small></article>
          <article class="highlight-card accent-violet"><span>${t('quotationSent')}</span><strong>${flattenQuotations(state.data.leads).length}</strong><small>${formatCurrency(flattenQuotations(state.data.leads).reduce((sum, item) => sum + item.quotationValue, 0))}</small></article>
          <article class="highlight-card accent-green"><span>${t('untouchedDiscovery')}</span><strong>${getFilteredDiscovery().filter((item) => !item.duplicate).length}</strong><small>${t('potentialNewCustomer')}</small></article>
        </section>

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
    </div>
  `
}

function renderTabs(): string {
  const tabs: Array<{ key: ViewKey; label: string }> = [
    { key: 'dashboard', label: t('dashboard') },
    { key: 'leads', label: t('leadDatabase') },
    { key: 'contacts', label: t('contactHistory') },
    { key: 'quotations', label: t('quotationManagement') },
    { key: 'discovery', label: t('leadDiscovery') },
    { key: 'assistant', label: t('aiAssistant') },
  ]
  return tabs
    .map(
      (tab) => `
        <button class="tab ${state.view === tab.key ? 'active' : ''}" data-view="${tab.key}">
          <span class="tab-dot"></span>
          <span>${tab.label}</span>
        </button>
      `,
    )
    .join('')
}

function renderDashboard(metrics: ReturnType<typeof getWindowMetrics>, reminders: Reminder[]): string {
  const countryCounts = Object.entries(getCountryCounts(state.data.leads)).sort((a, b) => b[1] - a[1]).slice(0, 6)
  const sourceCounts = Object.entries(getSourceCounts(state.data.leads)).sort((a, b) => b[1] - a[1]).slice(0, 6)
  const pipelineCounts = getPipelineCounts(state.data.leads)
  const recent = getRecentActivity()
  const priorityLeads = getPriorityLeads()
  return `
    <section class="section-stack">
      <div class="section-head">
        <div>
          <h2>${t('dashboardEasyScan')}</h2>
          <p>${t('dashboardEasyScanDesc')}</p>
        </div>
        <div class="segment-control">
          <button data-timeframe="daily" class="${state.timeframe === 'daily' ? 'active' : ''}">${t('daily')}</button>
          <button data-timeframe="weekly" class="${state.timeframe === 'weekly' ? 'active' : ''}">${t('weekly')}</button>
          <button data-timeframe="monthly" class="${state.timeframe === 'monthly' ? 'active' : ''}">${t('monthly')}</button>
        </div>
      </div>

      <div class="focus-grid">
        <article class="metric-card featured-metric">
          <span>${t('activityPeriod')}</span>
          <strong>${metrics.emailsSent + metrics.responsesReceived + metrics.quotationsSent}</strong>
          <small>${t('activityPeriodDesc')}</small>
        </article>
        <article class="metric-card"><span>${t('newLeads')}</span><strong>${metrics.newLeads}</strong><small>${t('newLeadsDesc')}</small></article>
        <article class="metric-card"><span>${t('responses')}</span><strong>${metrics.responsesReceived}</strong><small>${t('responsesDesc')}</small></article>
        <article class="metric-card"><span>${t('ordersWon')}</span><strong>${metrics.ordersWon}</strong><small>${t('ordersWonDesc')}</small></article>
      </div>

      <div class="dashboard-grid redesigned-grid">
        <article class="panel market-panel">
          <div class="panel-head"><h3>${t('databaseOverview')}</h3><span>${t('databaseOverviewDesc')}</span></div>
          <div class="overview-split">
            <div class="stack-list">
              <h4>${t('leadsByCountry')}</h4>
              ${countryCounts.map(([country, count]) => `<div class="row-line"><span>${escapeHtml(country)}</span><strong>${count}</strong></div>`).join('')}
            </div>
            <div class="stack-list">
              <h4>${t('leadsBySource')}</h4>
              ${sourceCounts.map(([source, count]) => `<div class="row-line"><span>${escapeHtml(source)}</span><strong>${count}</strong></div>`).join('')}
            </div>
            <div class="stack-list">
              <h4>${t('statusPipeline')}</h4>
              ${pipelineCounts
                .filter((item) => item.count > 0)
                .map((item) => `<div class="row-line"><span>${escapeHtml(translateLeadStatus(item.status))}</span><strong>${item.count}</strong></div>`)
                .join('')}
            </div>
          </div>
        </article>

        <article class="panel">
          <div class="panel-head"><h3>${t('priorityCustomers')}</h3><span>${t('priorityCustomersDesc')}</span></div>
          <div class="priority-list">
            ${
              priorityLeads.length
                ? priorityLeads
                    .map(
                      (lead) => `
                        <button class="priority-card" data-action="select-lead" data-id="${lead.id}">
                          <div>
                            <strong>${escapeHtml(lead.companyName)}</strong>
                            <p>${escapeHtml(lead.country)} • ${escapeHtml(lead.sourceOfLead)}</p>
                          </div>
                          <span class="status-pill">${escapeHtml(translateLeadStatus(lead.status))}</span>
                        </button>
                      `,
                    )
                    .join('')
                : `<p class="empty">${t('noPriorityCustomers')}</p>`
            }
          </div>
        </article>

        <article class="panel">
          <div class="panel-head"><h3>${t('reminderSystem')}</h3><span>${t('reminderSystemDesc')}</span></div>
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
                            <p>${escapeHtml(translateReminderSummary(reminder))}</p>
                          </div>
                          <small>${formatDate(reminder.dueDate)}</small>
                        </button>
                      `,
                    )
                    .join('')
                : `<p class="empty">${t('noReminder')}</p>`
            }
          </div>
        </article>

        <article class="panel full-span">
          <div class="panel-head"><h3>${t('recentMovements')}</h3><span>${t('recentMovementsDesc')}</span></div>
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
                : `<p class="empty">${t('noActivity')}</p>`
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
      <div class="database-layout">
        <article class="panel database-panel">
          <div class="panel-head">
            <div>
              <h2>${t('leadDatabaseFeels')}</h2>
              <span>${t('leadDatabaseFeelsDesc')}</span>
            </div>
            <button class="btn btn-primary" data-action="focus-create">${t('newCustomer')}</button>
          </div>
          <div class="filter-grid redesigned-filter">
            <label><span>${t('search')}</span><input id="lead-search" type="search" value="${escapeHtml(state.leadSearch)}" placeholder="Company, country, contact, website" /></label>
            <label><span>${t('country')}</span><select id="country-filter">${countries
              .map((country) => `<option value="${escapeHtml(country)}" ${state.countryFilter === country ? 'selected' : ''}>${escapeHtml(country === 'All' ? t('all') : country)}</option>`)
              .join('')}</select></label>
            <label><span>${t('status')}</span><select id="status-filter"><option value="All">${t('all')}</option>${pipelineStatuses
              .map((status) => `<option value="${escapeHtml(status)}" ${state.statusFilter === status ? 'selected' : ''}>${escapeHtml(translateLeadStatus(status))}</option>`)
              .join('')}</select></label>
          </div>
          <div class="lead-table">
            ${
              filteredLeads.length
                ? filteredLeads
                    .map(
                      (lead) => `
                        <button class="lead-row ${activeLead?.id === lead.id ? 'selected-row' : ''}" data-action="select-lead" data-id="${lead.id}">
                          <div class="lead-main">
                            <strong>${escapeHtml(lead.companyName)}</strong>
                            <p>${escapeHtml(lead.country)} • ${escapeHtml(lead.sourceOfLead)} • ${escapeHtml(lead.industryCategory)}</p>
                          </div>
                          <div class="lead-meta">
                            <span class="status-pill">${escapeHtml(translateLeadStatus(lead.status))}</span>
                            <small>${t('lastContact')} ${formatDate(getLastActivityDate(lead))}</small>
                            <small>${escapeHtml(getLastFeedback(lead))}</small>
                          </div>
                        </button>
                      `,
                    )
                    .join('')
                : `<div class="empty">${state.uiLanguage === 'id' ? 'Tidak ada lead yang cocok dengan filter saat ini.' : 'No leads match the current filter.'}</div>`
            }
          </div>
        </article>

        <aside class="workspace-side">
          <article class="panel create-panel" id="create-customer-panel">
            <div class="panel-head"><h2>${t('createNewCustomer')}</h2><span>${t('createNewCustomerDesc')}</span></div>
            <form id="lead-form" class="form-grid">
              <label><span>${t('companyName')}</span><input name="companyName" required /></label>
              <label><span>${t('country')}</span><input name="country" required /></label>
              <label><span>${t('city')}</span><input name="city" /></label>
              <label><span>${t('website')}</span><input name="website" type="url" /></label>
              <label><span>${t('contactPerson')}</span><input name="contactPerson" /></label>
              <label><span>${t('email')}</span><input name="email" type="email" /></label>
              <label><span>${t('phoneNumber')}</span><input name="phoneNumber" /></label>
              <label><span>${t('industryCategory')}</span><input name="industryCategory" required /></label>
              <label><span>${t('sourceOfLead')}</span><input name="sourceOfLead" required /></label>
              <label><span>${t('customerType')}</span><select name="customerType"><option>Distributor</option><option>Importer</option><option>Contractor</option><option>Stockist</option><option>Industrial Supplier</option></select></label>
              <label><span>${t('status')}</span><select name="status">${pipelineStatuses.map((status) => `<option value="${status}">${translateLeadStatus(status)}</option>`).join('')}</select></label>
              <label class="wide"><span>${t('notes')}</span><textarea name="notes" rows="3" placeholder="${t('initialCustomerNote')}"></textarea></label>
              <div class="form-actions wide"><button class="btn btn-primary" type="submit">${t('saveCustomer')}</button></div>
            </form>
          </article>

          <article class="panel quick-tip-panel">
            <div class="panel-head"><h3>${t('quickUxIdea')}</h3><span>${t('quickUxIdeaDesc')}</span></div>
            <div class="stack-list">
              <div class="row-line"><span>${t('searchAlwaysVisible')}</span><strong>${t('fasterDuplicateCheck')}</strong></div>
              <div class="row-line"><span>${t('leadRowAsCard')}</span><strong>${t('easierToScan')}</strong></div>
              <div class="row-line"><span>${t('customerWorkspace')}</span><strong>${t('updateWithoutSwitching')}</strong></div>
            </div>
          </article>
        </aside>
      </div>

      <article class="panel workspace-panel">
        <div class="panel-head"><h2>${t('customerWorkspaceTitle')}</h2><span>${t('customerWorkspaceDesc')}</span></div>
        ${
          activeLead
            ? `
              <div class="workspace-grid">
                <div class="workspace-profile">
                  <div class="workspace-hero">
                    <div>
                      <p class="eyebrow">${t('selectedCustomer')}</p>
                      <h3>${escapeHtml(activeLead.companyName)}</h3>
                      <p>${escapeHtml(activeLead.country)} • ${escapeHtml(activeLead.city)} • ${escapeHtml(translateCustomerType(activeLead.customerType))}</p>
                    </div>
                    <span class="status-pill">${escapeHtml(translateLeadStatus(activeLead.status))}</span>
                  </div>
                  <div class="info-pills">
                    <div class="info-pill"><strong>${t('source')}</strong><span>${escapeHtml(activeLead.sourceOfLead)}</span></div>
                    <div class="info-pill"><strong>${t('lastContact')}</strong><span>${formatDate(getLastActivityDate(activeLead))}</span></div>
                    <div class="info-pill"><strong>${t('feedback')}</strong><span>${escapeHtml(getLastFeedback(activeLead))}</span></div>
                  </div>
                  <div class="workspace-contact">
                    <p>${activeLead.website ? `<a href="${escapeHtml(activeLead.website)}" target="_blank">${escapeHtml(activeLead.website)}</a>` : t('noWebsite')}</p>
                    <p>${escapeHtml(activeLead.contactPerson || t('noContactPerson'))} • ${escapeHtml(activeLead.email || t('noEmail'))}</p>
                    <p>${escapeHtml(activeLead.phoneNumber || t('noPhoneNumber'))}</p>
                  </div>
                  <div class="notes-box">${escapeHtml(activeLead.notes || t('noNotes'))}</div>
                  <label class="inline-select status-select-row"><span>${t('updateStatus')}</span><select data-role="lead-status" data-id="${activeLead.id}">${pipelineStatuses
                    .map((status) => `<option value="${status}" ${status === activeLead.status ? 'selected' : ''}>${translateLeadStatus(status)}</option>`)
                    .join('')}</select></label>
                  ${
                    duplicate
                      ? `<div class="warning-box">${t('possibleDuplicateWith')} ${escapeHtml(duplicate.companyName)} • ${t('lastContact').toLowerCase()} ${formatDate(getLastActivityDate(duplicate))} • ${t('status').toLowerCase()} ${escapeHtml(translateLeadStatus(duplicate.status))}</div>`
                      : ''
                  }
                </div>

                <div class="workspace-columns">
                  <div class="detail-card timeline-card">
                    <div class="panel-head compact-head"><h3>${t('contactTimeline')}</h3><span>${t('latestActivity')}</span></div>
                    ${
                      activeLead.contacts.length
                        ? activeLead.contacts
                            .slice()
                            .sort((a, b) => b.dateContacted.localeCompare(a.dateContacted))
                            .map(
                              (contact) => `
                                <div class="timeline-item">
                                  <div class="timeline-head">
                                    <strong>${escapeHtml(translateContactMethod(contact.method))}</strong>
                                    <span>${formatDate(contact.dateContacted)}</span>
                                  </div>
                                  <p>${escapeHtml(contact.subject)}</p>
                                  <small>${escapeHtml(translateResponseStatus(contact.responseStatus))}${contact.followUpDate ? ` • ${t('followUp')} ${formatDate(contact.followUpDate)}` : ''}</small>
                                </div>
                              `,
                            )
                            .join('')
                        : `<p class="empty">${t('noContactHistoryYet')}</p>`
                    }
                  </div>

                  <div class="detail-card timeline-card">
                    <div class="panel-head compact-head"><h3>${t('quotationTimeline')}</h3><span>${t('offersSentBefore')}</span></div>
                    ${
                      activeLead.quotations.length
                        ? activeLead.quotations
                            .slice()
                            .sort((a, b) => b.dateSent.localeCompare(a.dateSent))
                            .map(
                              (quotation) => `
                                <div class="timeline-item">
                                  <div class="timeline-head">
                                    <strong>${escapeHtml(quotation.quotationNumber)}</strong>
                                    <span>${formatDate(quotation.dateSent)}</span>
                                  </div>
                                  <p>${escapeHtml(quotation.productType)} • ${formatCurrency(quotation.quotationValue)}</p>
                                  <small>${t('validityDate')} ${formatDate(quotation.validityDate)} • ${escapeHtml(quotation.customerFeedback || t('noFeedback'))}</small>
                                </div>
                              `,
                            )
                            .join('')
                        : `<p class="empty">${t('noQuotationHistoryYet')}</p>`
                    }
                  </div>
                </div>
              </div>
            `
            : `<p class="empty">${t('selectCustomerFromTable')}</p>`
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
          <div class="panel-head"><h2>${t('updateFollowUpTitle')}</h2><span>${t('updateFollowUpDesc')}</span></div>
          <form id="contact-form" class="form-grid">
            <label><span>${t('lead')}</span><select name="leadId" required><option value="">${t('selectCompany')}</option>${state.data.leads
              .map((lead) => `<option value="${lead.id}" ${activeLead?.id === lead.id ? 'selected' : ''}>${escapeHtml(lead.companyName)} • ${escapeHtml(lead.country)}</option>`)
              .join('')}</select></label>
            <label><span>${t('dateContacted')}</span><input name="dateContacted" type="date" value="${new Date().toISOString().slice(0, 10)}" required /></label>
            <label><span>${t('method')}</span><select name="method">${contactMethods.map((method) => `<option value="${method}">${translateContactMethod(method)}</option>`).join('')}</select></label>
            <label><span>${t('followUpDate')}</span><input name="followUpDate" type="date" /></label>
            <label class="wide"><span>${t('subject')}</span><input name="subject" required placeholder="${t('subjectPlaceholder')}" /></label>
            <label class="wide"><span>${t('messageSent')}</span><textarea name="messageSent" rows="4" placeholder="${t('messageSentPlaceholder')}"></textarea></label>
            <label><span>${t('responseStatus')}</span><select name="responseStatus">${responseStatuses.map((item) => `<option value="${item}">${translateResponseStatus(item)}</option>`).join('')}</select></label>
            <div class="form-actions wide"><button class="btn btn-primary" type="submit">${t('saveUpdate')}</button></div>
          </form>
        </article>

        <article class="panel">
          <div class="panel-head"><h2>${t('quickControl')}</h2><span>${t('quickControlDesc')}</span></div>
          <div class="stack-list">
            <div class="summary-card"><strong>${t('totalContactEntries')}</strong><p>${allContacts.length}</p><small>${t('allRecordedContactAttempts')}</small></div>
            <div class="summary-card"><strong>${t('responseRate')}</strong><p>${allContacts.length ? Math.round((allContacts.filter((item) => item.responseStatus === 'Responded').length / allContacts.length) * 100) : 0}%</p><small>${t('basedOnAllContactRecords')}</small></div>
            <div class="summary-card"><strong>${t('emailTouchpoints')}</strong><p>${allContacts.filter((item) => item.method === 'Email').length}</p><small>${t('helpsTrackOfferedCustomers')}</small></div>
          </div>
        </article>
      </div>

      <article class="panel">
        <div class="panel-head"><h2>${t('contactHistoryTitle')}</h2><span>${t('contactHistoryDesc')}</span></div>
        <div class="table-shell">
          <table>
            <thead>
              <tr>
                <th>${t('date')}</th>
                <th>${t('company')}</th>
                <th>${t('method')}</th>
                <th>${t('subject')}</th>
                <th>${t('followUp')}</th>
                <th>${t('status')}</th>
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
                            <td>${escapeHtml(translateContactMethod(contact.method))}</td>
                            <td>${escapeHtml(contact.subject)}</td>
                            <td>${formatDate(contact.followUpDate)}</td>
                            <td>${escapeHtml(translateResponseStatus(contact.responseStatus))}</td>
                          </tr>
                        `,
                      )
                      .join('')
                  : `<tr><td colspan="6"><div class="empty">${t('noContactHistory')}</div></td></tr>`
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
          <div class="panel-head"><h2>${t('quotationManagementTitle')}</h2><span>${t('quotationManagementDesc')}</span></div>
          <form id="quotation-form" class="form-grid">
            <label><span>${t('lead')}</span><select name="leadId" required><option value="">${t('selectCompany')}</option>${state.data.leads
              .map((lead) => `<option value="${lead.id}" ${activeLead?.id === lead.id ? 'selected' : ''}>${escapeHtml(lead.companyName)} • ${escapeHtml(lead.country)}</option>`)
              .join('')}</select></label>
            <label><span>${t('quotationNumber')}</span><input name="quotationNumber" required /></label>
            <label><span>${t('dateSent')}</span><input name="dateSent" type="date" value="${new Date().toISOString().slice(0, 10)}" required /></label>
            <label><span>${t('productType')}</span><input name="productType" required /></label>
            <label><span>${t('quotationValueUsd')}</span><input name="quotationValue" type="number" min="0" step="0.01" required /></label>
            <label><span>${t('validityDate')}</span><input name="validityDate" type="date" required /></label>
            <label class="wide"><span>${t('customerFeedback')}</span><textarea name="customerFeedback" rows="3" placeholder="${t('customerFeedbackPlaceholder')}"></textarea></label>
            <div class="form-actions wide"><button class="btn btn-primary" type="submit">${t('saveQuotation')}</button></div>
          </form>
        </article>

        <article class="panel">
          <div class="panel-head"><h2>${t('quotationSnapshot')}</h2><span>${t('quotationSnapshotDesc')}</span></div>
          <div class="stack-list">
            <div class="summary-card"><strong>${t('totalQuotations')}</strong><p>${allQuotations.length}</p><small>${t('allQuotationRecords')}</small></div>
            <div class="summary-card"><strong>${t('totalQuotedValue')}</strong><p>${formatCurrency(allQuotations.reduce((sum, item) => sum + item.quotationValue, 0))}</p><small>${t('forQuickControl')}</small></div>
            <div class="summary-card"><strong>${t('quotedCustomers')}</strong><p>${state.data.leads.filter((lead) => lead.quotations.length > 0).length}</p><small>${t('alreadyReceivedQuotation')}</small></div>
          </div>
        </article>
      </div>

      <article class="panel">
        <div class="panel-head"><h2>${t('quotationHistory')}</h2><span>${t('quotationHistoryDesc')}</span></div>
        <div class="table-shell">
          <table>
            <thead>
              <tr>
                <th>${t('quotation')}</th>
                <th>${t('company')}</th>
                <th>${t('date')}</th>
                <th>${t('product')}</th>
                <th>${t('value')}</th>
                <th>${t('feedback')}</th>
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
                            <td>${escapeHtml(quotation.customerFeedback || t('noFeedback'))}</td>
                          </tr>
                        `,
                      )
                      .join('')
                  : `<tr><td colspan="6"><div class="empty">${t('noQuotationsStored')}</div></td></tr>`
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
        <div class="panel-head"><h2>${t('leadDiscoveryTitle')}</h2><span>${t('leadDiscoveryDesc')}</span></div>
        <div class="filter-grid">
          <label><span>${t('search')}</span><input id="discovery-search" type="search" value="${escapeHtml(state.discoverySearch)}" placeholder="${t('discoverySearchPlaceholder')}" /></label>
          <label><span>${t('country')}</span><select id="discovery-country">${countries
            .map((country) => `<option value="${escapeHtml(country)}" ${state.discoveryCountry === country ? 'selected' : ''}>${escapeHtml(country === 'All' ? t('all') : country)}</option>`)
            .join('')}</select></label>
          <label><span>${t('industry')}</span><select id="discovery-category">${categories
            .map((category) => `<option value="${escapeHtml(category)}" ${state.discoveryCategory === category ? 'selected' : ''}>${escapeHtml(category)}</option>`)
            .join('')}</select></label>
          <label><span>${t('customerType')}</span><select id="discovery-type">${customerTypes
            .map((type) => `<option value="${escapeHtml(type)}" ${state.discoveryType === type ? 'selected' : ''}>${escapeHtml(type === 'All' ? t('all') : translateCustomerType(type))}</option>`)
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
                            <p>${escapeHtml(candidate.country)} • ${escapeHtml(translateCustomerType(candidate.customerType))}</p>
                          </div>
                          <span class="status-pill ${candidate.duplicate ? 'warn' : ''}">${candidate.duplicate ? t('alreadyExists') : t('newProspect')}</span>
                        </div>
                        <p>${escapeHtml(candidate.industryCategory)}</p>
                        <p>${escapeHtml(candidate.email || t('noEmail'))} • ${escapeHtml(candidate.contactPerson || t('noPic'))}</p>
                        <div class="tag-row">${candidate.keywords.map((keyword) => `<span class="tag">${escapeHtml(keyword)}</span>`).join('')}</div>
                        ${
                          candidate.duplicate
                            ? `<div class="warning-box">${escapeHtml(candidate.duplicate.companyName)} ${state.uiLanguage === 'id' ? 'sudah ada' : 'already exists'} • ${t('lastContact').toLowerCase()} ${formatDate(getLastActivityDate(candidate.duplicate))} • ${t('status').toLowerCase()} ${escapeHtml(translateLeadStatus(candidate.duplicate.status))}</div>`
                            : ''
                        }
                        <div class="form-actions">
                          <button class="btn btn-primary" data-action="import-candidate" data-id="${candidate.id}" ${candidate.duplicate ? 'disabled' : ''}>${t('addToCrm')}</button>
                        </div>
                      </article>
                    `,
                  )
                  .join('')
              : `<p class="empty">${t('noDiscoveryCandidates')}</p>`
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
          <div class="panel-head"><h2>${t('aiAssistantTitle')}</h2><span>${t('aiAssistantDesc')}</span></div>
          <form id="assistant-form" class="form-grid">
            <label><span>${t('lead')}</span><select name="leadId" required>${state.data.leads
              .map((lead) => `<option value="${lead.id}" ${state.assistantLeadId === lead.id ? 'selected' : ''}>${escapeHtml(lead.companyName)} • ${escapeHtml(lead.country)}</option>`)
              .join('')}</select></label>
            <label><span>${t('contentType')}</span><select name="assistantType">${assistantTypes
              .map((type) => `<option value="${type}" ${state.assistantType === type ? 'selected' : ''}>${translateAssistantType(type)}</option>`)
              .join('')}</select></label>
            <label class="wide"><span>${t('additionalContext')}</span><textarea name="context" rows="5" placeholder="${t('additionalContextPlaceholder')}"></textarea></label>
            <div class="form-actions wide">
              <button class="btn btn-primary" type="submit">${t('generateDraft')}</button>
              <button class="btn" type="button" data-action="copy-output">${t('copyOutput')}</button>
            </div>
          </form>
        </article>

        <article class="panel">
          <div class="panel-head"><h2>${t('draftOutput')}</h2><span>${escapeHtml(currentLead?.companyName || t('selectLead'))}</span></div>
          <div class="draft-box">
            <label><span>${t('subject')}</span><input readonly value="${escapeHtml(state.generatedSubject)}" /></label>
            <label><span>${t('content')}</span><textarea readonly rows="16">${escapeHtml(state.generatedOutput)}</textarea></label>
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
      setNotice(state.uiLanguage === 'id' ? 'Data sample berhasil dimuat.' : 'Sample data loaded.', 'success')
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
      setNotice(state.uiLanguage === 'id' ? 'Data CRM berhasil direset.' : 'CRM data reset.', 'info')
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
    if (target.id === 'language-switch') {
      state.uiLanguage = target.value === 'en' ? 'en' : 'id'
      persistUiLanguage()
      renderApp()
      return
    }
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
      setNotice(state.uiLanguage === 'id' ? `Status untuk ${lead.companyName} berhasil diperbarui.` : `Status updated for ${lead.companyName}.`, 'success')
      renderApp()
    }
  })

  app.addEventListener('input', (event) => {
    const target = event.target as HTMLInputElement
    if (target.id === 'lead-search' || target.id === 'global-search') {
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
    setNotice(
      state.uiLanguage === 'id'
        ? `Stop double sent: ${duplicate.companyName} sudah ada • kontak terakhir ${formatDate(getLastActivityDate(duplicate))} • status ${translateLeadStatus(duplicate.status)}`
        : `Stop double sent: ${duplicate.companyName} already exists • last contact ${formatDate(getLastActivityDate(duplicate))} • status ${duplicate.status}`,
      'error',
    )
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
  setNotice(state.uiLanguage === 'id' ? `Customer ${lead.companyName} berhasil disimpan.` : `Customer ${lead.companyName} saved.`, 'success')
  renderApp()
}

function handleContactSubmit(form: HTMLFormElement): void {
  const formData = new FormData(form)
  const lead = getLeadById(String(formData.get('leadId') || ''))
  if (!lead) {
    setNotice(state.uiLanguage === 'id' ? 'Silakan pilih company.' : 'Please select a company.', 'error')
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
  setNotice(state.uiLanguage === 'id' ? `Follow up untuk ${lead.companyName} berhasil diperbarui.` : `Follow up updated for ${lead.companyName}.`, 'success')
  renderApp()
}

function handleQuotationSubmit(form: HTMLFormElement): void {
  const formData = new FormData(form)
  const lead = getLeadById(String(formData.get('leadId') || ''))
  if (!lead) {
    setNotice(state.uiLanguage === 'id' ? 'Silakan pilih company.' : 'Please select a company.', 'error')
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
  setNotice(state.uiLanguage === 'id' ? `Quotation untuk ${lead.companyName} berhasil disimpan.` : `Quotation saved for ${lead.companyName}.`, 'success')
  renderApp()
}

function handleAssistantSubmit(form: HTMLFormElement): void {
  const formData = new FormData(form)
  const lead = getLeadById(String(formData.get('leadId') || ''))
  if (!lead) {
    setNotice(state.uiLanguage === 'id' ? 'Silakan pilih lead.' : 'Please choose a lead.', 'error')
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
  setNotice(
    state.uiLanguage === 'id'
      ? `${translateAssistantType(type)} untuk ${lead.companyName} berhasil dibuat.`
      : `${type} generated for ${lead.companyName}.`,
    'success',
  )
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
    setNotice(
      state.uiLanguage === 'id'
        ? `Stop double sent: ${duplicate.companyName} sudah ada di database.`
        : `Stop double sent: ${duplicate.companyName} already exists in database.`,
      'error',
    )
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
  setNotice(state.uiLanguage === 'id' ? `${candidate.companyName} berhasil ditambahkan ke CRM.` : `${candidate.companyName} added to CRM.`, 'success')
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
  setNotice(state.uiLanguage === 'id' ? 'Excel berhasil diexport.' : 'Excel exported successfully.', 'success')
  renderApp()
}

async function copyGeneratedOutput(): Promise<void> {
  const output = `${state.generatedSubject}\n\n${state.generatedOutput}`.trim()
  if (!output) {
    setNotice(state.uiLanguage === 'id' ? 'Buat draft terlebih dulu.' : 'Generate a draft first.', 'error')
    renderApp()
    return
  }
  await navigator.clipboard.writeText(output)
  setNotice(state.uiLanguage === 'id' ? 'Draft berhasil disalin ke clipboard.' : 'Draft copied to clipboard.', 'success')
  renderApp()
}
