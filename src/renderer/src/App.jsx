import { useState, useEffect, useCallback, useRef } from 'react'
import {
  Image,
  FileText,
  LayoutGrid,
  Folder,
  File,
  Settings,
  Upload,
  Search,
  Trash2,
  Edit2,
  X
} from 'lucide-react'
import * as XLSX from 'xlsx'
import './assets/main.css'
import logo from '../../Logo/RRM EMAIL APP.png?asset'

const formatExcelDate = (serial) => {
  if (!serial || isNaN(serial)) return serial
  try {
    const date = new Date(Math.round((serial - 25569) * 86400 * 1000))
    const dd = String(date.getDate()).padStart(2, '0')
    const mm = String(date.getMonth() + 1).padStart(2, '0')
    const yyyy = date.getFullYear()
    return `${dd}/${mm}/${yyyy}`
  } catch {
    return serial
  }
}

const formatExcelTime = (serial) => {
  if (serial === undefined || serial === null || isNaN(serial)) return serial
  try {
    const totalSeconds = Math.round(serial * 86400)
    const hours = Math.floor(totalSeconds / 3600)
    const minutes = Math.floor((totalSeconds % 3600) / 60)
    const ampm = hours >= 12 ? 'PM' : 'AM'
    const h12 = hours % 12 || 12
    const mm = String(minutes).padStart(2, '0')
    return `${String(h12).padStart(2, '0')}:${mm} ${ampm}`
  } catch {
    return serial
  }
}

const formatCurrency = (value) => {
  if (value === undefined || value === null || isNaN(value)) return value
  try {
    return new Intl.NumberFormat('en-IN', {
      style: 'currency',
      currency: 'INR',
      maximumFractionDigits: 0
    }).format(value)
  } catch {
    return value
  }
}

const DEFAULT_SIGNATURES = {
  tagging: [
    {
      regards: 'With Regards',
      name: 'Rasila Pansuriya',
      designation: 'Sr. Manager - Operations',
      company: 'Rural Risk Management Pvt. Ltd.',
      address:
        'B-TF-7, Signet Plaza, Near Krunal Char Rasta,\nSamta Gotri Road, Vadodara-390023, Gujarat.',
      mobile: '8160441446',
      email: 'Cattle.tagging@gmail.com',
      regardsSize: '12px',
      regardsColor: '#333333',
      nameSize: '16px',
      nameColor: '#4f46e5',
      designationSize: '12px',
      designationColor: '#3b82f6',
      companySize: '16px',
      companyColor: '#ef4444',
      addressSize: '12px',
      addressColor: '#333333',
      mobileSize: '12px',
      mobileColor: '#333333',
      emailSize: '12px',
      emailColor: '#3b82f6',
      isDefault: true
    },
    {
      regards: 'With Regards',
      name: 'Rasila Pansuriya',
      designation: 'Sr. Manager - Operations',
      company: 'Progressive Risk Management Pvt. Ltd.',
      address:
        'B-TF-7, Signet Plaza, Near Krunal Char Rasta,\nSamta Gotri Road, Vadodara-390023, Gujarat.',
      mobile: '8160441446',
      email: 'Cattle.tagging@gmail.com',
      regardsSize: '12px',
      regardsColor: '#333333',
      nameSize: '16px',
      nameColor: '#4f46e5',
      designationSize: '12px',
      designationColor: '#3b82f6',
      companySize: '16px',
      companyColor: '#ef4444',
      addressSize: '12px',
      addressColor: '#333333',
      mobileSize: '12px',
      mobileColor: '#333333',
      emailSize: '12px',
      emailColor: '#3b82f6',
      isDefault: false
    },
    {
      regards: 'With Regards',
      name: '',
      designation: '',
      company: '',
      address: '',
      mobile: '',
      email: '',
      regardsSize: '12px',
      regardsColor: '#333333',
      nameSize: '12px',
      nameColor: '#333333',
      designationSize: '12px',
      designationColor: '#333333',
      companySize: '12px',
      companyColor: '#333333',
      addressSize: '12px',
      addressColor: '#333333',
      mobileSize: '12px',
      mobileColor: '#333333',
      emailSize: '12px',
      emailColor: '#333333',
      isDefault: false
    }
  ],
  retagging: [
    {
      regards: 'With Regards',
      name: 'Rasila Pansuriya',
      designation: 'Sr. Manager - Operations',
      company: 'Rural Risk Management Pvt. Ltd.',
      address:
        'B-TF-7, Signet Plaza, Near Krunal Char Rasta,\nSamta Gotri Road, Vadodara-390023, Gujarat.',
      mobile: '8160441446',
      email: 'Cattle.tagging@gmail.com',
      regardsSize: '12px',
      regardsColor: '#333333',
      nameSize: '16px',
      nameColor: '#4f46e5',
      designationSize: '12px',
      designationColor: '#3b82f6',
      companySize: '16px',
      companyColor: '#ef4444',
      addressSize: '12px',
      addressColor: '#333333',
      mobileSize: '12px',
      mobileColor: '#333333',
      emailSize: '12px',
      emailColor: '#3b82f6',
      isDefault: true
    },
    {
      regards: 'With Regards',
      name: 'Rasila Pansuriya',
      designation: 'Sr. Manager - Operations',
      company: 'Progressive Risk Management Pvt. Ltd.',
      address:
        'B-TF-7, Signet Plaza, Near Krunal Char Rasta,\nSamta Gotri Road, Vadodara-390023, Gujarat.',
      mobile: '8160441446',
      email: 'Cattle.tagging@gmail.com',
      regardsSize: '12px',
      regardsColor: '#333333',
      nameSize: '16px',
      nameColor: '#4f46e5',
      designationSize: '12px',
      designationColor: '#3b82f6',
      companySize: '16px',
      companyColor: '#ef4444',
      addressSize: '12px',
      addressColor: '#333333',
      mobileSize: '12px',
      mobileColor: '#333333',
      emailSize: '12px',
      emailColor: '#3b82f6',
      isDefault: false
    },
    {
      regards: 'With Regards',
      name: '',
      designation: '',
      company: '',
      address: '',
      mobile: '',
      email: '',
      regardsSize: '12px',
      regardsColor: '#333333',
      nameSize: '12px',
      nameColor: '#333333',
      designationSize: '12px',
      designationColor: '#333333',
      companySize: '12px',
      companyColor: '#333333',
      addressSize: '12px',
      addressColor: '#333333',
      mobileSize: '12px',
      mobileColor: '#333333',
      emailSize: '12px',
      emailColor: '#333333',
      isDefault: false
    }
  ],
  claim: [
    {
      regards: 'With Regards',
      name: 'Dr. Ekta Vaghela',
      designation: 'Sr. Manager - Claims',
      company: 'Progressive Risk Management Pvt. Ltd.',
      address:
        'B-TF-7, Signet Plaza, Near Krunal Char Rasta,\nSamta Gotri Road, Vadodara-390023, Gujarat.',
      mobile: '8200194399',
      email: 'claim.cattle@gmail.com',
      regardsSize: '12px',
      regardsColor: '#333333',
      nameSize: '16px',
      nameColor: '#4f46e5',
      designationSize: '12px',
      designationColor: '#3b82f6',
      companySize: '16px',
      companyColor: '#ef4444',
      addressSize: '12px',
      addressColor: '#333333',
      mobileSize: '12px',
      mobileColor: '#333333',
      emailSize: '12px',
      emailColor: '#3b82f6',
      isDefault: true
    },
    {
      regards: 'With Regards',
      name: '',
      designation: '',
      company: '',
      address: '',
      mobile: '',
      email: '',
      regardsSize: '12px',
      regardsColor: '#333333',
      nameSize: '12px',
      nameColor: '#333333',
      designationSize: '12px',
      designationColor: '#333333',
      companySize: '12px',
      companyColor: '#333333',
      addressSize: '12px',
      addressColor: '#333333',
      mobileSize: '12px',
      mobileColor: '#333333',
      emailSize: '12px',
      emailColor: '#333333',
      isDefault: false
    },
    {
      regards: 'With Regards',
      name: '',
      designation: '',
      company: '',
      address: '',
      mobile: '',
      email: '',
      regardsSize: '12px',
      regardsColor: '#333333',
      nameSize: '12px',
      nameColor: '#333333',
      designationSize: '12px',
      designationColor: '#333333',
      companySize: '12px',
      companyColor: '#333333',
      addressSize: '12px',
      addressColor: '#333333',
      mobileSize: '12px',
      mobileColor: '#333333',
      emailSize: '12px',
      emailColor: '#333333',
      isDefault: false
    }
  ]
}

const DEFAULT_TEMPLATES = {
  tagging: [
    {
      body: 'Please Find Attached Health Certificates, Data Excel & Tagging Photos For Quick Policy Process.',
      isDefault: true
    },
    { body: '', isDefault: false },
    { body: '', isDefault: false }
  ],
  retagging: [
    {
      body: 'Please Find Attached Retagging Certificate, Tagging & Retagging Photos For Quick Endorsement Process.',
      isDefault: true
    },
    { body: '', isDefault: false },
    { body: '', isDefault: false }
  ],
  claim: [
    {
      body: 'Please Find attached all documents for quick claim process.',
      isDefault: true
    },
    {
      body: 'Please Find attached all documents and courier details for quick claim process.',
      isDefault: false
    },
    { body: '', isDefault: false }
  ]
}

const VALID_KEYS = [
  'KPYM2-XDY4Q-27BG6-C8J9Q-B8HBD',
  'NRV92-B7X6W-PQ98K-GRL2M-FJ54S',
  'XW87Z-LQ53M-RT92V-KJ41P-SB06N',
  'TL92P-MK31S-QV74N-RZ06D-FW85X'
]

function App() {
  // Application Main Component
  const [isActivated, setIsActivated] = useState(false)
  const [inputKey, setInputKey] = useState('')
  const [activationError, setActivationError] = useState('')
  const [checkingActivation, setCheckingActivation] = useState(true)

  const [folderName, setFolderName] = useState('')
  const [parsedData, setParsedData] = useState([])
  const [currentIndex, setCurrentIndex] = useState(0)
  const [attachments, setAttachments] = useState([])

  // Manage Settings State
  const [showManage, setShowManage] = useState(false)
  const [appVersion, setAppVersion] = useState('1.0.0')
  const [activeManageSection, setActiveManageSection] = useState(null) // 'smtp', 'settings', 'template'
  const [activeTemplateTab, setActiveTemplateTab] = useState('tagging') // 'tagging', 'retagging', 'claim'
  const [activeSignatureTab, setActiveSignatureTab] = useState('tagging') // 'tagging', 'retagging', 'claim'
  const [templates, setTemplates] = useState(DEFAULT_TEMPLATES)
  const [signatures, setSignatures] = useState(DEFAULT_SIGNATURES)
  const [smtpAccounts, setSmtpAccounts] = useState([
    {
      senderName: '',
      host: '',
      port: '465',
      user: '',
      pass: '',
      enableImapPop: false,
      imapHost: '',
      imapPort: '993',
      popHost: '',
      popPort: '995',
      isTagging: true,
      isRetagging: true,
      isClaim: true
    },
    {
      senderName: '',
      host: '',
      port: '465',
      user: '',
      pass: '',
      enableImapPop: false,
      imapHost: '',
      imapPort: '993',
      popHost: '',
      popPort: '995',
      isTagging: false,
      isRetagging: false,
      isClaim: false
    },
    {
      senderName: '',
      host: '',
      port: '465',
      user: '',
      pass: '',
      enableImapPop: false,
      imapHost: '',
      imapPort: '993',
      popHost: '',
      popPort: '995',
      isTagging: false,
      isRetagging: false,
      isClaim: false
    }
  ])
  const [activeSmtpTab, setActiveSmtpTab] = useState(0)
  const [smtpStatus, setSmtpStatus] = useState('idle') // idle, loading, connected, disconnected
  const [alwaysCc, setAlwaysCc] = useState('')
  const [manageData, setManageData] = useState([])
  const [newEntry, setNewEntry] = useState({
    greetingName: '',
    type: '',
    ccEmail: ''
  })
  const [saveStatus, setSaveStatus] = useState('idle') // idle, saving, saved
  const [editingIndex, setEditingIndex] = useState(null)
  const [searchTerm, setSearchTerm] = useState('')
  const [selectedIndices, setSelectedIndices] = useState([])

  // Batch Sending State
  const [isBatchSending, setIsBatchSending] = useState(false)
  const [batchProgress, setBatchProgress] = useState({
    current: 0,
    total: 0,
    success: 0,
    failed: 0,
    currentItemName: ''
  })
  const [showBatchSummary, setShowBatchSummary] = useState(false)
  const cancelBatchRef = useRef(false)

  const updateCurrentItem = useCallback(
    (field, value) => {
      setParsedData((prev) => {
        const updated = [...prev]
        updated[currentIndex] = { ...updated[currentIndex], [field]: value }
        return updated
      })
    },
    [currentIndex]
  )

  const currentItem = parsedData[currentIndex] || null
  const totalAttachmentsSizeMB = (
    attachments.reduce((acc, att) => acc + (att.size || 0), 0) /
    (1024 * 1024)
  ).toFixed(2)

  const generateSignatureHtml = useCallback((sig) => {
    if (!sig) return '<i style="color: #94a3b8;">No default signature found for this category.</i>'
    const {
      regards,
      name,
      designation,
      company,
      address,
      mobile,
      email,
      regardsSize,
      regardsColor,
      nameSize,
      nameColor,
      designationSize,
      designationColor,
      companySize,
      companyColor,
      addressSize,
      addressColor,
      mobileSize,
      mobileColor,
      emailSize,
      emailColor
    } = sig

    const rSize = regardsSize || '14px'
    const rColor = regardsColor || '#333'
    const nSize = nameSize || '16px'
    const nColor = nameColor || '#3b82f6'
    const dSize = designationSize || '14px'
    const dColor = designationColor || '#1e3a8a'
    const cSize = companySize || '14px'
    const cColor = companyColor || '#ef4444'
    const aSize = addressSize || '13px'
    const aColor = addressColor || '#475569'
    const mSize = mobileSize || '13px'
    const mColor = mobileColor || '#475569'
    const eSize = emailSize || '13px'
    const eColor = emailColor || '#3b82f6'

    // Create Mobile/Email combo line
    const mPart = mobile ? `M: ${mobile}` : ''
    const ePart = email
      ? `<a href="mailto:${email}" style="color: ${eColor}; text-decoration: none;">${email}</a>`
      : ''

    const contactLine = [
      mPart ? `<span style="font-size: ${mSize}; color: ${mColor};">${mPart}</span>` : '',
      ePart ? `<span style="font-size: ${eSize}; color: ${eColor};">${ePart}</span>` : ''
    ]
      .filter(Boolean)
      .join(', ')

    const rows = []
    if (regards) {
      rows.push(
        `<tr><td style="padding: 0; font-size: ${rSize}; color: ${rColor};">${regards},</td></tr>`
      )
    }
    if (name) {
      rows.push(
        `<tr><td style="padding: 2px 0; font-size: ${nSize}; font-weight: bold; color: ${nColor};">${name}</td></tr>`
      )
    }
    if (designation) {
      rows.push(
        `<tr><td style="padding: 0; font-size: ${dSize}; font-weight: 600; color: ${dColor};">${designation}</td></tr>`
      )
    }
    if (company) {
      rows.push(
        `<tr><td style="padding: 2px 0; font-size: ${cSize}; font-weight: bold; color: ${cColor};">${company}</td></tr>`
      )
    }
    if (address) {
      rows.push(
        `<tr><td style="padding: 2px 0; font-size: ${aSize}; color: ${aColor}; white-space: pre-wrap;">${address.replace(
          /\n/g,
          '<br>'
        )}</td></tr>`
      )
    }
    if (contactLine) {
      rows.push(`<tr><td style="padding: 0;">${contactLine}</td></tr>`)
    }

    const tableStyle =
      'font-family: Arial, Helvetica, sans-serif; margin-top: 25px; border-collapse: collapse;'
    return `
      <table border="0" cellpadding="0" cellspacing="0" style="${tableStyle}">
        ${rows.join('')}
      </table>
    `
  }, [])

  const generateClaimTableHtml = useCallback((item) => {
    if (!item || item.emailType !== 'Claim' || !item.fetchedEmails?.claimDetails) return ''
    const d = item.fetchedEmails.claimDetails
    const rows = [
      ['Claim Reg. No.', d.claimRegNo],
      ['Name of Beneficiary', d.nameOfBeneficiary],
      ['Phone No', d.phoneNo],
      ['Address', d.address],
      ['Village', d.village],
      ['Taluka', d.taluka],
      ['District', d.district],
      ['State', d.state],
      ['Pin Code', d.pinCode],
      ['Date of Death', formatExcelDate(d.dateOfDeath)],
      ['Time of Death', formatExcelTime(d.timeOfDeath)],
      ['Tag No.', d.tagNo],
      ['Animal Species', d.animalSpecies],
      ['Sum Insured', formatCurrency(d.sumInsured)],
      ['Insured Name Financing Bank', d.insuredNameFinancingBank],
      ['Insured Address /Bank Branch', d.insuredAddressBankBranch],
      ['Proposal / Loan Account No.', d.proposalLoanAccountNo],
      ['Insurance Company', d.insuranceCompany],
      ['Policy Number', d.policyNumber],
      ['Policy Date', formatExcelDate(d.policyDate)],
      ['Surveyer Name', d.surveyerName]
    ]

    return `
      <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; max-width: 600px; margin: 25px 0; font-family: sans-serif; font-size: 14px; border: 1px solid #e2e8f0; color: #1e293b;">
        <tbody>
          ${rows
            .map(
              ([label, value]) => `
            <tr>
              <th align="left" style="background-color: #f8fafc; width: 45%; font-weight: 600; border: 1px solid #e2e8f0; color: #475569;">${label}</th>
              <td style="border: 1px solid #e2e8f0; color: #0f172a;">${value || ''}</td>
            </tr>
          `
            )
            .join('')}
        </tbody>
      </table>
    `
  }, [])

  const generateClaimTableText = useCallback((item) => {
    if (!item || item.emailType !== 'Claim' || !item.fetchedEmails?.claimDetails) return ''
    const d = item.fetchedEmails.claimDetails
    const rows = [
      ['Claim Reg. No.', d.claimRegNo],
      ['Name of Beneficiary', d.nameOfBeneficiary],
      ['Phone No', d.phoneNo],
      ['Address', d.address],
      ['Village', d.village],
      ['Taluka', d.taluka],
      ['District', d.district],
      ['State', d.state],
      ['Pin Code', d.pinCode],
      ['Date of Death', formatExcelDate(d.dateOfDeath)],
      ['Time of Death', formatExcelTime(d.timeOfDeath)],
      ['Tag No.', d.tagNo],
      ['Animal Species', d.animalSpecies],
      ['Sum Insured', formatCurrency(d.sumInsured)],
      ['Insured Name Financing Bank', d.insuredNameFinancingBank],
      ['Insured Address /Bank Branch', d.insuredAddressBankBranch],
      ['Proposal / Loan Account No.', d.proposalLoanAccountNo],
      ['Insurance Company', d.insuranceCompany],
      ['Policy Number', d.policyNumber],
      ['Policy Date', formatExcelDate(d.policyDate)],
      ['Surveyer Name', d.surveyerName]
    ]
    return '\n\n' + rows.map(([label, value]) => `${label}: ${value || ''}`).join('\n')
  }, [])

  const generateSignatureText = useCallback((sig) => {
    if (!sig) return ''
    const { regards, name, designation, company, address, mobile, email } = sig

    // Create Mobile/Email combo line
    const contactLine = [mobile ? `M: ${mobile}` : '', email ? email : '']
      .filter(Boolean)
      .join(', ')

    const parts = []
    if (regards) parts.push(`${regards},`)
    if (name) parts.push(name)
    if (designation) parts.push(designation)
    if (company) parts.push(company)
    if (address) parts.push(address)
    if (contactLine) parts.push(contactLine)
    return parts.length > 0 ? '\n\n' + parts.join('\n') : ''
  }, [])

  const generateEmailBodyContent = useCallback(
    (item) => {
      if (!item) return ''
      const catKey = item.emailType.toLowerCase()
      const catTemplates = templates[catKey]
      if (!catTemplates) return ''

      const defaultTpl = catTemplates.find((t) => t.isDefault)
      if (!defaultTpl || !defaultTpl.body) return ''

      // Find custom greeting name from manage data
      const matchedRecipient = manageData.find(
        (m) => m.email?.trim().toLowerCase() === item.emailTo?.trim().toLowerCase()
      )
      const name = matchedRecipient ? matchedRecipient.greetingName : ''

      // Substitue {{name}} in the template body itself
      const processedTemplate = defaultTpl.body.replace(/{{name}}/g, name)

      // Construct final body
      return `${name ? name + '\n\n' : ''}${processedTemplate}`
    },
    [templates, manageData]
  )

  // Load settings on mount
  useEffect(() => {
    const getVersion = async () => {
      const res = await window.api.getAppVersion()
      if (res && res.version) setAppVersion(res.version)
    }
    getVersion()

    const load = async () => {
      try {
        const saved = await window.api.loadSettings()
        if (saved) {
          if (Array.isArray(saved.smtpAccounts)) {
            const accountsWithSenderName = saved.smtpAccounts.map((acc) => ({
              ...acc,
              senderName: acc.senderName || ''
            }))
            setSmtpAccounts(accountsWithSenderName)
          } else if (saved.host) {
            // Migrate from old single-account format
            setSmtpAccounts((prev) => {
              const updated = [...prev]
              updated[0] = {
                ...updated[0],
                senderName: saved.senderName || '',
                host: saved.host || '',
                port: saved.port || '465',
                user: saved.user || '',
                pass: saved.pass || '',
                enableImapPop: saved.enableImapPop || false,
                imapHost: saved.imapHost || '',
                imapPort: saved.imapPort || '993',
                popHost: saved.popHost || '',
                popPort: saved.popPort || '995'
              }
              return updated
            })
          }
          setAlwaysCc(saved.alwaysCc || '')
        }

        const savedManage = await window.api.loadManageData()
        if (savedManage) setManageData(savedManage)

        const savedTemplates = await window.api.loadTemplates()
        if (savedTemplates) {
          const normalized = {}
          ;['tagging', 'retagging', 'claim'].forEach((cat) => {
            if (typeof savedTemplates[cat] === 'string') {
              normalized[cat] = [
                {
                  body: savedTemplates[cat] || DEFAULT_TEMPLATES[cat][0].body,
                  isDefault: true
                },
                {
                  body: cat === 'claim' ? DEFAULT_TEMPLATES[cat][1].body : '',
                  isDefault: false
                },
                { body: '', isDefault: false }
              ]
            } else {
              normalized[cat] = savedTemplates[cat] || DEFAULT_TEMPLATES[cat]
            }
          })
          setTemplates(normalized)
        }

        const savedSignatures = await window.api.loadSignatures()
        if (savedSignatures) {
          const merged = {}
          ;['tagging', 'retagging', 'claim'].forEach((cat) => {
            if (Array.isArray(savedSignatures[cat])) {
              merged[cat] = savedSignatures[cat].map((sig, idx) => {
                const def = DEFAULT_SIGNATURES[cat][idx] || DEFAULT_SIGNATURES[cat][0]
                // If signature is empty (missing key fields), use default
                if (!sig.name && !sig.company && !sig.designation) {
                  return { ...def, ...sig, ...def } // def wins for empty fields
                }
                // Otherwise merge, ensuring all required properties exist
                return { ...def, ...sig }
              })
            } else {
              merged[cat] = DEFAULT_SIGNATURES[cat]
            }
          })
          setSignatures(merged)
        } else {
          setSignatures(DEFAULT_SIGNATURES)
        }
      } catch (err) {
        console.error('Error loading data:', err)
      }
    }
    load()
  }, [])

  useEffect(() => {
    const checkLicense = async () => {
      try {
        const licenseData = await window.api.loadLicense()
        if (licenseData?.key && VALID_KEYS.includes(licenseData.key.toUpperCase())) {
          setIsActivated(true)
        }
      } catch (err) {
        console.error('Error loading license:', err)
      } finally {
        setCheckingActivation(false)
      }
    }
    checkLicense()
  }, [])

  const handleActivate = async (e) => {
    e.preventDefault()
    const cleanedKey = inputKey.trim().toUpperCase()
    if (VALID_KEYS.includes(cleanedKey)) {
      try {
        await window.api.saveLicense({ key: cleanedKey, activatedAt: new Date().toISOString() })
      } catch (err) {
        console.error('Error saving license:', err)
      }
      setIsActivated(true)
      setActivationError('')
    } else {
      setActivationError('Invalid Product Key')
    }
  }

  // Auto-save removed to prevent race conditions that could lead to data loss.
  // Manual save buttons are used to persist settings.

  useEffect(() => {
    if (!currentItem) return
    const fetchAttachments = async () => {
      const finalPath = currentItem.rootDir + '/' + currentItem.folderName
      const atts = await window.api.getAttachments(finalPath, currentItem.emailType)
      const filtered = atts.filter((a) => !currentItem.removedAttachments?.includes(a.name))
      setAttachments(filtered)
    }
    fetchAttachments()
  }, [currentIndex, currentItem, currentItem?.removedAttachments])

  useEffect(() => {
    if (!currentItem || currentItem.emailBody) return
    const finalBody = generateEmailBodyContent(currentItem)
    if (finalBody) {
      updateCurrentItem('emailBody', finalBody)
    }
  }, [currentIndex, currentItem, generateEmailBodyContent, updateCurrentItem])

  const handleSelectFolder = async () => {
    const result = await window.api.selectFolder()
    if (result) {
      setFolderName(result.rootDir)

      let type = 'Tagging' // default fallback
      const rootLower = result.rootDir.toLowerCase()
      if (rootLower.includes('claim')) {
        type = 'Claim'
      } else if (rootLower.includes('retagging')) {
        type = 'Retagging'
      }

      const parsed = await Promise.all(
        result.subfolders.map(async (sfObj) => {
          const sf = sfObj.name
          const inner = sfObj.innerFolder

          let genSubject = sf
          if (type === 'Claim' && inner) {
            genSubject = inner
          }

          if (type === 'Claim' && sf.toLowerCase().includes('fraud')) {
            if (!genSubject.toLowerCase().startsWith('fraud')) {
              genSubject = `Fraud ${genSubject}`
            }
          }

          const regex = /^(.*?)\s*\((.*?)\)\s*Dt\.\s*(.*?)\s*-\s*(.*)$/i
          const folderMatch = sf.match(regex)

          // Fetch Email IDs from Excel
          const subfolderPath = result.rootDir + '/' + sf
          const fetchedEmails = await window.api.getEmailIds(subfolderPath, type)

          let finalTo = ''
          let finalCc = []

          if (fetchedEmails) {
            const { insurance, bank, leader } = fetchedEmails

            if (insurance) {
              finalTo = insurance
              const recipientMatch = manageData.find(
                (m) => m.email?.trim().toLowerCase() === insurance.trim().toLowerCase()
              )
              if (recipientMatch?.ccEmail) finalCc.push(recipientMatch.ccEmail)
              if (bank) finalCc.push(bank)
              if (leader) finalCc.push(leader)
            } else if (bank) {
              finalTo = bank
              const recipientMatch = manageData.find(
                (m) => m.email?.trim().toLowerCase() === bank.trim().toLowerCase()
              )
              if (recipientMatch?.ccEmail) finalCc.push(recipientMatch.ccEmail)
              if (leader) finalCc.push(leader)
            } else if (leader) {
              finalTo = leader
              const recipientMatch = manageData.find(
                (m) => m.email?.trim().toLowerCase() === leader.trim().toLowerCase()
              )
              if (recipientMatch?.ccEmail) finalCc.push(recipientMatch.ccEmail)
            } else {
              finalTo = 'No Email Id Found'
            }
          } else {
            finalTo = 'No Email Id Found'
          }

          if (alwaysCc) finalCc.push(alwaysCc)

          const baseItem = {
            rootDir: result.rootDir,
            folderName: sf,
            innerFolder: inner || null,
            subject: genSubject,
            emailType: type,
            emailTo: finalTo,
            ccTo: finalCc.filter(Boolean).join(', '),
            emailBody: '',
            status: 'Pending',
            fetchedEmails: fetchedEmails,
            removedAttachments: []
          }

          if (folderMatch) {
            return {
              ...baseItem,
              location: folderMatch[1].trim(),
              cattleCount: folderMatch[2].trim(),
              date: folderMatch[3].trim(),
              customerName: folderMatch[4].trim()
            }
          } else {
            return {
              ...baseItem,
              location: '',
              cattleCount: '',
              date: '',
              customerName: ''
            }
          }
        })
      )

      setParsedData(parsed)
      setCurrentIndex(0)
    }
  }

  const handleExcelUpload = (e) => {
    const file = e.target.files[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = async (evt) => {
      const bstr = evt.target.result
      const wb = XLSX.read(bstr, { type: 'binary' })
      const wsname = wb.SheetNames[0]
      const ws = wb.Sheets[wsname]

      const rows = XLSX.utils.sheet_to_json(ws, { header: 1 })
      if (rows.length < 2) return

      const headerRow = rows[0] || []
      const dataRows = rows.slice(1)

      const getColIndex = (names, defaultIdx) => {
        const idx = headerRow.findIndex((h) =>
          names.some((n) => h?.toString().trim().toLowerCase() === n.toLowerCase())
        )
        return idx !== -1 ? idx : defaultIdx
      }

      const greetingIdx = getColIndex(['Professional Greetings with Name', 'Name', 'Greeting'], 0)
      const typeIdx = getColIndex(['Email ID Type', 'Type'], 1)
      const emailIdx = getColIndex(['Email', 'Email ID'], 2) // Default Column C
      const ccEmailIdx = getColIndex(['CC Email', 'CC'], 3)

      const excelData = dataRows
        .map((row) => ({
          greetingName: row[greetingIdx] || '',
          email: row[emailIdx] || '',
          type: row[typeIdx] || '',
          ccEmail: row[ccEmailIdx] || ''
        }))
        .filter((entry) => entry.email)

      setManageData((prevData) => {
        const updatedData = [...prevData]

        excelData.forEach((newEntry) => {
          const existingIndex = updatedData.findIndex(
            (item) => item.email?.trim().toLowerCase() === newEntry.email.trim().toLowerCase()
          )

          if (existingIndex !== -1) {
            // Update matching entry (merge to preserve any other existing fields)
            updatedData[existingIndex] = { ...updatedData[existingIndex], ...newEntry }
          } else {
            // Add new entry
            updatedData.push(newEntry)
          }
        })

        window.api.saveManageData(updatedData)
        return updatedData
      })
    }
    reader.readAsBinaryString(file)
  }

  const handleAddEntry = async (e) => {
    e.preventDefault()
    if (!newEntry.greetingName || !newEntry.email || !newEntry.type) {
      alert('Please fill in Name, Type and Email')
      return
    }

    let updatedData = [...manageData]
    if (editingIndex !== null) {
      updatedData[editingIndex] = { ...newEntry }
      setEditingIndex(null)
    } else {
      updatedData.push({ ...newEntry })
    }

    setManageData(updatedData)
    await window.api.saveManageData(updatedData)

    // Reset form
    setNewEntry({
      greetingName: '',
      type: '',
      email: '',
      ccEmail: ''
    })
  }

  const handleDeleteEntry = async (idx) => {
    if (window.confirm('Are you sure you want to delete this entry?')) {
      const updated = manageData.filter((_, i) => i !== idx)
      setManageData(updated)
      await window.api.saveManageData(updated)
      // Reset selection if deleted
      setSelectedIndices((prev) => prev.filter((i) => i !== idx).map((i) => (i > idx ? i - 1 : i)))
    }
  }

  const handleBulkDelete = async () => {
    if (
      window.confirm(`Are you sure you want to delete ${selectedIndices.length} selected entries?`)
    ) {
      const updated = manageData.filter((_, i) => !selectedIndices.includes(i))
      setManageData(updated)
      await window.api.saveManageData(updated)
      setSelectedIndices([])
    }
  }

  const handleEditEntry = (idx) => {
    setEditingIndex(idx)
    setNewEntry({ ...manageData[idx] })
  }

  const cancelEdit = () => {
    setEditingIndex(null)
    setNewEntry({
      greetingName: '',
      type: '',
      email: '',
      ccEmail: ''
    })
  }

  const toggleSelectAll = () => {
    if (selectedIndices.length === manageData.length) {
      setSelectedIndices([])
    } else {
      setSelectedIndices(manageData.map((_, i) => i))
    }
  }

  const toggleSelect = (idx) => {
    setSelectedIndices((prev) =>
      prev.includes(idx) ? prev.filter((i) => i !== idx) : [...prev, idx]
    )
  }

  const saveSmtpSettings = async () => {
    setSaveStatus('saving')
    try {
      const res = await window.api.saveSettings({
        smtpAccounts,
        alwaysCc
      })
      if (res.success) {
        setSaveStatus('saved')
        setTimeout(() => setSaveStatus('idle'), 2000)
        alert('Settings saved successfully!')
      } else {
        setSaveStatus('idle')
        alert('Failed to save settings: ' + res.error)
      }
    } catch (e) {
      setSaveStatus('idle')
      alert('Error during settings save: ' + e.message)
    }
  }

  const handleTestConnection = async () => {
    const currentAcc = smtpAccounts[activeSmtpTab]
    if (!currentAcc.host || !currentAcc.port || !currentAcc.user || !currentAcc.pass) {
      alert('Please fill in all SMTP fields for the current account first.')
      return
    }
    setSmtpStatus('loading')
    try {
      const res = await window.api.testSmtpConnection(currentAcc)
      if (res && res.success) {
        setSmtpStatus('connected')
      } else {
        setSmtpStatus('disconnected')
        alert('SMTP Connection failed: ' + (res?.error || 'Unknown error'))
      }
    } catch (error) {
      setSmtpStatus('disconnected')
      alert('SMTP Connection error: ' + error.message)
    }
  }

  const saveTemplate = async () => {
    setSaveStatus('saving')
    try {
      const res = await window.api.saveTemplates(templates)
      if (res.success) {
        setSaveStatus('saved')
        setTimeout(() => setSaveStatus('idle'), 2000)
      } else {
        setSaveStatus('idle')
        alert('Failed to save templates: ' + res.error)
      }
    } catch (e) {
      setSaveStatus('idle')
      alert('Error during template save: ' + e.message)
    }
  }

  const handleSaveSignatures = async () => {
    setSaveStatus('saving')
    try {
      const res = await window.api.saveSignatures(signatures)
      if (res.success) {
        setSaveStatus('saved')
        setTimeout(() => setSaveStatus('idle'), 2000)
      } else {
        setSaveStatus('idle')
        alert('Failed to save signatures: ' + res.error)
      }
    } catch (e) {
      setSaveStatus('idle')
      alert('Error during signature save: ' + e.message)
    }
  }

  const handleTemplateChange = (category, index, value) => {
    setTemplates((prev) => {
      const updatedCat = [...prev[category]]
      updatedCat[index] = { ...updatedCat[index], body: value }
      return { ...prev, [category]: updatedCat }
    })
  }

  const handleSetDefault = (category, index) => {
    setTemplates((prev) => {
      const updatedCat = prev[category].map((t, i) => ({
        ...t,
        isDefault: i === index
      }))
      return { ...prev, [category]: updatedCat }
    })
  }

  const handleSignatureChange = (category, index, field, value) => {
    setSignatures((prev) => {
      const updated = { ...prev }
      const newItems = [...updated[category]]
      newItems[index] = { ...newItems[index], [field]: value }
      updated[category] = newItems
      return updated
    })
  }

  const handleSetSignatureDefault = (category, index) => {
    setSignatures((prev) => {
      const updated = { ...prev }
      updated[category] = updated[category].map((s, i) => ({
        ...s,
        isDefault: i === index
      }))
      return updated
    })
  }

  const resetTemplate = () => {
    if (
      window.confirm(
        `Are you sure you want to reset all ${activeTemplateTab} templates to defaults?`
      )
    ) {
      setTemplates((prev) => ({
        ...prev,
        [activeTemplateTab]: DEFAULT_TEMPLATES[activeTemplateTab]
      }))
    }
  }

  const resetSignature = () => {
    if (
      window.confirm(
        `Are you sure you want to reset all ${activeSignatureTab} signatures to defaults?`
      )
    ) {
      setSignatures((prev) => ({
        ...prev,
        [activeSignatureTab]: DEFAULT_SIGNATURES[activeSignatureTab]
      }))
    }
  }

  const handleSendEmail = async (index) => {
    const item = parsedData[index]
    if (!item) return

    // Ensure we have a body (especially for unviewed items in a batch)
    let emailBody = item.emailBody
    if (!emailBody) {
      emailBody = generateEmailBodyContent(item)
    }

    // Update status to Sending
    setParsedData((prev) => {
      const updated = [...prev]
      updated[index] = { ...updated[index], status: 'Sending' }
      return updated
    })

    const finalPath = item.rootDir + '/' + item.folderName

    // Generate Signature with fallback
    const catKey = item.emailType.toLowerCase()
    const catSigs = signatures[catKey] || []
    const defaultSig = catSigs.find((s) => s.isDefault) || catSigs[0]

    const sigHtml = generateSignatureHtml(defaultSig)
    const sigText = generateSignatureText(defaultSig)

    const claimTableHtml = generateClaimTableHtml(item)
    const claimTableText = generateClaimTableText(item)

    const fullBodyText = (emailBody || '') + claimTableText + sigText
    const fullBodyHtml = `
      <div style="font-family: 'Outfit', sans-serif; color: #1f2937; font-size: 15px; line-height: 1.6;">
        ${(emailBody || '').replace(/\n/g, '<br>')}
        ${claimTableHtml}
        ${sigHtml}
      </div>
    `

    // Determine the SMTP account to use
    const emailType = item.emailType // Tagging, Retagging, or Claim
    const targetAccount =
      smtpAccounts.find((acc) => {
        if (emailType === 'Tagging') return acc.isTagging
        if (emailType === 'Retagging') return acc.isRetagging
        if (emailType === 'Claim') return acc.isClaim
        return false
      }) || smtpAccounts[0]

    const res = await window.api.sendEmail({
      smtpConfig: targetAccount,
      to: item.emailTo,
      cc: item.ccTo,
      subject: item.subject,
      body: fullBodyText,
      html: fullBodyHtml,
      folderPath: finalPath,
      emailType: item.emailType,
      excludedAttachments: item.removedAttachments || []
    })

    setParsedData((prev) => {
      const updated = [...prev]
      updated[index] = {
        ...updated[index],
        status: res.success ? 'Sent' : 'Error',
        emailBody: emailBody // Save computed body for future reference
      }
      return updated
    })

    return res.success
  }

  const handleSendAll = async () => {
    const itemsToSend = parsedData.filter(
      (item) => item.status === 'Pending' || item.status === 'Error'
    )
    if (itemsToSend.length === 0) {
      alert('No pending or error emails to send.')
      return
    }

    setIsBatchSending(true)
    setShowBatchSummary(false)
    cancelBatchRef.current = false
    setBatchProgress({
      current: 0,
      total: itemsToSend.length,
      success: 0,
      failed: 0,
      currentItemName: ''
    })

    let successCount = 0
    let failedCount = 0

    for (let i = 0; i < parsedData.length; i++) {
      if (cancelBatchRef.current) break

      if (parsedData[i].status === 'Pending' || parsedData[i].status === 'Error') {
        setBatchProgress((prev) => ({
          ...prev,
          currentItemName: parsedData[i].folderName,
          current: prev.current + 1
        }))

        const success = await handleSendEmail(i)
        if (success) successCount++
        else failedCount++

        setBatchProgress((prev) => ({
          ...prev,
          success: successCount,
          failed: failedCount
        }))

        // Mini delay between emails to avoid rate limits
        if (i < parsedData.length - 1 && !cancelBatchRef.current) {
          await new Promise((r) => setTimeout(r, 1000))
        }
      }
    }

    setIsBatchSending(false)
    setShowBatchSummary(true)
  }

  const handleRemoveAttachment = (attName) => {
    setParsedData((prev) => {
      const updated = [...prev]
      const item = { ...updated[currentIndex] }
      const removed = item.removedAttachments ? [...item.removedAttachments] : []
      if (!removed.includes(attName)) {
        removed.push(attName)
      }
      item.removedAttachments = removed
      updated[currentIndex] = item
      return updated
    })
  }

  const formatSize = (bytes) => {
    if (bytes < 1024) return bytes + ' B'
    else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB'
    else return (bytes / 1048576).toFixed(1) + ' MB'
  }

  const getIcon = (ext) => {
    if (ext === 'pdf') return { icon: <FileText size={16} />, className: 'pdf' }
    if (['jpg', 'png', 'jpeg'].includes(ext))
      return { icon: <Image size={16} />, className: 'image' }
    if (['xls', 'xlsx'].includes(ext)) return { icon: <LayoutGrid size={16} />, className: 'excel' }
    if (!ext) return { icon: <Folder size={16} />, className: 'folder' }
    return { icon: <File size={16} />, className: 'default' }
  }

  if (checkingActivation) {
    return <div className="layout-overlay"></div>
  }

  if (!isActivated) {
    return (
      <div className="layout-overlay animate-fade-in">
        <div className="activation-container">
          <div className="card glass-effect activation-card">
            <div className="activation-header">
              <h2>Product Activation</h2>
              <p>Please enter your 25-character product key to continue.</p>
            </div>
            <form onSubmit={handleActivate} className="activation-form">
              <div className="form-group input-animated">
                <input
                  type="text"
                  id="productKey"
                  placeholder=" "
                  value={inputKey}
                  onChange={(e) => {
                    const val = e.target.value.toUpperCase()
                    if (val.length <= 30) {
                      // allowance for hyphens
                      setInputKey(val)
                    }
                  }}
                  required
                />
                <label htmlFor="productKey">Enter Product Key</label>
              </div>
              {activationError && <div className="error-message">{activationError}</div>}
              <button type="submit" className="btn-primary activation-btn">
                Activate
              </button>
            </form>
          </div>
        </div>
      </div>
    )
  }

  return (
    <div className="layout-overlay">
      <div className="container">
        <div className="card glass-effect tall-card">
          <button
            className="manage-toggle"
            onClick={() => {
              setShowManage(true)
              setActiveManageSection(null)
            }}
          >
            <Settings size={18} />
            <span>Manage</span>
          </button>

          {showManage && (
            <div className="manage-overlay">
              <div className="settings-header">
                <h2 style={{ margin: 0 }}>Management Center</h2>
                <button
                  className="about-btn"
                  onClick={() => {
                    alert(
                      `RRM Email Automation\nVersion: ${appVersion}\n\nDeveloped by Epexio Techno solutions.\nFor any queries contact: [office@epexio.in]`
                    )
                  }}
                  title="About App"
                >
                  About
                </button>
                <button
                  className="close-btn"
                  onClick={() => {
                    setShowManage(false)
                    setActiveManageSection(null)
                  }}
                  title="Close Management Center"
                >
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    width="20"
                    height="20"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <line x1="18" y1="6" x2="6" y2="18"></line>
                    <line x1="6" y1="6" x2="18" y2="18"></line>
                  </svg>
                  <span>Close</span>
                </button>
              </div>

              <div className="manage-nav">
                <button
                  className={`nav-item ${activeManageSection === 'smtp' ? 'active' : ''}`}
                  onClick={() => setActiveManageSection('smtp')}
                >
                  Manage SMTP / IMAP
                </button>
                <button
                  className={`nav-item ${activeManageSection === 'settings' ? 'active' : ''}`}
                  onClick={() => setActiveManageSection('settings')}
                >
                  Manage Email Settings
                </button>
                <button
                  className={`nav-item ${activeManageSection === 'template' ? 'active' : ''}`}
                  onClick={() => setActiveManageSection('template')}
                >
                  Manage Email Body Template
                </button>
                <button
                  className={`nav-item ${activeManageSection === 'signatures' ? 'active' : ''}`}
                  onClick={() => setActiveManageSection('signatures')}
                >
                  Manage Email Signature
                </button>
              </div>

              <div className="manage-content">
                {activeManageSection === 'smtp' && (
                  <div className="section-container animate-fade-in">
                    <div className="template-nav" style={{ marginBottom: '20px' }}>
                      {[0, 1, 2].map((idx) => (
                        <button
                          key={idx}
                          className={`sub-nav-item ${activeSmtpTab === idx ? 'active' : ''}`}
                          onClick={() => {
                            setActiveSmtpTab(idx)
                            setSmtpStatus('idle')
                          }}
                        >
                          Account {idx + 1}
                        </button>
                      ))}
                    </div>

                    <div className="settings-grid-3">
                      <div className="form-group input-animated">
                        <input
                          type="text"
                          id="senderName"
                          placeholder=" "
                          value={smtpAccounts[activeSmtpTab].senderName || ''}
                          onChange={(e) => {
                            const updated = [...smtpAccounts]
                            updated[activeSmtpTab] = {
                              ...updated[activeSmtpTab],
                              senderName: e.target.value
                            }
                            setSmtpAccounts(updated)
                          }}
                        />
                        <label htmlFor="senderName">Sender Name (Display Name)</label>
                      </div>
                      <div className="form-group input-animated">
                        <input
                          type="text"
                          id="smtpHost"
                          placeholder=" "
                          value={smtpAccounts[activeSmtpTab].host}
                          onChange={(e) => {
                            const updated = [...smtpAccounts]
                            updated[activeSmtpTab] = {
                              ...updated[activeSmtpTab],
                              host: e.target.value
                            }
                            setSmtpAccounts(updated)
                          }}
                          required
                        />
                        <label htmlFor="smtpHost">SMTP Host (e.g. smtp.gmail.com)</label>
                      </div>
                      <div className="form-group input-animated">
                        <input
                          type="text"
                          id="smtpPort"
                          placeholder=" "
                          value={smtpAccounts[activeSmtpTab].port}
                          onChange={(e) => {
                            const updated = [...smtpAccounts]
                            updated[activeSmtpTab] = {
                              ...updated[activeSmtpTab],
                              port: e.target.value
                            }
                            setSmtpAccounts(updated)
                          }}
                          required
                        />
                        <label htmlFor="smtpPort">Port (465, 587)</label>
                      </div>
                    </div>

                    <div className="settings-grid-2">
                      <div className="form-group input-animated">
                        <input
                          type="text"
                          id="smtpUser"
                          placeholder=" "
                          value={smtpAccounts[activeSmtpTab].user}
                          onChange={(e) => {
                            const updated = [...smtpAccounts]
                            updated[activeSmtpTab] = {
                              ...updated[activeSmtpTab],
                              user: e.target.value
                            }
                            setSmtpAccounts(updated)
                          }}
                          required
                        />
                        <label htmlFor="smtpUser">Username / Email</label>
                      </div>
                      <div className="form-group input-animated">
                        <input
                          type="password"
                          id="smtpPass"
                          placeholder=" "
                          value={smtpAccounts[activeSmtpTab].pass}
                          onChange={(e) => {
                            const updated = [...smtpAccounts]
                            updated[activeSmtpTab] = {
                              ...updated[activeSmtpTab],
                              pass: e.target.value
                            }
                            setSmtpAccounts(updated)
                          }}
                          required
                        />
                        <label htmlFor="smtpPass">Password / App Password</label>
                      </div>
                    </div>

                    <div
                      className="settings-header"
                      style={{ marginTop: '20px', display: 'flex', gap: '20px', flexWrap: 'wrap' }}
                    >
                      <label className="checkbox-container">
                        <input
                          type="checkbox"
                          checked={smtpAccounts[activeSmtpTab].isTagging}
                          onChange={(e) => {
                            const updated = [...smtpAccounts]
                            updated[activeSmtpTab] = {
                              ...updated[activeSmtpTab],
                              isTagging: e.target.checked
                            }
                            setSmtpAccounts(updated)
                          }}
                        />
                        <span className="checkmark"></span>
                        Enable for Tagging
                      </label>
                      <label className="checkbox-container">
                        <input
                          type="checkbox"
                          checked={smtpAccounts[activeSmtpTab].isRetagging}
                          onChange={(e) => {
                            const updated = [...smtpAccounts]
                            updated[activeSmtpTab] = {
                              ...updated[activeSmtpTab],
                              isRetagging: e.target.checked
                            }
                            setSmtpAccounts(updated)
                          }}
                        />
                        <span className="checkmark"></span>
                        Enable for Retagging
                      </label>
                      <label className="checkbox-container">
                        <input
                          type="checkbox"
                          checked={smtpAccounts[activeSmtpTab].isClaim}
                          onChange={(e) => {
                            const updated = [...smtpAccounts]
                            updated[activeSmtpTab] = {
                              ...updated[activeSmtpTab],
                              isClaim: e.target.checked
                            }
                            setSmtpAccounts(updated)
                          }}
                        />
                        <span className="checkmark"></span>
                        Enable for Claim
                      </label>
                    </div>

                    <div className="settings-header" style={{ marginTop: '20px' }}>
                      <label className="checkbox-container">
                        <input
                          type="checkbox"
                          checked={smtpAccounts[activeSmtpTab].enableImapPop}
                          onChange={(e) => {
                            const updated = [...smtpAccounts]
                            updated[activeSmtpTab] = {
                              ...updated[activeSmtpTab],
                              enableImapPop: e.target.checked
                            }
                            setSmtpAccounts(updated)
                          }}
                        />
                        <span className="checkmark"></span>
                        Enable IMAP / POP (Optional)
                      </label>
                    </div>

                    {smtpAccounts[activeSmtpTab].enableImapPop && (
                      <div
                        className="settings-grid-2 animate-fade-in"
                        style={{ marginTop: '10px' }}
                      >
                        <div className="form-group input-animated">
                          <input
                            type="text"
                            id="imapHost"
                            placeholder=" "
                            value={smtpAccounts[activeSmtpTab].imapHost}
                            onChange={(e) => {
                              const updated = [...smtpAccounts]
                              updated[activeSmtpTab] = {
                                ...updated[activeSmtpTab],
                                imapHost: e.target.value
                              }
                              setSmtpAccounts(updated)
                            }}
                          />
                          <label htmlFor="imapHost">IMAP Host</label>
                        </div>
                        <div className="form-group input-animated">
                          <input
                            type="text"
                            id="imapPort"
                            placeholder=" "
                            value={smtpAccounts[activeSmtpTab].imapPort}
                            onChange={(e) => {
                              const updated = [...smtpAccounts]
                              updated[activeSmtpTab] = {
                                ...updated[activeSmtpTab],
                                imapPort: e.target.value
                              }
                              setSmtpAccounts(updated)
                            }}
                          />
                          <label htmlFor="imapPort">IMAP Port (993)</label>
                        </div>
                        <div className="form-group input-animated">
                          <input
                            type="text"
                            id="popHost"
                            placeholder=" "
                            value={smtpAccounts[activeSmtpTab].popHost}
                            onChange={(e) => {
                              const updated = [...smtpAccounts]
                              updated[activeSmtpTab] = {
                                ...updated[activeSmtpTab],
                                popHost: e.target.value
                              }
                              setSmtpAccounts(updated)
                            }}
                          />
                          <label htmlFor="popHost">POP Host</label>
                        </div>
                        <div className="form-group input-animated">
                          <input
                            type="text"
                            id="popPort"
                            placeholder=" "
                            value={smtpAccounts[activeSmtpTab].popPort}
                            onChange={(e) => {
                              const updated = [...smtpAccounts]
                              updated[activeSmtpTab] = {
                                ...updated[activeSmtpTab],
                                popPort: e.target.value
                              }
                              setSmtpAccounts(updated)
                            }}
                          />
                          <label htmlFor="popPort">POP Port (995)</label>
                        </div>
                      </div>
                    )}

                    <div className="template-actions" style={{ marginTop: 'auto' }}>
                      {smtpStatus !== 'idle' && (
                        <div className={`connection-status ${smtpStatus}`}>
                          <span className="status-dot"></span>
                          <span className="status-text">
                            {smtpStatus.charAt(0).toUpperCase() + smtpStatus.slice(1)}
                          </span>
                        </div>
                      )}
                      <div style={{ flex: 1 }}></div>
                      <button className="btn-secondary" onClick={handleTestConnection}>
                        {smtpStatus === 'loading' ? 'Testing...' : 'Test SMTP Connection'}
                      </button>
                      <button className="btn-primary" onClick={saveSmtpSettings}>
                        Save SMTP Settings
                      </button>
                    </div>
                  </div>
                )}

                {activeManageSection === 'settings' && (
                  <div className="section-container animate-fade-in">
                    <div
                      className="form-group input-animated"
                      style={{ marginBottom: '15px', marginTop: '25px' }}
                    >
                      <input
                        type="text"
                        id="alwaysCc"
                        placeholder=" "
                        value={alwaysCc}
                        onChange={(e) => setAlwaysCc(e.target.value)}
                        onBlur={saveSmtpSettings}
                      />
                      <label htmlFor="alwaysCc">Always CC Email ID(s)</label>
                    </div>
                    <div className="upload-section">
                      <p style={{ color: 'var(--text-light)', fontSize: '14px', margin: 0 }}>
                        Upload Excel file to manage recipient details.
                      </p>
                      <label className="btn-primary ripple-effect" style={{ cursor: 'pointer' }}>
                        <Upload size={16} />
                        <span>Upload Excel File</span>
                        <input
                          type="file"
                          accept=".xlsx, .xls"
                          onChange={handleExcelUpload}
                          style={{ display: 'none' }}
                        />
                      </label>
                    </div>

                    <form onSubmit={handleAddEntry} className="manual-entry-form">
                      <div className="form-group input-animated">
                        <input
                          type="text"
                          id="manualName"
                          placeholder=" "
                          value={newEntry.greetingName}
                          onChange={(e) =>
                            setNewEntry({ ...newEntry, greetingName: e.target.value })
                          }
                          required
                        />
                        <label htmlFor="manualName">Professional Greeting Name</label>
                      </div>

                      <div className="form-group input-animated">
                        <select
                          id="manualType"
                          value={newEntry.type}
                          onChange={(e) => setNewEntry({ ...newEntry, type: e.target.value })}
                          required
                        >
                          <option value="" disabled hidden></option>
                          <option value="Insurance RM">Insurance RM</option>
                          <option value="Bank RM">Bank RM</option>
                          <option value="Team Leader">Team Leader</option>
                        </select>
                        <label htmlFor="manualType" className="select-label">
                          Email ID Type
                        </label>
                      </div>

                      <div className="form-group input-animated">
                        <input
                          type="text"
                          id="manualEmail"
                          placeholder=" "
                          value={newEntry.email}
                          onChange={(e) => setNewEntry({ ...newEntry, email: e.target.value })}
                          required
                        />
                        <label htmlFor="manualEmail">Email Address</label>
                      </div>

                      <div className="form-group input-animated">
                        <input
                          type="text"
                          id="manualCc"
                          placeholder=" "
                          value={newEntry.ccEmail}
                          onChange={(e) => setNewEntry({ ...newEntry, ccEmail: e.target.value })}
                        />
                        <label htmlFor="manualCc">CC Email</label>
                      </div>

                      <div className="form-actions">
                        <button type="submit" className="btn-primary ripple-effect">
                          <span>{editingIndex !== null ? 'Update Member' : 'Add Member'}</span>
                          {editingIndex !== null ? (
                            <Edit2 size={16} />
                          ) : (
                            <svg
                              xmlns="http://www.w3.org/2000/svg"
                              width="16"
                              height="16"
                              viewBox="0 0 24 24"
                              fill="none"
                              stroke="currentColor"
                              strokeWidth="2"
                              strokeLinecap="round"
                              strokeLinejoin="round"
                            >
                              <line x1="12" y1="5" x2="12" y2="19"></line>
                              <line x1="5" y1="12" x2="19" y2="12"></line>
                            </svg>
                          )}
                        </button>
                        {editingIndex !== null && (
                          <button
                            type="button"
                            className="btn-secondary ripple-effect"
                            onClick={cancelEdit}
                          >
                            <X size={16} />
                            <span>Cancel Edit</span>
                          </button>
                        )}
                      </div>
                    </form>

                    <div className="manage-controls" style={{ marginTop: '20px' }}>
                      <div className="search-bar">
                        <Search size={18} />
                        <input
                          type="text"
                          placeholder="Search by name, email or type..."
                          value={searchTerm}
                          onChange={(e) => setSearchTerm(e.target.value)}
                        />
                      </div>
                      {selectedIndices.length > 0 && (
                        <button
                          className="btn-secondary danger-hover animate-fade-in"
                          onClick={handleBulkDelete}
                          style={{ padding: '8px 15px', display: 'flex', gap: '8px' }}
                        >
                          <Trash2 size={16} />
                          <span>Delete Selected ({selectedIndices.length})</span>
                        </button>
                      )}
                    </div>

                    <div className="data-list-container">
                      {manageData.length === 0 ? (
                        <div
                          style={{
                            textAlign: 'center',
                            padding: '40px',
                            color: 'var(--text-light)',
                            fontStyle: 'italic'
                          }}
                        >
                          No data uploaded yet.
                        </div>
                      ) : (
                        <div className="data-table-wrapper">
                          <table className="data-table">
                            <thead>
                              <tr>
                                <th style={{ width: '40px' }}>
                                  <label className="checkbox-container no-label">
                                    <input
                                      type="checkbox"
                                      checked={selectedIndices.length === manageData.length}
                                      onChange={toggleSelectAll}
                                    />
                                    <span className="checkmark"></span>
                                  </label>
                                </th>
                                <th>Professional Greeting Name</th>
                                <th>Email ID Type</th>
                                <th>Email</th>
                                <th>CC Email</th>
                                <th style={{ textAlign: 'right' }}>Actions</th>
                              </tr>
                            </thead>
                            <tbody>
                              {manageData
                                .map((item, idx) => ({ ...item, originalIdx: idx }))
                                .filter((item) => {
                                  if (!searchTerm) return true
                                  const search = searchTerm.toLowerCase()
                                  return (
                                    item.greetingName.toLowerCase().includes(search) ||
                                    item.email.toLowerCase().includes(search) ||
                                    item.type.toLowerCase().includes(search) ||
                                    item.ccEmail.toLowerCase().includes(search)
                                  )
                                })
                                .map((item) => (
                                  <tr
                                    key={item.originalIdx}
                                    className={editingIndex === item.originalIdx ? 'editing' : ''}
                                  >
                                    <td>
                                      <label className="checkbox-container no-label">
                                        <input
                                          type="checkbox"
                                          checked={selectedIndices.includes(item.originalIdx)}
                                          onChange={() => toggleSelect(item.originalIdx)}
                                        />
                                        <span className="checkmark"></span>
                                      </label>
                                    </td>
                                    <td>{item.greetingName}</td>
                                    <td>{item.type}</td>
                                    <td>{item.email}</td>
                                    <td>{item.ccEmail}</td>
                                    <td style={{ textAlign: 'right' }}>
                                      <div
                                        className="table-actions"
                                        style={{
                                          display: 'flex',
                                          gap: '10px',
                                          justifyContent: 'flex-end'
                                        }}
                                      >
                                        <button
                                          className={`icon-btn ${editingIndex === item.originalIdx ? 'active' : ''}`}
                                          onClick={() => handleEditEntry(item.originalIdx)}
                                          title="Edit Member"
                                        >
                                          <Edit2 size={16} />
                                        </button>
                                        <button
                                          className="icon-btn danger"
                                          onClick={() => handleDeleteEntry(item.originalIdx)}
                                          title="Delete Member"
                                        >
                                          <Trash2 size={16} />
                                        </button>
                                      </div>
                                    </td>
                                  </tr>
                                ))}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  </div>
                )}

                {activeManageSection === 'template' && (
                  <div className="section-container animate-fade-in">
                    <div className="template-nav">
                      <button
                        className={`sub-nav-item ${activeTemplateTab === 'tagging' ? 'active' : ''}`}
                        onClick={() => setActiveTemplateTab('tagging')}
                      >
                        Tagging Templates
                      </button>
                      <button
                        className={`sub-nav-item ${activeTemplateTab === 'retagging' ? 'active' : ''}`}
                        onClick={() => setActiveTemplateTab('retagging')}
                      >
                        Retagging Templates
                      </button>
                      <button
                        className={`sub-nav-item ${activeTemplateTab === 'claim' ? 'active' : ''}`}
                        onClick={() => setActiveTemplateTab('claim')}
                      >
                        Claim Templates
                      </button>
                    </div>

                    <div className="multi-template-container">
                      {templates[activeTemplateTab].map((template, idx) => (
                        <div key={idx} className="template-box">
                          <label className="editor-label">
                            Template {idx + 1} ({activeTemplateTab})
                          </label>
                          <textarea
                            className="template-textarea mini-editor"
                            value={template.body}
                            onChange={(e) =>
                              handleTemplateChange(activeTemplateTab, idx, e.target.value)
                            }
                            placeholder={`Enter template ${idx + 1} text...`}
                          />
                          <div className="signature-preview-container">
                            <span className="signature-preview-label">Appended Signature:</span>
                            <div className="signature-preview-box">
                              {(() => {
                                const ds = signatures[activeTemplateTab]?.find((s) => s.isDefault)
                                if (!ds)
                                  return (
                                    <i>No default signature selected for {activeTemplateTab}</i>
                                  )

                                const contactLine = [
                                  ds.mobile ? `M: ${ds.mobile}` : '',
                                  ds.email ? ds.email : ''
                                ]
                                  .filter(Boolean)
                                  .join(', ')

                                const displayLines = [
                                  ds.regards ? ds.regards + ',' : '',
                                  ds.name,
                                  ds.designation,
                                  ds.company,
                                  ds.address,
                                  contactLine
                                ].filter(Boolean)

                                return displayLines.map((line, lidx) => {
                                  let className = ''
                                  if (line === ds.name) className = 'label-blue'
                                  if (line === ds.company) className = 'label-red'
                                  if (line === ds.designation) className = 'label-navy'
                                  return (
                                    <div key={lidx} className={className}>
                                      {line}
                                    </div>
                                  )
                                })
                              })()}
                            </div>
                          </div>
                          <div className="default-toggle">
                            <label className="checkbox-container">
                              <input
                                type="checkbox"
                                checked={template.isDefault}
                                onChange={() => handleSetDefault(activeTemplateTab, idx)}
                              />
                              <span className="checkmark"></span>
                              Set as Default
                            </label>
                          </div>
                        </div>
                      ))}
                    </div>

                    <div className="template-actions sticky-actions">
                      {saveStatus === 'saving' && (
                        <span className="save-indicator saving">Saving...</span>
                      )}
                      {saveStatus === 'saved' && (
                        <span className="save-indicator saved">● Saved</span>
                      )}
                      <button className="btn-secondary" onClick={resetTemplate}>
                        Reset All{' '}
                        {activeTemplateTab.charAt(0).toUpperCase() + activeTemplateTab.slice(1)}
                      </button>
                      <button className="btn-primary" onClick={saveTemplate}>
                        Save All Templates
                      </button>
                    </div>
                  </div>
                )}

                {activeManageSection === 'signatures' && (
                  <div className="section-container animate-fade-in">
                    <div className="template-nav">
                      <button
                        className={`sub-nav-item ${activeSignatureTab === 'tagging' ? 'active' : ''}`}
                        onClick={() => setActiveSignatureTab('tagging')}
                      >
                        Tagging Signatures
                      </button>
                      <button
                        className={`sub-nav-item ${activeSignatureTab === 'retagging' ? 'active' : ''}`}
                        onClick={() => setActiveSignatureTab('retagging')}
                      >
                        Retagging Signatures
                      </button>
                      <button
                        className={`sub-nav-item ${activeSignatureTab === 'claim' ? 'active' : ''}`}
                        onClick={() => setActiveSignatureTab('claim')}
                      >
                        Claim Signatures
                      </button>
                    </div>

                    <div className="multi-template-container">
                      {signatures[activeSignatureTab].map((sig, idx) => (
                        <div key={idx} className="signature-box template-box">
                          <label className="editor-label">
                            Signature {idx + 1} ({activeSignatureTab})
                          </label>
                          <div className="signature-form-grid">
                            <div
                              style={{
                                display: 'flex',
                                gap: '15px',
                                alignItems: 'flex-start'
                              }}
                            >
                              <div
                                className="form-group input-animated sm-input"
                                style={{ flex: 1, marginBottom: '20px' }}
                              >
                                <input
                                  type="text"
                                  placeholder=" "
                                  value={sig.regards}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'regards',
                                      e.target.value
                                    )
                                  }
                                />
                                <label>Regards Text</label>
                              </div>
                              <div
                                style={{
                                  display: 'flex',
                                  gap: '8px',
                                  alignItems: 'center',
                                  paddingTop: '8px'
                                }}
                              >
                                <select
                                  value={sig.regardsSize || '14px'}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'regardsSize',
                                      e.target.value
                                    )
                                  }
                                  style={{
                                    fontSize: '11px',
                                    padding: '3px 6px',
                                    borderRadius: '4px',
                                    border: '1px solid #e2e8f0',
                                    background: 'white',
                                    cursor: 'pointer'
                                  }}
                                >
                                  {[
                                    '12px',
                                    '13px',
                                    '14px',
                                    '15px',
                                    '16px',
                                    '17px',
                                    '18px',
                                    '19px',
                                    '20px'
                                  ].map((s) => (
                                    <option key={s} value={s}>
                                      {s}
                                    </option>
                                  ))}
                                </select>
                                <div style={{ display: 'flex', gap: '5px' }}>
                                  {['#333333', '#3b82f6', '#ef4444', '#10b981', '#4f46e5'].map(
                                    (c) => (
                                      <button
                                        key={c}
                                        onClick={() =>
                                          handleSignatureChange(
                                            activeSignatureTab,
                                            idx,
                                            'regardsColor',
                                            c
                                          )
                                        }
                                        style={{
                                          width: '16px',
                                          height: '16px',
                                          borderRadius: '4px',
                                          backgroundColor: c,
                                          border:
                                            sig.regardsColor === c
                                              ? '2px solid white'
                                              : '1px solid #e2e8f0',
                                          outline:
                                            sig.regardsColor === c ? '2px solid #3b82f6' : 'none',
                                          cursor: 'pointer',
                                          padding: 0,
                                          transform:
                                            sig.regardsColor === c ? 'scale(1.15)' : 'scale(1)',
                                          boxShadow:
                                            sig.regardsColor === c
                                              ? '0 2px 4px rgba(0,0,0,0.15)'
                                              : 'none',
                                          zIndex: sig.regardsColor === c ? 2 : 1,
                                          transition: 'all 0.2s ease'
                                        }}
                                        title={c}
                                      />
                                    )
                                  )}
                                </div>
                              </div>
                            </div>
                            <div
                              style={{
                                display: 'flex',
                                gap: '15px',
                                alignItems: 'flex-start'
                              }}
                            >
                              <div
                                className="form-group input-animated sm-input"
                                style={{ flex: 1, marginBottom: '20px' }}
                              >
                                <input
                                  type="text"
                                  placeholder=" "
                                  value={sig.name}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'name',
                                      e.target.value
                                    )
                                  }
                                />
                                <label className="label-blue">Full Name</label>
                              </div>
                              <div
                                style={{
                                  display: 'flex',
                                  gap: '8px',
                                  alignItems: 'center',
                                  paddingTop: '8px'
                                }}
                              >
                                <select
                                  value={sig.nameSize || '16px'}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'nameSize',
                                      e.target.value
                                    )
                                  }
                                  style={{
                                    fontSize: '11px',
                                    padding: '3px 6px',
                                    borderRadius: '4px',
                                    border: '1px solid #e2e8f0',
                                    background: 'white',
                                    cursor: 'pointer'
                                  }}
                                >
                                  {[
                                    '12px',
                                    '13px',
                                    '14px',
                                    '15px',
                                    '16px',
                                    '17px',
                                    '18px',
                                    '19px',
                                    '20px'
                                  ].map((s) => (
                                    <option key={s} value={s}>
                                      {s}
                                    </option>
                                  ))}
                                </select>
                                <div style={{ display: 'flex', gap: '5px' }}>
                                  {['#333333', '#3b82f6', '#ef4444', '#10b981', '#4f46e5'].map(
                                    (c) => (
                                      <button
                                        key={c}
                                        onClick={() =>
                                          handleSignatureChange(
                                            activeSignatureTab,
                                            idx,
                                            'nameColor',
                                            c
                                          )
                                        }
                                        style={{
                                          width: '16px',
                                          height: '16px',
                                          borderRadius: '4px',
                                          backgroundColor: c,
                                          border:
                                            sig.nameColor === c
                                              ? '2px solid white'
                                              : '1px solid #e2e8f0',
                                          outline:
                                            sig.nameColor === c ? '2px solid #3b82f6' : 'none',
                                          cursor: 'pointer',
                                          padding: 0,
                                          transform:
                                            sig.nameColor === c ? 'scale(1.15)' : 'scale(1)',
                                          boxShadow:
                                            sig.nameColor === c
                                              ? '0 2px 4px rgba(0,0,0,0.15)'
                                              : 'none',
                                          zIndex: sig.nameColor === c ? 2 : 1,
                                          transition: 'all 0.2s ease'
                                        }}
                                        title={c}
                                      />
                                    )
                                  )}
                                </div>
                              </div>
                            </div>
                            <div
                              style={{
                                display: 'flex',
                                gap: '15px',
                                alignItems: 'flex-start'
                              }}
                            >
                              <div
                                className="form-group input-animated sm-input"
                                style={{ flex: 1, marginBottom: '20px' }}
                              >
                                <input
                                  type="text"
                                  placeholder=" "
                                  value={sig.designation}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'designation',
                                      e.target.value
                                    )
                                  }
                                />
                                <label className="label-navy">Designation</label>
                              </div>
                              <div
                                style={{
                                  display: 'flex',
                                  gap: '8px',
                                  alignItems: 'center',
                                  paddingTop: '8px'
                                }}
                              >
                                <select
                                  value={sig.designationSize || '14px'}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'designationSize',
                                      e.target.value
                                    )
                                  }
                                  style={{
                                    fontSize: '11px',
                                    padding: '3px 6px',
                                    borderRadius: '4px',
                                    border: '1px solid #e2e8f0',
                                    background: 'white',
                                    cursor: 'pointer'
                                  }}
                                >
                                  {[
                                    '12px',
                                    '13px',
                                    '14px',
                                    '15px',
                                    '16px',
                                    '17px',
                                    '18px',
                                    '19px',
                                    '20px'
                                  ].map((s) => (
                                    <option key={s} value={s}>
                                      {s}
                                    </option>
                                  ))}
                                </select>
                                <div style={{ display: 'flex', gap: '5px' }}>
                                  {['#333333', '#3b82f6', '#ef4444', '#10b981', '#4f46e5'].map(
                                    (c) => (
                                      <button
                                        key={c}
                                        onClick={() =>
                                          handleSignatureChange(
                                            activeSignatureTab,
                                            idx,
                                            'designationColor',
                                            c
                                          )
                                        }
                                        style={{
                                          width: '16px',
                                          height: '16px',
                                          borderRadius: '4px',
                                          backgroundColor: c,
                                          border:
                                            sig.designationColor === c
                                              ? '2px solid white'
                                              : '1px solid #e2e8f0',
                                          outline:
                                            sig.designationColor === c
                                              ? '2px solid #3b82f6'
                                              : 'none',
                                          cursor: 'pointer',
                                          padding: 0,
                                          transform:
                                            sig.designationColor === c ? 'scale(1.15)' : 'scale(1)',
                                          boxShadow:
                                            sig.designationColor === c
                                              ? '0 2px 4px rgba(0,0,0,0.15)'
                                              : 'none',
                                          zIndex: sig.designationColor === c ? 2 : 1,
                                          transition: 'all 0.2s ease'
                                        }}
                                        title={c}
                                      />
                                    )
                                  )}
                                </div>
                              </div>
                            </div>
                            <div
                              style={{
                                display: 'flex',
                                gap: '15px',
                                alignItems: 'flex-start'
                              }}
                            >
                              <div
                                className="form-group input-animated sm-input"
                                style={{ flex: 1, marginBottom: '20px' }}
                              >
                                <input
                                  type="text"
                                  placeholder=" "
                                  value={sig.company}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'company',
                                      e.target.value
                                    )
                                  }
                                />
                                <label className="label-red">Company Name</label>
                              </div>
                              <div
                                style={{
                                  display: 'flex',
                                  gap: '8px',
                                  alignItems: 'center',
                                  paddingTop: '8px'
                                }}
                              >
                                <select
                                  value={sig.companySize || '14px'}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'companySize',
                                      e.target.value
                                    )
                                  }
                                  style={{
                                    fontSize: '11px',
                                    padding: '3px 6px',
                                    borderRadius: '4px',
                                    border: '1px solid #e2e8f0',
                                    background: 'white',
                                    cursor: 'pointer'
                                  }}
                                >
                                  {[
                                    '12px',
                                    '13px',
                                    '14px',
                                    '15px',
                                    '16px',
                                    '17px',
                                    '18px',
                                    '19px',
                                    '20px'
                                  ].map((s) => (
                                    <option key={s} value={s}>
                                      {s}
                                    </option>
                                  ))}
                                </select>
                                <div style={{ display: 'flex', gap: '5px' }}>
                                  {['#333333', '#3b82f6', '#ef4444', '#10b981', '#4f46e5'].map(
                                    (c) => (
                                      <button
                                        key={c}
                                        onClick={() =>
                                          handleSignatureChange(
                                            activeSignatureTab,
                                            idx,
                                            'companyColor',
                                            c
                                          )
                                        }
                                        style={{
                                          width: '16px',
                                          height: '16px',
                                          borderRadius: '4px',
                                          backgroundColor: c,
                                          border:
                                            sig.companyColor === c
                                              ? '2px solid white'
                                              : '1px solid #e2e8f0',
                                          outline:
                                            sig.companyColor === c ? '2px solid #3b82f6' : 'none',
                                          cursor: 'pointer',
                                          padding: 0,
                                          transform:
                                            sig.companyColor === c ? 'scale(1.15)' : 'scale(1)',
                                          boxShadow:
                                            sig.companyColor === c
                                              ? '0 2px 4px rgba(0,0,0,0.15)'
                                              : 'none',
                                          zIndex: sig.companyColor === c ? 2 : 1,
                                          transition: 'all 0.2s ease'
                                        }}
                                        title={c}
                                      />
                                    )
                                  )}
                                </div>
                              </div>
                            </div>
                            <div
                              style={{
                                display: 'flex',
                                gap: '15px',
                                alignItems: 'flex-start'
                              }}
                            >
                              <div
                                className="form-group input-animated sm-input full-width"
                                style={{ flex: 1, marginBottom: '20px' }}
                              >
                                <textarea
                                  placeholder=" "
                                  value={sig.address}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'address',
                                      e.target.value
                                    )
                                  }
                                />
                                <label className="label-navy">Address</label>
                              </div>
                              <div
                                style={{
                                  display: 'flex',
                                  gap: '8px',
                                  alignItems: 'center',
                                  paddingTop: '8px'
                                }}
                              >
                                <select
                                  value={sig.addressSize || '13px'}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'addressSize',
                                      e.target.value
                                    )
                                  }
                                  style={{
                                    fontSize: '11px',
                                    padding: '3px 6px',
                                    borderRadius: '4px',
                                    border: '1px solid #e2e8f0',
                                    background: 'white',
                                    cursor: 'pointer'
                                  }}
                                >
                                  {[
                                    '12px',
                                    '13px',
                                    '14px',
                                    '15px',
                                    '16px',
                                    '17px',
                                    '18px',
                                    '19px',
                                    '20px'
                                  ].map((s) => (
                                    <option key={s} value={s}>
                                      {s}
                                    </option>
                                  ))}
                                </select>
                                <div style={{ display: 'flex', gap: '5px' }}>
                                  {['#333333', '#3b82f6', '#ef4444', '#10b981', '#4f46e5'].map(
                                    (c) => (
                                      <button
                                        key={c}
                                        onClick={() =>
                                          handleSignatureChange(
                                            activeSignatureTab,
                                            idx,
                                            'addressColor',
                                            c
                                          )
                                        }
                                        style={{
                                          width: '16px',
                                          height: '16px',
                                          borderRadius: '4px',
                                          backgroundColor: c,
                                          border:
                                            sig.addressColor === c
                                              ? '2px solid white'
                                              : '1px solid #e2e8f0',
                                          outline:
                                            sig.addressColor === c ? '2px solid #3b82f6' : 'none',
                                          cursor: 'pointer',
                                          padding: 0,
                                          transform:
                                            sig.addressColor === c ? 'scale(1.15)' : 'scale(1)',
                                          boxShadow:
                                            sig.addressColor === c
                                              ? '0 2px 4px rgba(0,0,0,0.15)'
                                              : 'none',
                                          zIndex: sig.addressColor === c ? 2 : 1,
                                          transition: 'all 0.2s ease'
                                        }}
                                        title={c}
                                      />
                                    )
                                  )}
                                </div>
                              </div>
                            </div>
                            <div
                              style={{
                                display: 'flex',
                                gap: '15px',
                                alignItems: 'flex-start'
                              }}
                            >
                              <div
                                className="form-group input-animated sm-input"
                                style={{ flex: 1, marginBottom: '20px' }}
                              >
                                <input
                                  type="text"
                                  placeholder=" "
                                  value={sig.mobile}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'mobile',
                                      e.target.value
                                    )
                                  }
                                />
                                <label className="label-navy">Mobile</label>
                              </div>
                              <div
                                style={{
                                  display: 'flex',
                                  gap: '8px',
                                  alignItems: 'center',
                                  paddingTop: '8px'
                                }}
                              >
                                <select
                                  value={sig.mobileSize || '13px'}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'mobileSize',
                                      e.target.value
                                    )
                                  }
                                  style={{
                                    fontSize: '11px',
                                    padding: '3px 6px',
                                    borderRadius: '4px',
                                    border: '1px solid #e2e8f0',
                                    background: 'white',
                                    cursor: 'pointer'
                                  }}
                                >
                                  {[
                                    '12px',
                                    '13px',
                                    '14px',
                                    '15px',
                                    '16px',
                                    '17px',
                                    '18px',
                                    '19px',
                                    '20px'
                                  ].map((s) => (
                                    <option key={s} value={s}>
                                      {s}
                                    </option>
                                  ))}
                                </select>
                                <div style={{ display: 'flex', gap: '5px' }}>
                                  {['#333333', '#3b82f6', '#ef4444', '#10b981', '#4f46e5'].map(
                                    (c) => (
                                      <button
                                        key={c}
                                        onClick={() =>
                                          handleSignatureChange(
                                            activeSignatureTab,
                                            idx,
                                            'mobileColor',
                                            c
                                          )
                                        }
                                        style={{
                                          width: '16px',
                                          height: '16px',
                                          borderRadius: '4px',
                                          backgroundColor: c,
                                          border:
                                            sig.mobileColor === c
                                              ? '2px solid white'
                                              : '1px solid #e2e8f0',
                                          outline:
                                            sig.mobileColor === c ? '2px solid #3b82f6' : 'none',
                                          cursor: 'pointer',
                                          padding: 0,
                                          transform:
                                            sig.mobileColor === c ? 'scale(1.15)' : 'scale(1)',
                                          boxShadow:
                                            sig.mobileColor === c
                                              ? '0 2px 4px rgba(0,0,0,0.15)'
                                              : 'none',
                                          zIndex: sig.mobileColor === c ? 2 : 1,
                                          transition: 'all 0.2s ease'
                                        }}
                                        title={c}
                                      />
                                    )
                                  )}
                                </div>
                              </div>
                            </div>
                            <div
                              style={{
                                display: 'flex',
                                gap: '15px',
                                alignItems: 'flex-start'
                              }}
                            >
                              <div
                                className="form-group input-animated sm-input"
                                style={{ flex: 1, marginBottom: '20px' }}
                              >
                                <input
                                  type="text"
                                  placeholder=" "
                                  value={sig.email}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'email',
                                      e.target.value
                                    )
                                  }
                                />
                                <label className="label-navy">Email</label>
                              </div>
                              <div
                                style={{
                                  display: 'flex',
                                  gap: '8px',
                                  alignItems: 'center',
                                  paddingTop: '8px'
                                }}
                              >
                                <select
                                  value={sig.emailSize || '13px'}
                                  onChange={(e) =>
                                    handleSignatureChange(
                                      activeSignatureTab,
                                      idx,
                                      'emailSize',
                                      e.target.value
                                    )
                                  }
                                  style={{
                                    fontSize: '11px',
                                    padding: '3px 6px',
                                    borderRadius: '4px',
                                    border: '1px solid #e2e8f0',
                                    background: 'white',
                                    cursor: 'pointer'
                                  }}
                                >
                                  {[
                                    '12px',
                                    '13px',
                                    '14px',
                                    '15px',
                                    '16px',
                                    '17px',
                                    '18px',
                                    '19px',
                                    '20px'
                                  ].map((s) => (
                                    <option key={s} value={s}>
                                      {s}
                                    </option>
                                  ))}
                                </select>
                                <div style={{ display: 'flex', gap: '5px' }}>
                                  {['#333333', '#3b82f6', '#ef4444', '#10b981', '#4f46e5'].map(
                                    (c) => (
                                      <button
                                        key={c}
                                        onClick={() =>
                                          handleSignatureChange(
                                            activeSignatureTab,
                                            idx,
                                            'emailColor',
                                            c
                                          )
                                        }
                                        style={{
                                          width: '16px',
                                          height: '16px',
                                          borderRadius: '4px',
                                          backgroundColor: c,
                                          border:
                                            sig.emailColor === c
                                              ? '2px solid white'
                                              : '1px solid #e2e8f0',
                                          outline:
                                            sig.emailColor === c ? '2px solid #3b82f6' : 'none',
                                          cursor: 'pointer',
                                          padding: 0,
                                          transform:
                                            sig.emailColor === c ? 'scale(1.15)' : 'scale(1)',
                                          boxShadow:
                                            sig.emailColor === c
                                              ? '0 2px 4px rgba(0,0,0,0.15)'
                                              : 'none',
                                          zIndex: sig.emailColor === c ? 2 : 1,
                                          transition: 'all 0.2s ease'
                                        }}
                                        title={c}
                                      />
                                    )
                                  )}
                                </div>
                              </div>
                            </div>
                          </div>

                          <div className="default-toggle">
                            <label className="checkbox-container">
                              <input
                                type="checkbox"
                                checked={sig.isDefault}
                                onChange={() => handleSetSignatureDefault(activeSignatureTab, idx)}
                              />
                              <span className="checkmark"></span>
                              Set as Default
                            </label>
                          </div>
                        </div>
                      ))}
                    </div>

                    <div className="template-actions sticky-actions">
                      {saveStatus === 'saving' && (
                        <span className="save-indicator saving">Saving...</span>
                      )}
                      {saveStatus === 'saved' && (
                        <span className="save-indicator saved">● Auto-saved</span>
                      )}
                      <button className="btn-secondary" onClick={resetSignature}>
                        Reset All {activeSignatureTab}
                      </button>
                      <button className="btn-primary" onClick={handleSaveSignatures}>
                        Save All Signatures
                      </button>
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}

          <div className="header-container compact-header">
            <div style={{ display: 'flex', alignItems: 'center', gap: '15px' }}>
              <img src={logo} alt="RRM Logo" className="app-logo" />
              <div className="status-indicator"></div>
              <h1
                className="header"
                style={{ margin: 0, padding: 0, border: 'none', fontSize: '20px' }}
              >
                Email Automation Tool
              </h1>
            </div>
            <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              <div
                className="form-group input-animated m-0"
                style={{ minWidth: '250px', marginBottom: 0 }}
              >
                <input
                  type="text"
                  id="rootFolderName"
                  placeholder=" "
                  readOnly
                  value={folderName}
                />
                <label htmlFor="rootFolderName">Root Folder</label>
              </div>
              <button className="btn-secondary" onClick={handleSelectFolder}>
                Browse
              </button>
            </div>
          </div>

          {parsedData.length > 0 && currentItem && (
            <div className="preview-container">
              <div className="carousel-nav">
                <button
                  className="btn-secondary nav-btn"
                  disabled={currentIndex === 0}
                  onClick={() => setCurrentIndex(currentIndex - 1)}
                >
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    width="16"
                    height="16"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <polyline points="15 18 9 12 15 6"></polyline>
                  </svg>
                  <span>Previous</span>
                </button>
                <div className="carousel-counter">
                  <b>{currentIndex + 1}</b> / {parsedData.length}
                </div>
                <span className={`status-badge status-${currentItem.status.toLowerCase()}`}>
                  {currentItem.status}
                </span>
                <button
                  className="btn-secondary nav-btn"
                  disabled={currentIndex === parsedData.length - 1}
                  onClick={() => setCurrentIndex(currentIndex + 1)}
                >
                  <span>Next</span>
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    width="16"
                    height="16"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <polyline points="9 18 15 12 9 6"></polyline>
                  </svg>
                </button>
              </div>

              <div
                className={`editor-grid ${currentItem.emailType === 'Claim' ? 'editor-grid-3' : ''}`}
              >
                <div className="form-group input-animated full-width">
                  <input
                    type="text"
                    id="itemName"
                    placeholder=" "
                    value={currentItem.folderName}
                    onChange={(e) => updateCurrentItem('folderName', e.target.value)}
                    required
                  />
                  <label htmlFor="itemName">Folder Name</label>
                </div>

                <div className="form-group input-animated">
                  <input
                    type="text"
                    id="itemSubject"
                    placeholder=" "
                    value={currentItem.subject}
                    onChange={(e) => updateCurrentItem('subject', e.target.value)}
                    required
                  />
                  <label htmlFor="itemSubject">Subject</label>
                </div>

                <div className="form-group input-animated">
                  <select
                    id="itemTemplate"
                    required
                    value={currentItem.emailType}
                    onChange={async (e) => {
                      const newType = e.target.value
                      updateCurrentItem('emailType', newType)

                      // Re-fetch email IDs and Claim Details for the new type
                      const subfolderPath = currentItem.rootDir + '/' + currentItem.folderName
                      const fetchedEmails = await window.api.getEmailIds(subfolderPath, newType)

                      if (fetchedEmails) {
                        setParsedData((prev) => {
                          const updated = [...prev]
                          const item = updated[currentIndex]
                          const updatedItem = {
                            ...item,
                            emailType: newType,
                            fetchedEmails: fetchedEmails
                            // We don't overwrite emailBody here to preserve any manual edits,
                            // but the user can use a 'Reset Template' button if needed.
                            // However, let's auto-refresh if it's the default placeholder.
                          }

                          // Optional: Auto-refresh body if it's currently empty or default
                          if (!item.emailBody || item.emailBody.includes('Respected Sir')) {
                            updatedItem.emailBody = generateEmailBodyContent(updatedItem)
                          }

                          updated[currentIndex] = updatedItem
                          return updated
                        })
                      }
                    }}
                  >
                    <option value="" disabled hidden></option>
                    <option value="Claim">Claim</option>
                    <option value="Tagging">Tagging</option>
                    <option value="Retagging">Retagging</option>
                  </select>
                  <label htmlFor="itemTemplate" className="select-label">
                    Email Type
                  </label>
                </div>

                <div className="form-group input-animated">
                  <input
                    type="text"
                    id="itemEmailTo"
                    placeholder=" "
                    className={currentItem.emailTo === 'No Email Id Found' ? 'text-danger' : ''}
                    value={currentItem.emailTo}
                    onChange={(e) => updateCurrentItem('emailTo', e.target.value)}
                    required
                  />
                  <label htmlFor="itemEmailTo">Email To</label>
                </div>

                <div className="form-group input-animated">
                  <input
                    type="text"
                    id="itemCcTo"
                    placeholder=" "
                    value={currentItem.ccTo}
                    onChange={(e) => updateCurrentItem('ccTo', e.target.value)}
                    required
                  />
                  <label htmlFor="itemCcTo">CC To</label>
                </div>

                <div className="editor-left-column">
                  <div className="form-group input-animated preview-body">
                    <textarea
                      id="itemEmailBody"
                      placeholder=" "
                      value={currentItem.emailBody}
                      onChange={(e) => updateCurrentItem('emailBody', e.target.value)}
                      required
                    />
                    <label htmlFor="itemEmailBody">Email Body</label>
                  </div>

                  <div className="signature-main-preview">
                    <span className="signature-preview-label">Appended Signature:</span>
                    <div
                      className="signature-preview-box"
                      dangerouslySetInnerHTML={{
                        __html: generateSignatureHtml(
                          signatures[currentItem.emailType.toLowerCase()]?.find((s) => s.isDefault)
                        )
                      }}
                    />
                  </div>
                </div>

                {currentItem.emailType === 'Claim' && (
                  <div className="editor-middle-column animate-fade-in">
                    {currentItem.fetchedEmails?.claimDetails ? (
                      <div className="claim-details-preview-section border-none">
                        <span className="signature-preview-label">Claim Details:</span>
                        <div
                          className="claim-table-preview"
                          dangerouslySetInnerHTML={{ __html: generateClaimTableHtml(currentItem) }}
                        />
                      </div>
                    ) : (
                      <div className="claim-details-error">
                        <span className="signature-preview-label">Claim Details:</span>
                        <div
                          className="no-data-warning"
                          style={{
                            color: '#ef4444',
                            padding: '15px',
                            border: '1px dashed #ef4444',
                            borderRadius: '8px',
                            marginTop: '10px',
                            fontSize: '13px'
                          }}
                        >
                          <strong>No &quot;Claim Data&quot; file found!</strong>
                          <br />
                          Looking for a file containing &quot;claim data&quot; in:
                          <br />
                          <small style={{ wordBreak: 'break-all' }}>{currentItem.folderName}</small>
                        </div>
                      </div>
                    )}
                  </div>
                )}

                <div className="editor-right-column">
                  <div className="form-group input-animated h-100">
                    <div className="attachments-box">
                      {attachments.length === 0 ? (
                        <div className="no-attachments">No valid attachments found</div>
                      ) : (
                        <div className="attachment-list">
                          {attachments.map((att, i) => {
                            const { icon, className } = getIcon(att.ext)
                            return (
                              <div key={i} className="att-item">
                                <span className={`att-icon ${className}`}>{icon}</span>
                                <span className="att-name" title={att.name}>
                                  {att.name}
                                </span>
                                <span className="att-size">({formatSize(att.size)})</span>
                                <button
                                  className="att-remove-btn"
                                  onClick={() => handleRemoveAttachment(att.name)}
                                  title="Remove attachment"
                                >
                                  <Trash2 size={14} />
                                </button>
                              </div>
                            )
                          })}
                        </div>
                      )}
                    </div>
                    <span
                      className={`fixed-float-label ${Number(totalAttachmentsSizeMB) > 5 ? 'size-warning' : ''}`}
                    >
                      Attachments {attachments.length > 0 && `(${totalAttachmentsSizeMB} MB)`}
                      {Number(totalAttachmentsSizeMB) > 5 && ' Is Too Big Email Might Get Bounced'}
                    </span>
                  </div>
                </div>
              </div>

              <div className="action-container duo-actions">
                <button
                  className="btn-secondary ripple-effect"
                  onClick={() => handleSendEmail(currentIndex)}
                >
                  Send Current Email
                </button>
                <button className="btn-primary ripple-effect" onClick={handleSendAll}>
                  <span>Send All Emails</span>
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    width="16"
                    height="16"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    className="btn-icon"
                  >
                    <line x1="22" y1="2" x2="11" y2="13"></line>
                    <polygon points="22 2 15 22 11 13 2 9 22 2"></polygon>
                  </svg>
                </button>
              </div>
            </div>
          )}
        </div>
      </div>
      {/* Batch Progress Modal */}
      {(isBatchSending || showBatchSummary) && (
        <div className="batch-modal-overlay">
          <div className="batch-modal glass-effect">
            <div className="batch-modal-header">
              <h2>{isBatchSending ? 'Sending Emails...' : 'Email Summary'}</h2>
              {isBatchSending && <div className="loading-spinner-small"></div>}
            </div>

            <div className="batch-modal-body">
              {isBatchSending ? (
                <>
                  <div className="current-item-info">
                    <span className="label">Processing:</span>
                    <span className="value">{batchProgress.currentItemName}</span>
                  </div>

                  <div className="progress-container">
                    <div
                      className="progress-bar-fill"
                      style={{ width: `${(batchProgress.current / batchProgress.total) * 100}%` }}
                    ></div>
                  </div>

                  <div className="progress-text">
                    {batchProgress.current} / {batchProgress.total} Emails Processed
                  </div>
                </>
              ) : (
                <div className="summary-stats">
                  <div className="stat-card success">
                    <span className="stat-value">{batchProgress.success}</span>
                    <span className="stat-label">Successfully Sent</span>
                  </div>
                  <div className="stat-card failed">
                    <span className="stat-value">{batchProgress.failed}</span>
                    <span className="stat-label">Failed to Send</span>
                  </div>
                  <div className="stat-card total">
                    <span className="stat-value">{batchProgress.total}</span>
                    <span className="stat-label">Total emails</span>
                  </div>
                </div>
              )}
            </div>

            <div className="batch-modal-footer">
              {isBatchSending ? (
                <button
                  className="btn-primary ripple-effect"
                  onClick={() => {
                    cancelBatchRef.current = true
                  }}
                  style={{ backgroundColor: '#ef4444' }}
                >
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    width="16"
                    height="16"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    className="btn-icon"
                  >
                    <line x1="18" y1="6" x2="6" y2="18"></line>
                    <line x1="6" y1="6" x2="18" y2="18"></line>
                  </svg>
                  <span>Cancel Batch</span>
                </button>
              ) : (
                <button
                  className="btn-primary ripple-effect"
                  onClick={() => setShowBatchSummary(false)}
                >
                  Close Summary
                </button>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

export default App
