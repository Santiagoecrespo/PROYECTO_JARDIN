import { useEffect, useMemo, useRef, useState } from 'react'
import './App.css'

const formatCurrency = (value) => {
  const numericValue = Number(value) || 0
  const hasDecimals = Math.abs(numericValue % 1) > 0.001

  return `$ ${new Intl.NumberFormat('es-AR', {
    minimumFractionDigits: hasDecimals ? 1 : 0,
    maximumFractionDigits: 1,
  }).format(numericValue)}`
}

const formatPercent = (value) =>
  `${new Intl.NumberFormat('es-AR', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 3,
  }).format(Number(value) || 0)}%`

const formatDate = (value) => {
  if (!value) {
    return '--/--/----'
  }

  return new Intl.DateTimeFormat('es-AR', { timeZone: 'UTC' }).format(
    new Date(`${value}T00:00:00`),
  )
}

const roundToSingleDecimal = (value) => Math.round((value + Number.EPSILON) * 10) / 10
const normalizeSearchValue = (value) =>
  String(value || '')
    .normalize('NFD')
    .replace(/[^\w\s.-]|_/g, '')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim()
  const normalizeRecordValue = (value) => normalizeSearchValue(value).replace(/\s+/g, '')

const KEYWORDS = {
  name: ['apellido', 'nombre', 'cliente', 'titular', 'beneficiario'],
  record: ['credito', 'recibo', 'numero', 'nro', 'cuenta', 'legajo', 'codigo', 'cto'],
}

const findColumnKey = (row, type) => {
  const keys = Object.keys(row || {})

  return (
    keys.find((key) =>
      KEYWORDS[type].some((keyword) => normalizeSearchValue(key).includes(keyword)),
    ) || ''
  )
}

const findRecordColumnKey = (row, nameKey) => {
  const explicitRecordKey = findColumnKey(row, 'record')

  if (explicitRecordKey) {
    return explicitRecordKey
  }

  const keys = Object.keys(row || {})
  const nameColumnIndex = keys.indexOf(nameKey)

  if (nameColumnIndex > 0) {
    return keys[nameColumnIndex - 1]
  }

  return ''
}

const parseClientWorkbook = (XLSX, workbook) => {
  const sheetName = workbook.SheetNames[0]

  if (!sheetName) {
    return []
  }

  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    defval: '',
    raw: false,
  })

  if (!rows.length) {
    return []
  }

  const nameKey = findColumnKey(rows[0], 'name') || Object.keys(rows[0])[0] || ''
  const recordKey = findRecordColumnKey(rows[0], nameKey)
  const seenEntries = new Set()

  return rows
    .map((row) => {
      const name = String(row[nameKey] || '').trim()
      const recordNumber = recordKey ? String(row[recordKey] || '').trim() : ''

      return {
        id: `${name}-${recordNumber}`,
        name,
        recordNumber,
      }
    })
    .filter((entry) => {
      if (!entry.name) {
        return false
      }

      const dedupeKey = `${normalizeSearchValue(entry.name)}-${entry.recordNumber}`

      if (seenEntries.has(dedupeKey)) {
        return false
      }

      seenEntries.add(dedupeKey)
      return true
    })
    .sort((left, right) => left.name.localeCompare(right.name, 'es'))
}

const findClientByName = (catalog, value) => {
  const normalizedValue = normalizeSearchValue(value)

  if (!normalizedValue) {
    return null
  }

  return catalog.find((client) => normalizeSearchValue(client.name) === normalizedValue) || null
}

const findClientByRecordNumber = (catalog, value) => {
  const normalizedValue = normalizeRecordValue(value)

  if (!normalizedValue) {
    return null
  }

  return (
    catalog.find((client) => normalizeRecordValue(client.recordNumber) === normalizedValue) || null
  )
}

const FREQUENCY_META = {
  semanal: { label: 'Semanal', installmentLabel: 'Cta. semanal', badge: 'Ruta semanal' },
  quincenal: { label: 'Quincenal', installmentLabel: 'Cuota quincenal', badge: 'Ruta quincenal' },
  mensual: { label: 'Mensual', installmentLabel: 'Cuota mensual', badge: 'Ruta mensual' },
  alternativo: {
    label: 'Mensual alternativo',
    installmentLabel: 'Cuota mensual',
    badge: 'Ruta alternativa',
  },
}

const PLAN_OPTIONS = {
  semanal: [
    { id: '11', label: '11', term: '55 dias', percent: 3, cycleDays: 55, installmentDays: 5 },
    { id: '14.2', label: '14.2', term: '72 dias', percent: 2.16, cycleDays: 72, installmentDays: 5 },
    { id: '22', label: '22', term: '110 dias', percent: 2.64, cycleDays: 110, installmentDays: 5 },
    { id: '32', label: '32', term: '160 dias', percent: 2.88, cycleDays: 160, installmentDays: 5 },
    { id: '40', label: '40', term: '210 dias', percent: 3.12, cycleDays: 210, installmentDays: 5 },
    { id: '48', label: '48', term: '240 dias', percent: 3.375, cycleDays: 240, installmentDays: 5 },
  ],
  quincenal: [
    { id: '6Q', label: '6Q', term: '3 meses', percent: 3, cycleDays: 60, installmentDays: 10 },
    { id: '7Q', label: '7Q', term: '3,5 meses', percent: 2.16, cycleDays: 70, installmentDays: 10 },
    { id: '11Q', label: '11Q', term: '5,5 meses', percent: 2.64, cycleDays: 110, installmentDays: 10 },
    { id: '16Q', label: '16Q', term: '8 meses', percent: 2.88, cycleDays: 160, installmentDays: 10 },
    { id: '20Q', label: '20Q', term: '10 meses', percent: 3.12, cycleDays: 200, installmentDays: 10 },
    { id: '24Q', label: '24Q', term: '12 meses', percent: 3.375, cycleDays: 240, installmentDays: 10 },
  ],
  mensual: [
    { id: '4M', label: '4 meses', term: '4 meses', percent: 2.16, cycleDays: 80, installmentDays: 20 },
    { id: '6M', label: '6 meses', term: '6 meses', percent: 2.64, cycleDays: 120, installmentDays: 20 },
    { id: '8M', label: '8 meses', term: '8 meses', percent: 2.88, cycleDays: 160, installmentDays: 20 },
    { id: '10M', label: '10 meses', term: '10 meses', percent: 3.12, cycleDays: 200, installmentDays: 20 },
    { id: '12M', label: '12 meses', term: '12 meses', percent: 3.375, cycleDays: 240, installmentDays: 20 },
  ],
  alternativo: [
    { id: '3A', label: '3 meses', term: '3 meses', percent: 1.92, cycleDays: 60, installmentDays: 20 },
    { id: '5A', label: '5 meses', term: '5 meses', percent: 2.25, cycleDays: 100, installmentDays: 20 },
    { id: '7A', label: '7 meses', term: '7 meses', percent: 2.75, cycleDays: 140, installmentDays: 20 },
  ],
}

const defaultCredits = [
  { id: 1, concept: '', amount: 0 },
]

const today = new Date().toISOString().slice(0, 10)

function App() {
  const receiptRef = useRef(null)
  const [recordNumber, setRecordNumber] = useState('')
  const [customerName, setCustomerName] = useState('')
  const [receiptDate, setReceiptDate] = useState(today)
  const [usePlanDetails, setUsePlanDetails] = useState(false)
  const [baseAmount, setBaseAmount] = useState('')
  const [frequency, setFrequency] = useState('mensual')
  const [selectedPlanId, setSelectedPlanId] = useState('4M')
  const [credits, setCredits] = useState(defaultCredits)
  const [notes, setNotes] = useState('')
  const [clientCatalog, setClientCatalog] = useState([])
  const [catalogFileName, setCatalogFileName] = useState('')
  const [catalogError, setCatalogError] = useState('')
  const [previewImageUrl, setPreviewImageUrl] = useState('')
  const [isGeneratingImage, setIsGeneratingImage] = useState(false)
  const [isCopyingImage, setIsCopyingImage] = useState(false)
  const [imageActionError, setImageActionError] = useState('')
  const [imageActionSuccess, setImageActionSuccess] = useState('')

  const numericBaseAmount = Number(baseAmount) || 0
  const frequencyMeta = FREQUENCY_META[frequency]
  const plans = PLAN_OPTIONS[frequency]

  const computedPlans = useMemo(
    () =>
      plans.map((plan) => {
        const totalRaw = numericBaseAmount * plan.percent
        const perDayRaw = totalRaw / plan.cycleDays
        const installmentRaw = totalRaw * (plan.installmentDays / plan.cycleDays)

        const total = roundToSingleDecimal(totalRaw)
        const perDay = roundToSingleDecimal(perDayRaw)
        const installment = roundToSingleDecimal(installmentRaw)

        return {
          ...plan,
          total,
          perDay,
          installment,
        }
      }),
    [numericBaseAmount, plans],
  )

  const selectedPlan =
    computedPlans.find((plan) => plan.id === selectedPlanId) || computedPlans[0]

  const totalPaid = useMemo(
    () => credits.reduce((accumulator, credit) => accumulator + (Number(credit.amount) || 0), 0),
    [credits],
  )

  const remainingAmount = useMemo(
    () => Math.max((selectedPlan?.total || 0) - totalPaid, 0),
    [selectedPlan, totalPaid],
  )

  const showExtendedReceipt = usePlanDetails && numericBaseAmount > 0

  const nextCreditId = useMemo(
    () => credits.reduce((maximum, credit) => Math.max(maximum, credit.id), 0) + 1,
    [credits],
  )

  const suggestedClients = useMemo(() => {
    const query = normalizeSearchValue(customerName)

    if (!query || !clientCatalog.length) {
      return []
    }

    return clientCatalog
      .filter((client) => normalizeSearchValue(client.name).includes(query))
      .slice(0, 8)
  }, [clientCatalog, customerName])

  const matchedClientByRecord = useMemo(
    () => findClientByRecordNumber(clientCatalog, recordNumber),
    [clientCatalog, recordNumber],
  )

  const matchedClientByName = useMemo(
    () => findClientByName(clientCatalog, customerName),
    [clientCatalog, customerName],
  )

  const suggestedRecordClients = useMemo(() => {
    const query = normalizeRecordValue(recordNumber)

    if (!query || !clientCatalog.length || matchedClientByRecord) {
      return []
    }

    return clientCatalog
      .filter((client) => normalizeRecordValue(client.recordNumber).includes(query))
      .slice(0, 8)
  }, [clientCatalog, matchedClientByRecord, recordNumber])

  useEffect(() => {
    if (matchedClientByRecord && matchedClientByRecord.name !== customerName) {
      setCustomerName(matchedClientByRecord.name)
    }
  }, [customerName, matchedClientByRecord])

  useEffect(() => {
    if (matchedClientByName?.recordNumber && matchedClientByName.recordNumber !== recordNumber) {
      setRecordNumber(matchedClientByName.recordNumber)
    }
  }, [matchedClientByName, recordNumber])

  const handleFrequencyChange = (event) => {
    const newFrequency = event.target.value
    setFrequency(newFrequency)
    setSelectedPlanId(PLAN_OPTIONS[newFrequency][0].id)
  }

  const handleCreditChange = (id, field, value) => {
    setCredits((currentCredits) =>
      currentCredits.map((credit) =>
        credit.id === id
          ? {
              ...credit,
              [field]: field === 'amount' ? Number(value) || 0 : value,
            }
          : credit,
      ),
    )
  }

  const handleAddCredit = () => {
    setCredits((currentCredits) => [
      ...currentCredits,
      { id: nextCreditId, concept: '', amount: 0 },
    ])
  }

  const handleRemoveCredit = (id) => {
    setCredits((currentCredits) => currentCredits.filter((credit) => credit.id !== id))
  }

  const handleCatalogUpload = async (event) => {
    const file = event.target.files?.[0]

    if (!file) {
      return
    }

    try {
      const buffer = await file.arrayBuffer()
      const XLSX = await import('xlsx')
      const workbook = XLSX.read(buffer, { type: 'array' })
      const parsedClients = parseClientWorkbook(XLSX, workbook)

      if (!parsedClients.length) {
        throw new Error('No se encontraron nombres validos en la primera hoja.')
      }

      setClientCatalog(parsedClients)
      setCatalogFileName(file.name)
      setCatalogError('')

      const clientMatchedByRecord = findClientByRecordNumber(parsedClients, recordNumber)
      const clientMatchedByName = findClientByName(parsedClients, customerName)

      if (clientMatchedByRecord) {
        setCustomerName(clientMatchedByRecord.name)
      } else if (clientMatchedByName?.recordNumber) {
        setRecordNumber(clientMatchedByName.recordNumber)
      }
    } catch (error) {
      setClientCatalog([])
      setCatalogFileName('')
      setCatalogError(error instanceof Error ? error.message : 'No se pudo leer la planilla.')
    }

    event.target.value = ''
  }

  const handleClientPick = (client) => {
    setCustomerName(client.name)

    if (client.recordNumber) {
      setRecordNumber(client.recordNumber)
    }
  }

  const handleRecordNumberChange = (event) => {
    const nextRecordNumber = event.target.value

    setRecordNumber(nextRecordNumber)
  }

  const handleCustomerNameChange = (event) => {
    const nextCustomerName = event.target.value

    setCustomerName(nextCustomerName)
  }

  const handlePrint = () => {
    window.print()
  }

  const generateReceiptImage = async () => {
    if (!receiptRef.current) {
      throw new Error('No se encontró el recibo para generar la imagen.')
    }

    const { toPng } = await import('html-to-image')

    return toPng(receiptRef.current, {
      cacheBust: true,
      backgroundColor: '#fffefa',
      pixelRatio: 2,
    })
  }

  const handlePreviewImage = async () => {
    if (isGeneratingImage) {
      return
    }

    setIsGeneratingImage(true)
    setImageActionError('')
    setImageActionSuccess('')

    try {
      const imageUrl = await generateReceiptImage()
      setPreviewImageUrl(imageUrl)
    } catch {
      setImageActionError('No se pudo generar la vista de la imagen del recibo.')
    } finally {
      setIsGeneratingImage(false)
    }
  }

  const handleCopyImage = async () => {
    if (isCopyingImage) {
      return
    }

    setIsCopyingImage(true)
    setImageActionError('')
    setImageActionSuccess('')

    try {
      const imageUrl = previewImageUrl || (await generateReceiptImage())

      if (!previewImageUrl) {
        setPreviewImageUrl(imageUrl)
      }

      if (!navigator.clipboard || typeof window.ClipboardItem === 'undefined') {
        throw new Error('Clipboard no disponible')
      }

      const response = await fetch(imageUrl)
      const blob = await response.blob()

      await navigator.clipboard.write([
        new window.ClipboardItem({
          [blob.type]: blob,
        }),
      ])

      setImageActionSuccess('Imagen copiada. Ya puedes pegarla donde necesites.')
    } catch {
      setImageActionError('No se pudo copiar la imagen. Puedes usar la vista previa para copiarla manualmente.')
    } finally {
      setIsCopyingImage(false)
    }
  }

  return (
    <div className="app-shell">
      <header className="hero-banner">
        <div className="brand-cluster">
          <img
            className="brand-logo"
            src="/creditos-jardin-logo.svg"
            alt="Logo de Créditos Jardín"
          />
          <div>
            <p className="eyebrow">Gestión de recibos y cuotas</p>
            <h1>Créditos Jardín</h1>
            <p className="hero-copy">
              v1.0-Master, recibo de cuotas y planillas de clientes
            </p>
          </div>
        </div>

        <div className="hero-stats">
          <article>
            <span>Total abonado</span>
            <strong>{formatCurrency(totalPaid)}</strong>
          </article>
          {showExtendedReceipt ? (
            <>
              <article>
                <span>Monto base</span>
                <strong>{formatCurrency(numericBaseAmount)}</strong>
              </article>
              <article>
                <span>Restante</span>
                <strong>{formatCurrency(remainingAmount)}</strong>
              </article>
            </>
          ) : (
            <article>
              <span>Clientes en planilla</span>
              <strong>{clientCatalog.length}</strong>
            </article>
          )}
        </div>
      </header>

      <main className="workspace-grid">
        <section className="panel form-panel">
          <div className="section-heading">
            <div>
              <h2>Datos del recibo</h2>
              <p>
                Se puede trabajar en modo básico o activar los datos del plan para mostrar
                modalidad, cuota, restante y cálculos completos.
              </p>
            </div>
            <span className="pill-highlight">
              {showExtendedReceipt ? frequencyMeta.badge : 'Modo básico'}
            </span>
          </div>

          <div className="form-grid">
            <label className="client-search-field">
              <span>Número de crédito</span>
              <input
                value={recordNumber}
                placeholder={
                  clientCatalog.length
                    ? 'Escribe el número para buscarlo en la planilla'
                    : 'Puedes escribir el número manualmente o cargar una planilla'
                }
                onChange={handleRecordNumberChange}
              />
              {matchedClientByRecord ? (
                <strong className="field-success">Cliente identificado: {matchedClientByRecord.name}</strong>
              ) : suggestedRecordClients.length > 0 ? (
                <div className="client-suggestions">
                  {suggestedRecordClients.map((client) => (
                    <button
                      type="button"
                      key={client.id}
                      className="suggestion-item"
                      onClick={() => handleClientPick(client)}
                    >
                      <span>{client.recordNumber || 'Sin número'}</span>
                      <strong>{client.name}</strong>
                    </button>
                  ))}
                </div>
              ) : recordNumber && clientCatalog.length ? (
                <small className="field-hint">No hay coincidencias con ese número.</small>
              ) : null}
            </label>

            <label className="full-width">
              <span>Planilla de clientes</span>
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleCatalogUpload} />
              <small className="field-hint">
                Sube una planilla Excel y luego busca por nombre. Si la hoja tiene una columna de
                número de crédito o recibo, se completa sola al seleccionar el cliente.
              </small>
              {catalogError ? (
                <strong className="field-error">{catalogError}</strong>
              ) : catalogFileName ? (
                <strong className="field-success">
                  {clientCatalog.length} clientes cargados desde {catalogFileName}
                </strong>
              ) : null}
            </label>

            <label className="full-width client-search-field">
              <span>Apellido y nombre</span>
              <input
                value={customerName}
                placeholder={
                  clientCatalog.length
                    ? 'Escribe parte del nombre para buscarlo en la planilla'
                    : 'Puedes escribir el nombre manualmente o cargar una planilla'
                }
                onChange={handleCustomerNameChange}
              />
              {suggestedClients.length > 0 ? (
                <div className="client-suggestions">
                  {suggestedClients.map((client) => (
                    <button
                      type="button"
                      key={client.id}
                      className="suggestion-item"
                      onClick={() => handleClientPick(client)}
                    >
                      <span>{client.name}</span>
                      {client.recordNumber ? <strong>{client.recordNumber}</strong> : null}
                    </button>
                  ))}
                </div>
              ) : customerName && clientCatalog.length ? (
                <small className="field-hint">No hay coincidencias con ese texto.</small>
              ) : null}
            </label>

            <label>
              <span>Fecha</span>
              <input type="date" value={receiptDate} onChange={(event) => setReceiptDate(event.target.value)} />
            </label>

            <label className="full-width optional-toggle">
              <span>Datos del plan</span>
              <button
                type="button"
                className={`toggle-chip ${usePlanDetails ? 'is-active' : ''}`}
                onClick={() => setUsePlanDetails((currentValue) => !currentValue)}
              >
                {usePlanDetails ? 'Ocultar datos avanzados' : 'Agregar modalidad, plan y monto base'}
              </button>
              <small className="field-hint">
                  Si no lo activas, el recibo se muestra en formato básico como tu modelo.
              </small>
            </label>

            {usePlanDetails ? (
              <>
                <label>
                  <span>Monto base</span>
                  <input
                    type="number"
                    min="0"
                    step="0.1"
                    value={baseAmount}
                    placeholder="30000"
                    onChange={(event) => setBaseAmount(event.target.value)}
                  />
                </label>

                <label>
                  <span>Modalidad</span>
                  <select value={frequency} onChange={handleFrequencyChange}>
                    {Object.entries(FREQUENCY_META).map(([key, meta]) => (
                      <option key={key} value={key}>
                        {meta.label}
                      </option>
                    ))}
                  </select>
                </label>

                <label>
                  <span>Plan</span>
                  <select value={selectedPlanId} onChange={(event) => setSelectedPlanId(event.target.value)}>
                    {computedPlans.map((plan) => (
                      <option key={plan.id} value={plan.id}>
                        {plan.label}
                      </option>
                    ))}
                  </select>
                </label>
              </>
            ) : null}

            <label className="full-width">
              <span>Observaciones</span>
              <textarea
                rows="4"
                value={notes}
                onChange={(event) => setNotes(event.target.value)}
                placeholder="Detalle de pagos parciales, aclaraciones para el cliente o notas de cobranzas."
              />
            </label>
          </div>

          <div className="section-heading compact">
            <h2>Créditos abonados</h2>
            <button type="button" className="secondary-button" onClick={handleAddCredit}>
              Agregar fila
            </button>
          </div>

          <div className="credits-editor">
            {credits.map((credit) => (
              <div className="credit-row" key={credit.id}>
                <input
                  aria-label={`Concepto ${credit.id}`}
                  placeholder="Concepto o Nro. credito"
                  value={credit.concept}
                  onChange={(event) => handleCreditChange(credit.id, 'concept', event.target.value)}
                />
                <input
                  aria-label={`Monto ${credit.id}`}
                  type="number"
                  min="0"
                  step="0.1"
                  placeholder="Monto"
                  value={credit.amount}
                  onChange={(event) => handleCreditChange(credit.id, 'amount', event.target.value)}
                />
                <button type="button" className="ghost-button" onClick={() => handleRemoveCredit(credit.id)}>
                  Quitar
                </button>
              </div>
            ))}
          </div>

          {usePlanDetails ? (
            <>
              <section className="plan-card">
                <div className="section-heading compact">
                  <h2>Resumen del plan</h2>
                  <span className="plan-pill">{selectedPlan.label}</span>
                </div>
                <div className="plan-metrics">
                  <article>
                    <span>Plazo</span>
                    <strong>{selectedPlan.term}</strong>
                  </article>
                  <article>
                    <span>{frequencyMeta.installmentLabel}</span>
                    <strong>{formatCurrency(selectedPlan.installment)}</strong>
                  </article>
                  <article>
                    <span>Total financiado</span>
                    <strong>{formatCurrency(selectedPlan.total)}</strong>
                  </article>
                  <article>
                    <span>Valor por dia</span>
                    <strong>{formatCurrency(selectedPlan.perDay)}</strong>
                  </article>
                  <article>
                    <span>% aplicado</span>
                    <strong>{formatPercent(selectedPlan.percent)}</strong>
                  </article>
                  {showExtendedReceipt ? (
                    <article>
                      <span>Restante</span>
                      <strong>{formatCurrency(remainingAmount)}</strong>
                    </article>
                  ) : null}
                </div>
              </section>

              <section className="reference-card">
                <div className="section-heading compact">
                  <h2>Tabla de planes calculada</h2>
                  <span className="table-caption">Base {formatCurrency(numericBaseAmount)}</span>
                </div>
                <div className="plan-table-wrapper">
                  <table className="plan-reference-table">
                    <thead>
                      <tr>
                        <th>Plan</th>
                        <th>Plazo</th>
                        <th>Cuota</th>
                        <th>Total</th>
                        <th>Dia</th>
                        <th>%</th>
                      </tr>
                    </thead>
                    <tbody>
                      {computedPlans.map((plan) => (
                        <tr key={plan.id} className={plan.id === selectedPlan.id ? 'is-selected' : ''}>
                          <td>{plan.label}</td>
                          <td>{plan.term}</td>
                          <td>{formatCurrency(plan.installment)}</td>
                          <td>{formatCurrency(plan.total)}</td>
                          <td>{formatCurrency(plan.perDay)}</td>
                          <td>{formatPercent(plan.percent)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            </>
          ) : null}
        </section>

        <section className="panel report-panel">
          <div className="report-actions no-print">
            <button type="button" className="primary-button" onClick={handlePrint}>
              Imprimir recibo
            </button>
            <button
              type="button"
              className="secondary-button"
              onClick={handlePreviewImage}
              disabled={isGeneratingImage}
            >
              {isGeneratingImage ? 'Generando imagen...' : 'Ver imagen'}
            </button>
            {previewImageUrl ? (
              <button
                type="button"
                className="secondary-button"
                onClick={handleCopyImage}
                disabled={isCopyingImage}
              >
                {isCopyingImage ? 'Copiando...' : 'Copiar imagen'}
              </button>
            ) : null}
          </div>
          {imageActionError ? <p className="report-error no-print">{imageActionError}</p> : null}
          {imageActionSuccess ? <p className="report-success no-print">{imageActionSuccess}</p> : null}
          {previewImageUrl ? (
            <section className="image-preview-card no-print">
              <div className="image-preview-header">
                <div>
                  <h3>Vista previa de la imagen</h3>
                  <p>Puedes copiarla con el botón o desde esta vista previa.</p>
                </div>
                <button
                  type="button"
                  className="ghost-button"
                  onClick={() => {
                    setPreviewImageUrl('')
                    setImageActionError('')
                    setImageActionSuccess('')
                  }}
                >
                  Cerrar vista
                </button>
              </div>
              <img
                className="image-preview-image"
                src={previewImageUrl}
                alt="Vista previa del recibo como imagen"
              />
            </section>
          ) : null}

          <article className="receipt-sheet" ref={receiptRef}>
            <header className="receipt-topbar">
              <div className="receipt-branding">
                <img
                  className="receipt-logo"
                  src="/creditos-jardin-logo.svg"
                  alt="Logo de Créditos Jardín"
                />
                <div>
                  <span className="receipt-label">Recibo para el cliente</span>
                  <h2>Créditos Jardín</h2>
                </div>
              </div>
              <div className="receipt-number-boxes">
                <div>
                  <span>Número de crédito</span>
                  <strong>{recordNumber || '---'}</strong>
                </div>
                <div>
                  <span>Fecha</span>
                  <strong>{formatDate(receiptDate)}</strong>
                </div>
              </div>
            </header>

            <div className="receipt-banner">
              {notes.trim() || 'Recibo generado para registrar pagos de cuotas o parciales.'}
            </div>

            <section className={`receipt-client-grid ${showExtendedReceipt ? 'is-extended' : 'is-basic'}`}>
              <div>
                <span>Apellido y nombre</span>
                <strong>{customerName || 'Sin completar'}</strong>
              </div>
              {showExtendedReceipt ? (
                <>
                  <div>
                    <span>Modalidad</span>
                    <strong>{frequencyMeta.label}</strong>
                  </div>
                  <div>
                    <span>Plan elegido</span>
                    <strong>{selectedPlan.label}</strong>
                  </div>
                  <div>
                    <span>Monto base</span>
                    <strong>{formatCurrency(numericBaseAmount)}</strong>
                  </div>
                  <div>
                    <span>{frequencyMeta.installmentLabel}</span>
                    <strong>{formatCurrency(selectedPlan.installment)}</strong>
                  </div>
                </>
              ) : null}
            </section>

            <table className="receipt-table">
              <thead>
                <tr>
                  <th>Concepto / Nro. credito</th>
                  <th>Unidad</th>
                  <th>Subtotal</th>
                </tr>
              </thead>
              <tbody>
                {credits.map((credit) => (
                  <tr key={credit.id}>
                    <td>{credit.concept || 'Sin detalle'}</td>
                    <td>{formatCurrency(credit.amount)}</td>
                    <td>{formatCurrency(credit.amount)}</td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr>
                  <td colSpan="2">Total abonado</td>
                  <td>{formatCurrency(totalPaid)}</td>
                </tr>
              </tfoot>
            </table>

            {showExtendedReceipt ? (
              <section className="receipt-summary-grid">
                <article>
                  <span>{frequencyMeta.installmentLabel}</span>
                  <strong>{formatCurrency(selectedPlan.installment)}</strong>
                </article>
                <article>
                  <span>Total financiado</span>
                  <strong>{formatCurrency(selectedPlan.total)}</strong>
                </article>
                <article>
                  <span>Valor por dia</span>
                  <strong>{formatCurrency(selectedPlan.perDay)}</strong>
                </article>
                <article>
                  <span>% del plan</span>
                  <strong>{formatPercent(selectedPlan.percent)}</strong>
                </article>
              </section>
            ) : null}

            <footer className={`receipt-footer ${showExtendedReceipt ? 'is-extended' : 'is-basic'}`}>
              <div>
                <span>Total abonado</span>
                <strong>{formatCurrency(totalPaid)} pesos argentinos</strong>
              </div>
              {showExtendedReceipt ? (
                <>
                  <div>
                    <span>Total del plan</span>
                    <strong>{formatCurrency(selectedPlan.total)}</strong>
                  </div>
                  <div>
                    <span>Restante</span>
                    <strong>{formatCurrency(remainingAmount)}</strong>
                  </div>
                </>
              ) : null}
            </footer>
          </article>
        </section>
      </main>
    </div>
  )
}

export default App
