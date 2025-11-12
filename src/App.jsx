import { useState, useEffect, useRef } from 'react'
import TimeLine from './TimeLine'
import * as XLSX from 'xlsx-js-style'

function App() {
  const [programs, setPrograms] = useState([])
  const [today, setToday] = useState(new Date())
  const [errorMessage, setErrorMessage] = useState('')
  const [isLocked, setIsLocked] = useState(false)


  const processConfig = (config) => {
    if (!config || typeof config !== 'object') {
      throw new Error('Invalid config file format: not a valid JSON object')
    }

    if (!Array.isArray(config.programs)) {
      throw new Error("Invalid config file format: missing 'programs' array")
    }

    if (config.programs.length === 0) {
      throw new Error("Invalid config file format: 'programs' array is empty")
    }

    // DDL gap days (default: 30 days)
    const ddlGapDays = config.ddlGapDays || 30

    const loadedPrograms = config.programs.map((p, index) => {
      // Validate program structure
      if (!p.id || typeof p.id !== 'string') {
        throw new Error(`Program ${index + 1} is missing a valid id field`)
      }
      if (!p.name || typeof p.name !== 'string') {
        throw new Error(`Program "${p.id}" is missing a valid name field`)
      }
      if (!Array.isArray(p.timePoints)) {
        throw new Error(`Program "${p.name}" is missing the timePoints array`)
      }
      if (p.timePoints.length === 0) {
        throw new Error(`Program "${p.name}" has an empty timePoints array`)
      }

      const timePoints = p.timePoints.map((tp, tpIndex) => {
        // Validate timePoint structure
        if (!tp.id || typeof tp.id !== 'string') {
          throw new Error(`Program "${p.name}" TimePoint ${tpIndex + 1} is missing a valid id field`)
        }
        if (!tp.name || typeof tp.name !== 'string') {
          throw new Error(`Program "${p.name}" TimePoint "${tp.id}" is missing a valid name field`)
        }
        if (!tp.date) {
          throw new Error(`Program "${p.name}" TimePoint "${tp.name}" is missing the date field`)
        }

        // Validate date format
        const date = new Date(tp.date)
        if (isNaN(date.getTime())) {
          throw new Error(`Program "${p.name}" TimePoint "${tp.name}" has invalid date format: "${tp.date}"`)
        }

        return {
          ...tp,
          date: date
        }
      })

      // Handle conference (DDL) node - Fixed to 2026-10-05
      const fixedConferenceDate = new Date('2026-10-05')
      const conferenceNode = {
        id: `${p.id}-conference`,
        name: 'Conference',
        date: fixedConferenceDate
      }

      return {
        ...p,
        ddl: conferenceNode ? conferenceNode.date : null,
        conference: conferenceNode,
        timePoints
      }
    })

    return loadedPrograms
  }

  // Load data from the default config file
  const loadConfig = async () => {
    try {
      setErrorMessage('')
      const response = await fetch('/timeline-config.json')

      if (!response.ok) {
        throw new Error(`Failed to load config file: HTTP ${response.status}`)
      }

      const config = await response.json()
      const loadedPrograms = processConfig(config)
      setPrograms(loadedPrograms)

      // Set today to 15 days before the first event of the first program
      if (loadedPrograms.length > 0 && loadedPrograms[0].timePoints.length > 0) {
        const firstEventDate = new Date(loadedPrograms[0].timePoints[0].date)
        const newToday = new Date(firstEventDate)
        newToday.setDate(newToday.getDate() - 15)
        setToday(newToday)
      }
    } catch (error) {
      console.error('Failed to load config file:', error)
      setErrorMessage(`Load failed: ${error.message}`)
    }
  }


  // Initial load
  useEffect(() => {
    loadConfig()
  }, [])

  // Update the date of a time point
  const updateTimePointDate = (programId, timePointId, newDate) => {
    if (isLocked) return // Don't allow changes when locked

    setPrograms(programs.map(p => {
      if (p.id === programId) {
        return {
          ...p,
          timePoints: p.timePoints.map(tp =>
            tp.id === timePointId ? { ...tp, date: newDate } : tp
          )
        }
      }
      return p
    }))
  }

  // 16-color text palette (dark colors, suitable for white background)
  const COLOR_POOL = [
    'DC143C', // crimson
    '1E90FF', // dodger blue
    '228B22', // forest green
    'FF8C00', // dark orange
    '9370De', // medium purple
    'FF1493', // deep pink
    '00CED1', // dark turquoise
    'DAA520', // goldenrod
    'C71585', // medium violet red
    '32CD32', // lime green
    'BA55D3', // medium orchid
    'FF6347', // tomato
    '4169E1', // royal blue
    '9ACD32', // yellow green
    'FF69B4', // hot pink
    '4682B4'  // steel blue
  ]

  // Export to Excel (two-column format, colored, grouped by program)
  const exportToExcel = () => {
  // Prepare dataRows (in program order, no extra sorting)
    const dataRows = []

    programs.forEach((p, programIndex) => {
      const color = COLOR_POOL[programIndex % COLOR_POOL.length]

  // Add regular events
      p.timePoints.forEach(tp => {
        dataRows.push({
          date: tp.date,
          event: `${p.name} - ${tp.name}`,
          color: color
        })
      })
    })

    // Add single conference node at the end (use the last program's conference)
    const lastProgram = programs[programs.length - 1]
    if (lastProgram && lastProgram.conference) {
      dataRows.push({
        date: lastProgram.conference.date,
        event: 'Conference',
        color: '000000'  // Use black color for Conference
      })
    }

    // Create worksheet data (including header row)
    const wsData = [
      ['DATE', 'EVENT'], // header row
      ...dataRows.map(row => [
        row.date.toLocaleDateString('en-US', {
          year: 'numeric',
          month: 'long',
          day: 'numeric'
        }),
        row.event
      ])
    ]

  // Create worksheet
  const ws = XLSX.utils.aoa_to_sheet(wsData)

    // Set column widths
    ws['!cols'] = [
      { wch: 20 }, // DATE column width
      { wch: 50 }  // EVENT column width
    ]

  // Set header row style (dark background, white text, bold)
    const headerStyle = {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '2C3E50' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: '000000' } },
        bottom: { style: 'thin', color: { rgb: '000000' } },
        left: { style: 'thin', color: { rgb: '000000' } },
        right: { style: 'thin', color: { rgb: '000000' } }
      }
    }

    ws['A1'].s = headerStyle
    ws['B1'].s = headerStyle

  // Apply style to each data row
    dataRows.forEach((row, index) => {
      const rowNum = index + 2 // +2 because Excel rows start at 1 and there's a header row

      const cellStyle = {
        font: { color: { rgb: row.color } },
        alignment: { horizontal: 'left', vertical: 'center' },
        border: {
          top: { style: 'thin', color: { rgb: 'CCCCCC' } },
          bottom: { style: 'thin', color: { rgb: 'CCCCCC' } },
          left: { style: 'thin', color: { rgb: 'CCCCCC' } },
          right: { style: 'thin', color: { rgb: 'CCCCCC' } }
        }
      }

      const dateCell = `A${rowNum}`
      const eventCell = `B${rowNum}`

      if (ws[dateCell]) ws[dateCell].s = cellStyle
      if (ws[eventCell]) ws[eventCell].s = cellStyle
    })

  // Create workbook and export
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Timeline')
    XLSX.writeFile(wb, 'timeline-export.xlsx')
  }

  // Assign a color to each program based on its index
  const programsWithColors = programs.map((program, index) => ({
    ...program,
    color: COLOR_POOL[index % COLOR_POOL.length],
  }));

  return (
    <div className="app">
      <div className="header-container">
        <div className="title-section">
          <h1 className="main-title">Time Deadlines(AoE)</h1>
          <div className="info-icon-wrapper">
            <div className="info-icon">i</div>
            <div className="info-tooltip">
              <div className="tooltip-item">
                <strong>Dragg:</strong> Drag red nodes to adjust the date
              </div>
              <div className="tooltip-item">
                <strong>Double-click:</strong> Double-click red nodes to enter a specific date
              </div>
            </div>
          </div>
        </div>
        <div className="header-actions">
          <button className="refresh-btn" onClick={loadConfig}>
            Restore
          </button>
          <button className="export-btn" onClick={exportToExcel}>
            Export
          </button>
          {isLocked ? (
            <button className="unlock-btn" onClick={() => setIsLocked(false)}>
              Unlock
            </button>
          ) : (
            <button className="lock-btn" onClick={() => setIsLocked(true)}>
              Lock
            </button>
          )}
        </div>
      </div>

      {/* Error message display */}
      {errorMessage && (
        <div className="error-message">
          {errorMessage}
        </div>
      )}

      <div className="timeline-container">
        {programsWithColors.map((program, index) => (
          <TimeLine
            key={program.id}
            program={program}
            today={today}
            color={program.color} // Pass the color to TimeLine
            isLocked={isLocked} // Pass the lock state to TimeLine
            isFirst={index === 0} // Pass whether this is the first program
            onTimePointChange={(timePointId, newDate) =>
              updateTimePointDate(program.id, timePointId, newDate)
            }
          />
        ))}
      </div>
    </div>
  )
}

export default App
