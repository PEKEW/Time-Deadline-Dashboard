import { useState, useEffect, useRef } from 'react'
import TimeLine from './TimeLine'
import * as XLSX from 'xlsx-js-style'

function App() {
  const [programs, setPrograms] = useState([])
  const [today] = useState(new Date())
  const [currentAoETime, setCurrentAoETime] = useState('')
  const [errorMessage, setErrorMessage] = useState('')
  const fileInputRef = useRef(null)

  const getCurrentAoETime = () => {
    const now = new Date()
    const utc = now.getTime() + (now.getTimezoneOffset() * 60000)
    const aoeTime = new Date(utc - (12 * 60 * 60 * 1000))

    const year = aoeTime.getFullYear()
    const month = String(aoeTime.getMonth() + 1).padStart(2, '0')
    const day = String(aoeTime.getDate()).padStart(2, '0')
    const hours = String(aoeTime.getHours()).padStart(2, '0')
    const minutes = String(aoeTime.getMinutes()).padStart(2, '0')
    const seconds = String(aoeTime.getSeconds()).padStart(2, '0')

    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`
  }

  useEffect(() => {
    const updateTime = () => {
      setCurrentAoETime(getCurrentAoETime())
    }

    updateTime()
    const timer = setInterval(updateTime, 1000)

    return () => clearInterval(timer)
  }, [])

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

      // Handle conference (DDL) node
      let conferenceNode = null
      if (p.conference) {
        if (!p.conference.name || typeof p.conference.name !== 'string') {
          throw new Error(`Program "${p.name}" conference is missing a valid name field`)
        }
        if (!p.conference.date) {
          throw new Error(`Program "${p.name}" conference is missing the date field`)
        }

        const conferenceDate = new Date(p.conference.date)
        if (isNaN(conferenceDate.getTime())) {
          throw new Error(`Program "${p.name}" conference has invalid date format: "${p.conference.date}"`)
        }

        conferenceNode = {
          id: `${p.id}-conference`,
          name: p.conference.name,
          date: conferenceDate
        }
      } else {
        // If no conference field, automatically compute DDL
        if (timePoints.length > 0) {
          const lastDate = new Date(timePoints[timePoints.length - 1].date)
          const autoDDL = new Date(lastDate)
          autoDDL.setDate(autoDDL.getDate() + ddlGapDays)

          conferenceNode = {
            id: `${p.id}-conference`,
            name: 'Conference',
            date: autoDDL
          }
        }
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
    } catch (error) {
      console.error('Failed to load config file:', error)
      setErrorMessage(`Load failed: ${error.message}`)
    }
  }

  // Import data from Excel
  const handleExcelUpload = (event) => {
    const file = event.target.files[0]
    if (!file) return

    setErrorMessage('')
    const reader = new FileReader()

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { type: 'array' })

  // Read the first worksheet
  const sheetName = workbook.SheetNames[0]
  const worksheet = workbook.Sheets[sheetName]

  // Convert worksheet to JSON
  const jsonData = XLSX.utils.sheet_to_json(worksheet)

        if (!jsonData || jsonData.length === 0) {
          throw new Error('Excel file is empty or has invalid format')
        }

        // Validate required columns
        const firstRow = jsonData[0]
        if (!firstRow.DATE || !firstRow.EVENT) {
          throw new Error('Excel file must contain DATE and EVENT columns')
        }

  // Group rows by program (process in order)
  const programsMap = new Map()
        let currentProgramName = null

        jsonData.forEach((row, index) => {
          if (!row.DATE || !row.EVENT) {
            console.warn(`Skipping row ${index + 2}: missing required fields`)
            return
          }

          // Parse date
          const date = new Date(row.DATE)
          if (isNaN(date.getTime())) {
            throw new Error(`Invalid date format on row ${index + 2}: "${row.DATE}"`)
          }

          // Check if this is a Conference row
          if (row.EVENT.trim() === 'Conference') {
            // This is a conference row and belongs to the current program
            if (!currentProgramName) {
              throw new Error(`Row ${index + 2} is Conference but has no corresponding Program`)
            }

            programsMap.get(currentProgramName).conference = {
              id: `${currentProgramName.toLowerCase().replace(/\s+/g, '-')}-conference`,
              name: 'Conference',
              date: date
            }
          } else {
            // Parse EVENT field: format is "ProgramName - EventName"
            const eventParts = row.EVENT.split(' - ')
            if (eventParts.length < 2) {
              throw new Error(`Invalid EVENT format on row ${index + 2}; expected "ProgramName - EventName" or "Conference"`)
            }

            const programName = eventParts[0].trim()
            const eventName = eventParts.slice(1).join(' - ').trim()

            // Update current program
            currentProgramName = programName

            // Add to the corresponding program
            if (!programsMap.has(programName)) {
              programsMap.set(programName, { events: [], conference: null })
            }

            // This is a regular event
            programsMap.get(programName).events.push({
              id: `${programName.toLowerCase().replace(/\s+/g, '-')}-${index}`,
              name: eventName,
              date: date
            })
          }
        })

        // Convert to the format required by the app
        const loadedPrograms = Array.from(programsMap.entries()).map(([programName, data], index) => {
          // Sort timePoints by date
          data.events.sort((a, b) => a.date - b.date)

          let conferenceNode = data.conference

          // If there is no conference node, auto-compute the DDL
          if (!conferenceNode && data.events.length > 0) {
            const lastDate = new Date(data.events[data.events.length - 1].date)
            const autoDDL = new Date(lastDate)
            autoDDL.setDate(autoDDL.getDate() + 30)

            conferenceNode = {
              id: `${programName.toLowerCase().replace(/\s+/g, '-')}-conference`,
              name: 'Conference',
              date: autoDDL
            }
          }

          return {
            id: programName.toLowerCase().replace(/\s+/g, '-'),
            name: programName,
            timePoints: data.events,
            conference: conferenceNode,
            ddl: conferenceNode ? conferenceNode.date : null
          }
        })

        setPrograms(loadedPrograms)
        } catch (error) {
        console.error('Excel parsing failed:', error)
        setErrorMessage(`Excel parsing failed: ${error.message}`)
      }
    }

    reader.onerror = () => {
      setErrorMessage('File read failed, please try again')
    }

    reader.readAsArrayBuffer(file)
  }

  // Trigger file input
  const triggerFileInput = () => {
    fileInputRef.current?.click()
  }

  // Initial load
  useEffect(() => {
    loadConfig()
  }, [])

  // Update the date of a time point
  const updateTimePointDate = (programId, timePointId, newDate) => {
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
      <div className="title-container">
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

      <div
        className="header-actions"
        style={{
          display: 'flex',
          gap: '8px',
          flexWrap: 'nowrap',
          alignItems: 'center'
        }}
      >
        <button className="refresh-btn" onClick={loadConfig}>
          Restore to initial state
        </button>
        <button className="refresh-btn" onClick={triggerFileInput}>
          Import Excel
        </button>
        <button className="export-btn" onClick={exportToExcel}>
          Export Excel
        </button>
      </div>

      {/* Hidden file input */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={handleExcelUpload}
        style={{ display: 'none' }}
      />

      {/* Error message display */}
      {errorMessage && (
        <div className="error-message">
          {errorMessage}
        </div>
      )}

      <div className="today-info">
        <strong>Current AoE Time:</strong>{' '}
        <span className="date">{currentAoETime}</span>
      </div>

      <div className="timeline-container">
        {programsWithColors.map((program, index) => (
          <TimeLine
            key={program.id}
            program={program}
            today={today}
            color={program.color} // Pass the color to TimeLine
            isLastProgram={index === programsWithColors.length - 1} // Only show conference for last program
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
