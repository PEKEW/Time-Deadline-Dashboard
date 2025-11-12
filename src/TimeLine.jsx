import { useState, useRef, useEffect } from 'react'

// Single time point component
function TimePoint({ timePoint, position, today, ddl, canDrag, onDateChange, pixelsPerDay, prevDate, nextDate, offsetLevel, isConference = false, color }) {
  const [isDragging, setIsDragging] = useState(false)
  const [isEditing, setIsEditing] = useState(false)
  const [editValue, setEditValue] = useState('')
  const [countdown, setCountdown] = useState('')
  const dragStartX = useRef(0)
  const dragStartDate = useRef(null)

  // Get current AoE time (UTC-12)
  const getCurrentAoETime = () => {
    const now = new Date()
    const utc = now.getTime() + (now.getTimezoneOffset() * 60000)
    const aoeTime = new Date(utc - (12 * 60 * 60 * 1000))
    return aoeTime
  }

  // Calculate countdown to the target date's AoE 00:00:00
  const calculateCountdown = () => {
    const currentAoE = getCurrentAoETime()

  // Target date's AoE 00:00:00
  const targetDate = new Date(timePoint.date)
  // Convert to midnight in AoE timezone
    const targetAoE = new Date(Date.UTC(
      targetDate.getFullYear(),
      targetDate.getMonth(),
      targetDate.getDate(),
      12, 0, 0, 0 // UTC+12 = AoE 00:00:00
    ))

    const diff = targetAoE - currentAoE

    if (diff <= 0) {
      const absDiff = Math.abs(diff)
      const days = Math.floor(absDiff / (1000 * 60 * 60 * 24))
      const hours = Math.floor((absDiff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60))
      const minutes = Math.floor((absDiff % (1000 * 60 * 60)) / (1000 * 60))
      const seconds = Math.floor((absDiff % (1000 * 60)) / 1000)
      return `Passed ${days} days ${hours}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`
    } else {
      const days = Math.floor(diff / (1000 * 60 * 60 * 24))
      const hours = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60))
      const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60))
      const seconds = Math.floor((diff % (1000 * 60)) / 1000)

      if (days > 0) {
        return `Remaining ${days} days ${hours}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`
      } else {
        return `Remaining ${hours}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`
      }
    }
  }

  // Update countdown every second
  useEffect(() => {
    const updateCountdown = () => {
      setCountdown(calculateCountdown())
    }

    updateCountdown()
    const timer = setInterval(updateCountdown, 1000)

    return () => clearInterval(timer)
  }, [timePoint.date])

  // Calculate days from today (used for static day labels)
  const getDaysFromToday = (date) => {
    const todayCopy = new Date(today)
    todayCopy.setHours(0, 0, 0, 0)
    const targetDate = new Date(date)
    targetDate.setHours(0, 0, 0, 0)
    return Math.ceil((targetDate - todayCopy) / (1000 * 60 * 60 * 24))
  }

  // Clamp date between today and DDL, and avoid crossing adjacent nodes
  const clampDate = (date) => {
    const todayTime = new Date(today)
    todayTime.setHours(0, 0, 0, 0)
    const ddlTime = new Date(ddl)
    ddlTime.setHours(0, 0, 0, 0)

    const dateTime = new Date(date)
    dateTime.setHours(0, 0, 0, 0)

  // Basic constraint: cannot go beyond today and DDL
    let clampedDate = dateTime
    if (clampedDate < todayTime) clampedDate = todayTime
    if (clampedDate > ddlTime) clampedDate = ddlTime

  // Order constraint: cannot cross the previous node
    if (prevDate) {
      const prevTime = new Date(prevDate)
      prevTime.setHours(0, 0, 0, 0)
  // Current node cannot be earlier than previous node (0-day buffer, same day allowed)
      if (clampedDate < prevTime) clampedDate = prevTime
    }

  // Order constraint: cannot cross the next node
    if (nextDate) {
      const nextTime = new Date(nextDate)
      nextTime.setHours(0, 0, 0, 0)
  // Current node cannot be later than next node (0-day buffer, same day allowed)
      if (clampedDate > nextTime) clampedDate = nextTime
    }

    return clampedDate
  }

  // Handle drag start
  const handleMouseDown = (e) => {
    if (!canDrag) return
    e.preventDefault()
    setIsDragging(true)
    dragStartX.current = e.clientX
    dragStartDate.current = new Date(timePoint.date)
  }

  // Handle dragging
  const handleMouseMove = (e) => {
    if (!isDragging || !canDrag) return

    const deltaX = e.clientX - dragStartX.current
    const daysChange = Math.round(deltaX / pixelsPerDay) // Calculate day change based on pixelsPerDay

    const newDate = new Date(dragStartDate.current)
    newDate.setDate(newDate.getDate() + daysChange)

  // Clamp between today and DDL
    const clampedDate = clampDate(newDate)
    onDateChange(clampedDate)
  }

  // Handle drag end
  const handleMouseUp = () => {
    setIsDragging(false)
  }

  // Handle double-click to edit
  const handleDoubleClick = () => {
    setIsEditing(true)
    setEditValue(timePoint.date.toISOString().split('T')[0])
  }

  // Handle date input
  const handleDateSubmit = () => {
    if (editValue) {
  const newDate = new Date(editValue)
  // Clamp between today and DDL
      const clampedDate = clampDate(newDate)
      onDateChange(clampedDate)
    }
    setIsEditing(false)
  }

  // Listen to global mouse events
  useEffect(() => {
    if (isDragging) {
      window.addEventListener('mousemove', handleMouseMove)
      window.addEventListener('mouseup', handleMouseUp)
      return () => {
        window.removeEventListener('mousemove', handleMouseMove)
        window.removeEventListener('mouseup', handleMouseUp)
      }
    }
  }, [isDragging, pixelsPerDay])

  const isInteracting = isDragging || isEditing

  // Determine classes for position line and card based on offsetLevel
  let lineClass = 'position-line normal'
  let infoClass = 'date-info top'

  if (offsetLevel === 1) {
    lineClass = 'position-line offset-high'
    infoClass = 'date-info top offset-high'
  } else if (offsetLevel === 2) {
    lineClass = 'position-line offset-higher'
    infoClass = 'date-info top offset-higher'
  } else if (offsetLevel === 3) {
    lineClass = 'position-line offset-highest'
    infoClass = 'date-info top offset-highest'
  }

  return (
    <div
      className="timeline-point-wrapper"
      style={{
        left: `${position}px`,
        zIndex: isInteracting ? 100 : 10
      }}
    >
      <div
        className={`timeline-point ${isDragging ? 'dragging' : ''} ${!canDrag ? 'disabled' : ''}`}
        onMouseDown={handleMouseDown}
        onDoubleClick={handleDoubleClick}
        style={{ cursor: canDrag ? 'grab' : 'not-allowed' }}
      >
        <div
          className={`point-dot ${isConference ? 'conference-node' : ''}`}
          style={{ 
            background: `#${color}`,
            '--pulse-color': `#${color}`,
            '--pulse-color-alpha': `#${color}66`,
            '--pulse-color-fade': `#${color}00`
          }}
        ></div>

        {/* Vertical positioning line */}
        <div className={lineClass} style={{ color: `#${color}` }}></div>

        {/* Display node info */}
        <div className={infoClass}>
          <div className="point-name" style={{ color: `#${color}` }}>{timePoint.name}</div>
          {isEditing ? (
            <input
              type="date"
              value={editValue}
              onChange={(e) => setEditValue(e.target.value)}
              onBlur={handleDateSubmit}
              onKeyPress={(e) => e.key === 'Enter' && handleDateSubmit()}
              autoFocus
              className="date-input"
            />
          ) : (
            <div className="days-info">
              <div style={{ color: `#${color}` }}>{timePoint.date.toLocaleDateString('en-US')}</div>
              <div className="days-count" style={{ color: `#${color}` }}>
                {countdown}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

// Compute number of days between two dates
function getDaysBetween(date1, date2) {
  const d1 = new Date(date1)
  const d2 = new Date(date2)
  d1.setHours(0, 0, 0, 0)
  d2.setHours(0, 0, 0, 0)
  return Math.ceil(Math.abs((d2 - d1) / (1000 * 60 * 60 * 24)))
}

// Timeline main component
function TimeLine({ program, today, onTimePointChange, isLastProgram = false }) {
  const timelineRef = useRef(null)
  const [containerWidth, setContainerWidth] = useState(1000)
  const START_OFFSET = 100 // start offset (space for Today marker)
  const END_OFFSET = 100 // end offset (space for DDL marker)
  const PADDING = 30 // extra padding

  // Watch container width changes
  useEffect(() => {
    const updateWidth = () => {
      if (timelineRef.current) {
        setContainerWidth(timelineRef.current.offsetWidth)
      }
    }

    updateWidth()
    window.addEventListener('resize', updateWidth)
    return () => window.removeEventListener('resize', updateWidth)
  }, [])

  // Calculate timeline layout (responsive)
  const calculateLayout = () => {
    if (!program.ddl || !program.timePoints || program.timePoints.length === 0) {
      return { positions: [], width: containerWidth, todayPos: START_OFFSET, ddlPos: containerWidth - END_OFFSET, pixelsPerDay: 1 }
    }

    const todayTime = new Date(today)
    todayTime.setHours(0, 0, 0, 0)
    const ddlTime = new Date(program.ddl)
    ddlTime.setHours(0, 0, 0, 0)

  // Total days on timeline (from today to DDL)
    const totalDays = Math.max(1, (ddlTime - todayTime) / (1000 * 60 * 60 * 24))

  // availableWidth = containerWidth - offsets - extra padding
    const availableWidth = containerWidth - START_OFFSET - END_OFFSET - PADDING

  // Dynamically compute pixels per day
  const PIXELS_PER_DAY = Math.max(2, availableWidth / totalDays) // minimum 2px/day

  // Timeline total width
    const timelineWidth = totalDays * PIXELS_PER_DAY
    const totalWidth = timelineWidth + START_OFFSET + END_OFFSET

  // Positions for Today and DDL
    const todayPos = START_OFFSET
    const ddlPos = START_OFFSET + timelineWidth

  // Compute position for each time point (relative to today)
    const positions = program.timePoints.map(tp => {
      const tpTime = new Date(tp.date)
      tpTime.setHours(0, 0, 0, 0)
      const daysSinceToday = (tpTime - todayTime) / (1000 * 60 * 60 * 24)
      return START_OFFSET + daysSinceToday * PIXELS_PER_DAY
    })

    return { positions, width: totalWidth, todayPos, ddlPos, timelineWidth, pixelsPerDay: PIXELS_PER_DAY }
  }

  const { positions, width, todayPos, ddlPos, timelineWidth, pixelsPerDay } = calculateLayout()

  return (
    <div className="timeline-row">
      <div className="program-name">
        <span style={{ color: `#${program.color}` }}>{program.name}</span>
      </div>

      <div className="timeline" ref={timelineRef}>
        {/* Main timeline line - use program color */}
        {program.timePoints && program.timePoints.length > 0 && (
          <div className="timeline-line" style={{
            left: `${todayPos}px`,
            width: `${positions[positions.length - 1] - todayPos}px`,
            background: `#${program.color}`
          }}></div>
        )}

        {/* Gray timeline segment - from last event to Conference */}
        {program.conference && program.timePoints && program.timePoints.length > 0 && (
          <div className="timeline-line-gray" style={{
            left: `${positions[positions.length - 1]}px`,
            width: `${ddlPos - positions[positions.length - 1]}px`
          }}></div>
        )}

        {/* Today marker */}
        {program.ddl && (
          <div className="timeline-marker today-marker" style={{ left: `${todayPos}px` }}>
            <div className="marker-dot today-dot"></div>
            <div className="marker-label">
              <div className="marker-name">Today</div>
              <div className="marker-date">{today.toLocaleDateString('en-US')}</div>
            </div>
          </div>
        )}

        {/* Render Conference DDL node (full card) */}
        {program.conference && (
          <TimePoint
            timePoint={program.conference}
            position={ddlPos}
            today={today}
            ddl={program.ddl}
            canDrag={false}
            onDateChange={() => {}}
            pixelsPerDay={pixelsPerDay}
            prevDate={null}
            nextDate={null}
            offsetLevel={0}
            isConference={true}
            color="000000"
          />
        )}

        {/* Render time points */}
        {(() => {
          // Smart offset algorithm: use occupied ranges to avoid overlap
          const CARD_HALF_WIDTH = 130 // half card width (increased to fit larger cards)
          const occupiedRanges = [[], [], [], []] // occupied ranges for 4 vertical levels

          return program.timePoints && program.timePoints.map((timePoint, index) => {
            const prevTimePoint = index > 0 ? program.timePoints[index - 1] : null
            const nextTimePoint = index < program.timePoints.length - 1 ? program.timePoints[index + 1] : null

            const currentPos = positions[index]
            const cardStart = currentPos - CARD_HALF_WIDTH
            const cardEnd = currentPos + CARD_HALF_WIDTH

            // Function to check overlap between two ranges
            const checkOverlap = (range1Start, range1End, range2Start, range2End) => {
              return !(range1End <= range2Start || range1Start >= range2End)
            }

            // Find the first non-overlapping vertical level
            let offsetLevel = -1

            for (let level = 0; level <= 3; level++) {
              let hasOverlap = false

              // Check if current level overlaps with existing cards
              for (const range of occupiedRanges[level]) {
                if (checkOverlap(cardStart, cardEnd, range.start, range.end)) {
                  hasOverlap = true
                  break
                }
              }

              if (!hasOverlap) {
                offsetLevel = level
                break
              }
            }

            // If no suitable level found (all overlap), force using level 3
            if (offsetLevel === -1) {
              offsetLevel = 3
            }

            // Record occupied range for current card
            occupiedRanges[offsetLevel].push({ start: cardStart, end: cardEnd })

            return (
              <TimePoint
                key={timePoint.id}
                timePoint={timePoint}
                position={positions[index]}
                today={today}
                ddl={program.ddl}
                canDrag={!!program.ddl}
                onDateChange={(newDate) => onTimePointChange(timePoint.id, newDate)}
                pixelsPerDay={pixelsPerDay}
                prevDate={prevTimePoint ? prevTimePoint.date : null}
                nextDate={nextTimePoint ? nextTimePoint.date : null}
                offsetLevel={offsetLevel}
                color={program.color}
              />
            )
          })
        })()}

  {/* Render days label (displayed under the later node) */}
        {program.timePoints && program.timePoints.map((timePoint, index) => {
          if (index < program.timePoints.length - 1) {
            const nextTimePoint = program.timePoints[index + 1]
            const pos2 = positions[index + 1]

            const daysBetween = getDaysBetween(timePoint.date, nextTimePoint.date)

            return (
              <div key={`label-${timePoint.id}`}>
                {/* Days label displayed under the later node */}
                <div
                  className="days-between-label"
                  style={{ left: `${pos2}px`, color: `#${program.color}` }}
                >
                  <span style={{ color: `#${program.color}` }}>{daysBetween} days apart</span>
                </div>
              </div>
            )
          }
          return null
        })}
      </div>
    </div>
  )
}

export default TimeLine
