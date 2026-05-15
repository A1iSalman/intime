import { useState, useEffect } from 'react'
import './App.css'
import ExcelJS from 'exceljs'
import calendarImg from './assets/calendar.png'

function App() {
  // Date variables
  const now = new Date()
  const month = now.getMonth() + 1
  const monthName = now.toLocaleString('default', { month: 'long' })
  const year = now.getFullYear()
  const prevMonthName = new Date(year, month - 2).toLocaleString('default', { month: 'long' })
  const prevMonth = month === 1 ? 12 : month - 1
  const prevYear = month === 1 ? year - 1 : year

  // User state variables
  const [userName, setName] = useState("")
  const [userPosition, setPosition] = useState("")
  const [userIGG, setIGG] = useState("")
  const [userSGCnu, setSGCnu] = useState("")
  const [forCurrentMonth, setForCurrentMonth] = useState(true)

  // Input state
  const [leaveInput, setLeaveInput] = useState("")
  const [missionInput, setMissionInput] = useState("")
  const [trainingInput, setTrainingInput] = useState("")

  // Load saved user info
  useEffect(() => {
    const saved = JSON.parse(localStorage.getItem('userData') || '{}')
    if (saved.name) setName(saved.name)
    if (saved.position) setPosition(saved.position)
    if (saved.igg) setIGG(saved.igg)
    if (saved.sgc) setSGCnu(saved.sgc)
  }, [])

  async function generateTimesheet() {
    const response = await fetch('/Timesheet.xlsx')
    const buffer = await response.arrayBuffer()
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(buffer)
    const sheet = workbook.worksheets[0]

    const targetMonth = forCurrentMonth ? month : prevMonth
    const targetYear = forCurrentMonth ? year : prevYear
    const targetMonthName = forCurrentMonth ? monthName : prevMonthName

    // User Information
    sheet.getCell('C9').value = userName
    sheet.getCell('G9').value = userPosition
    sheet.getCell('L9').value = userIGG
    sheet.getCell('S9').value = parseInt(userSGCnu)
    sheet.getCell('AB5').value = targetMonthName
    sheet.getCell('AB7').value = targetYear

    // Holidays
    const holidays = await fetch('/holidays.txt')
      .then(res => res.text())
      .then(text => text.split('\n')
        .map(line => line.split('#')[0].trim())
        .filter(line => line.length > 0)
      )
    
    // Leave/ Mission/ Training
    const parseInput = (input) => input
      .replace(/\s/g, '')
      .split(',')
      .map(n => parseInt(n))
      .filter(n => !isNaN(n))

    const leaveDays = parseInput(leaveInput)
    const missionDays = parseInput(missionInput)
    const trainingDays = parseInput(trainingInput)

    // Fill Timesheet Data
    const daysInMonth = new Date(targetYear, targetMonth, 0).getDate()
    const startCell = 3
    const endCell = daysInMonth + startCell

    for (let col = startCell; col < endCell; col++) {
      const dayNum = col - 2
      const date = new Date(targetYear, targetMonth - 1, dayNum)
      const dayName = date.toLocaleString('default', { weekday: 'long' })
      const dateStr = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`

      if (dayName === 'Friday') {
        sheet.getCell(21, col).value = 1
        sheet.getCell(26, col).value = 'F'

      } else if (dayName === 'Saturday') {
        sheet.getCell(21, col).value = 1
        sheet.getCell(26, col).value = 'Y'

      } else if (holidays.includes(dateStr)) {
        sheet.getCell(21, col).value = 1
        sheet.getCell(26, col).value = 'H'

      } else if (leaveDays.includes(dayNum)) {
        sheet.getCell(22, col).value = 1
        sheet.getCell(26, col).value = 'N'

      } else if (missionDays.includes(dayNum)) {
        sheet.getCell(17, col).value = 1
        sheet.getCell(26, col).value = 'P'

      } else if (trainingDays.includes(dayNum)) {
        sheet.getCell(14, col).value = 1
        sheet.getCell(26, col).value = 'L'

      } else {
        sheet.getCell(14, col).value = 1
        sheet.getCell(26, col).value = 'T5'
      }
  }

    localStorage.setItem('userData', JSON.stringify({
      name: userName,
      position: userPosition,
      igg: userIGG,
      sgc: userSGCnu
    }))

    const out = await workbook.xlsx.writeBuffer()
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `${userName} Timesheet ${targetMonthName} ${targetYear}.xlsx`
    a.click()
  }

  return (
    <>
      <div className="mainWrapper">
        <div className="rightsCard">
          <div className="titleWrapper">
            <img src={calendarImg} />
            <h1>InTime</h1>
          </div>
          <p className='appInfo'>A light weight web application used to generate monthly timesheets for SGC secondees in Basra.</p>
          <p className='devInfo'>Developed by <span>Ali ALQATRANI</span></p>
        </div>
        <div className='card'>
          <div className="subWrapper">
            <p className="infoHeader">Employee Info</p>
            <div className="infoWrapper">
              <div>
                <p>Full Name</p>
                <input type="text" value={userName} spellCheck="false" onChange={(e) => {setName(e.target.value)}}/>
              </div>
              <div>
                <p>Position</p>
                <input type="text" value={userPosition} spellCheck="false" onChange={(e) => {setPosition(e.target.value)}}/>
              </div>
              <div>
                <p>IGG</p>
                <input type="text" value={userIGG} spellCheck="false" onChange={(e) => {setIGG(e.target.value)}}/>
              </div>
              <div>
                <p>SGC Number</p>
                <input type="text" value={userSGCnu} spellCheck="false" onChange={(e) => {setSGCnu(e.target.value)}}/>
              </div>
            </div>
          </div>
          <div className="subWrapper">
            <p className="infoHeader">period</p>
            <div className="monthSwitcher">
              <div 
                className={forCurrentMonth ? 'current activeMonth' : 'current'}
                onClick={() => setForCurrentMonth(true)}
              >{monthName + ", " + year}</div>
              <div 
                className={!forCurrentMonth ? 'previous activeMonth' : 'previous'}
                onClick={() => setForCurrentMonth(false)}
              >{prevMonthName + ", " + prevYear}</div>
            </div>
          </div> 
          <div className="subWrapper">
            <p className="infoHeader">attendance</p>
            <div className="infoWrapper">
              <div>
                <p>Leave</p>
                <input type="text" spellCheck="false" onChange={(e) => {setLeaveInput(e.target.value)}}/>
              </div>
              <div>
                <p>Mission</p>
                <input type="text" spellCheck="false" onChange={(e) => {setMissionInput(e.target.value)}}/>
              </div>
              <div>
                <p>Training</p>
                <input type="text" spellCheck="false" onChange={(e) => {setTrainingInput(e.target.value)}}/>
              </div>
            </div>
          </div>
          <button className='generateButton' onClick={generateTimesheet}>Generate</button>
        </div>
      </div>
    </>
  )
}

export default App
