const mappingState = {}

if (!localStorage.getItem("isLoggedIn")) {
    window.location.href = "login.html"
}

function convertToNested(rows) {

    const result = {}
    rows.forEach(row => {
        const type = (row.Type || "").toLowerCase()       
        const group = (row.Group || "").toLowerCase()      
        if (!type || !group || !row.Number) return

        if (!result[type]) {
            result[type] = { group: {} }
        }
        if (!result[type].group[group]) {
            result[type].group[group] = {}
        }

        const items = result[type].group[group]
        const index = Object.keys(items).length
        items[`item ${index}`] = {
            number: Number(row.Number),
            name: row.Name,
            subGroup: row["Sub-Group"],
            description: row.Description
        }
    })

    return result
}

function convertDestinationToNested(rows) {

    const result = {}

    rows.forEach(row => {

        let type = (row.AccountTypeName || "").toLowerCase()

        if (type.includes("liab")) type = "liability"
        if (type.includes("asset")) type = "assets"
        if (type.includes("equity")) type = "equity/capital"
        if (type.includes("revenue")) type = "revenue"
        if (type.includes("cog")) type = "cogs"
        if (type.includes("expense")) type = "g&a expenses"
        if (type.includes("other")) type = "other revenue and expense"

        const group = (row.SubAccountName || "").toLowerCase()

        if (!result[type]) {
            result[type] = { group: {} }
        }

        if (!result[type].group[group]) {
            result[type].group[group] = {}
        }

        const items = result[type].group[group]

        const index = Object.keys(items).length

        items[`item ${index}`] = {
            number: Number(row.AccountCode),
            name: row.AccountName,

            flags: {
                SpecialtyEmergencyCashBasis: row.SpecialtyEmergencyCashBasis,
                SpecialtyEmergencyAccrualBasis: row.SpecialtyEmergencyAccrualBasis,
                GeneralPracticeCashBasis: row.GeneralPracticeCashBasis,
                GeneralPracticeAccrualBasis: row.GeneralPracticeAccrualBasis,
                GPSpecERHybridCashBasis: row.GPSpecERHybridCashBasis,
                GPSpecERHybridAccrualBasis: row.GPSpecERHybridAccrualBasis,
                EquineCashBasis: row.EquineCashBasis,
                EquineAccuralBasis: row.EquineAccuralBasis
            }
        }
    })

    return result
}

function fetchExcel(filePath, storageKey, converterFn, done)
{

    const xhr = new XMLHttpRequest()
    xhr.open("GET", filePath, true)
    xhr.responseType = "arraybuffer"
    xhr.onload = function () {

        if (xhr.status !== 200) {
            console.error("Excel load failed:", filePath)
            if (done) done()
            return
        }
        const buffer = xhr.response
        const workbook = XLSX.read(buffer, { type: "array" })
        const finalResult = {}

        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName]
            const rows = XLSX.utils.sheet_to_json(sheet, {
                defval: "",
                raw: false
            })
            const structured = converterFn(rows)
            finalResult[sheetName] = structured
        })

        localStorage.setItem(storageKey, JSON.stringify(finalResult))
        if (done) done()
        console.log(storageKey, finalResult)
    }
    xhr.onerror = () => {
  console.error("XHR error")
  if (done) done()
}

    xhr.send()
}

function createCard(text, number, extra = {}) {
  const div = document.createElement("div")
  div.className = "account-card"
  div.dataset.name = text.toLowerCase()
  div.dataset.number = number
  div.dataset.type = (extra.type || "").toLowerCase()

  div.innerHTML = `
      <b>${number}</b> — ${text}
  `

  return div
}
function flattenData(data) {
  const list = []

  if (!data) return list

  Object.keys(data).forEach(type => {
    const groups = data[type]?.group || {}

    Object.keys(groups).forEach(group => {
      const items = groups[group] || {}

      Object.values(items).forEach(item => {
        list.push({
          number: item.number,
          name: item.name,
          type: type
        })
      })
    })
  })

  return list
}

function populateLists() {
  const master = JSON.parse(localStorage.getItem("masterSheetData"))
  const destination = JSON.parse(localStorage.getItem("destinationSheetData"))

  if (!master || !destination) return
  const table = document.getElementById("mappingTableBody")
  const destList = document.getElementById("a-accountContainer")
  
  if (!table) {
      console.error("mappingTableBody not found! Check your HTML.")
      return
  }
  
  if (!destList) {
      console.error("a-accountContainer not found! Check your HTML.")
      return
  }
  
  table.innerHTML = ""
  destList.innerHTML = ""

  const masterSheet = Object.values(master)[0]
  flattenData(masterSheet).forEach(acc => {

      if (!mappingState[acc.number]) {
          mappingState[acc.number] = {
              mostLikely: null,
              likely: null,
              possible: null
          }
      }
      const row = document.createElement("div")
      row.className = "row g-2 mb-2 mapping-row"
      row.dataset.source = acc.number

      row.innerHTML = `
          <div class="col-3">
              <div class="source-cell p-2 bg-light border rounded">${acc.number} — ${acc.name}</div>
          </div>
          <div class="col-3">
              <div class="slot-cell p-2 border rounded" data-slot="mostLikely"></div>
          </div>
          <div class="col-3">
              <div class="slot-cell p-2 border rounded" data-slot="likely"></div>
          </div>
          <div class="col-3">
              <div class="slot-cell p-2 border rounded" data-slot="possible"></div>
          </div>
      `

      table.appendChild(row)
  })
  
  const destSheet = Object.values(destination)[0]
  
  flattenData(destSheet).forEach(acc => {
      const card = createCard(acc.name, acc.number, { type: acc.type })
      card.setAttribute("draggable", "true")
      card.addEventListener("dragstart", (e) => {
          e.dataTransfer.setData("text/plain", card.dataset.number)
          e.dataTransfer.effectAllowed = "copy"
      })
      
      destList.appendChild(card)
  })
}
function renderRowFromState(sourceId) {
  const row = document.querySelector(`.mapping-row[data-source="${sourceId}"]`)
  if (!row) return

  const state = mappingState[sourceId]
  const destination = JSON.parse(localStorage.getItem("destinationSheetData"))
  const destSheet = Object.values(destination)[0]
  const destAccounts = flattenData(destSheet)

  row.querySelectorAll(".slot-cell").forEach(cell => {
      const slot = cell.dataset.slot
      cell.innerHTML = ""

      const destNumber = state[slot]
      if (!destNumber) return

      const destAcc = destAccounts.find(a => a.number == destNumber)
      if (!destAcc) return
      
      const dropped = document.createElement("div")
      dropped.className = "dropped-card bg-white p-2 border rounded"
      dropped.innerHTML = `<b>${destAcc.number}</b> — ${destAcc.name}`

      cell.appendChild(dropped)
  })
}
  
  
  function initDragDrop() {
    document.querySelectorAll(".slot-cell").forEach(cell => {
  
        cell.addEventListener("dragover", (e) => {
            e.preventDefault()
            e.dataTransfer.dropEffect = "copy"
            cell.classList.add("drag-over")
        })
  
        cell.addEventListener("dragleave", () => {
            cell.classList.remove("drag-over")
        })
  
        cell.addEventListener("drop", (e) => {
            e.preventDefault()
            cell.classList.remove("drag-over")
  
            const destNumber = e.dataTransfer.getData("text/plain")
            const sourceRow = cell.closest(".mapping-row")
            const sourceId = sourceRow.dataset.source
            const slot = cell.dataset.slot
  
            handleDrop(sourceId, slot, destNumber)
        })
    })
}
  function handleDrop(sourceId, slot, destNumber) {
    const state = mappingState[sourceId]
    if (state.mostLikely == destNumber ||
        state.likely == destNumber ||
        state.possible == destNumber) {
      return
    }
  
    if (slot === "mostLikely") {
      state.possible = state.likely
      state.likely = state.mostLikely
      state.mostLikely = destNumber
    } else if (slot === "likely") {
      state.possible = state.likely
      state.likely = destNumber
    } else if (slot === "possible") {
      state.possible = destNumber
    }
  
    renderRowFromState(sourceId)
  }
  

function setupDestinationFilter() {
    const search = document.getElementById("search")
    const list = document.getElementById("a-accountContainer")
    const buttons = document.querySelectorAll(".nav-item-link")
  
    let currentType = "all"
  
    function applyFilter() {
        const text = search.value.toLowerCase()
        const normalize = s => (s || "").replace(/[^a-z0-9]/g,"")
      
        list.querySelectorAll(".account-card").forEach(card => {
            const matchText =
                card.dataset.name.includes(text) ||
                card.dataset.number.includes(text)
      
            const matchType =
                currentType === "all" ||
                normalize(card.dataset.type) === normalize(currentType)
      
            card.style.display = (matchText && matchType) ? "" : "none"
        })
    }
  
    search.addEventListener("input", applyFilter)
  
    buttons.forEach(btn => {
        btn.addEventListener("click", (e) => {
            e.preventDefault()
            
            buttons.forEach(b => b.classList.remove("active"))
            btn.classList.add("active")
  
            currentType = btn.dataset.type
            applyFilter()
        })
    })
    const slider = document.getElementById('navSlider')
    const leftArrow = document.getElementById('slideLeft')
    const rightArrow = document.getElementById('slideRight')
 
    if (leftArrow && rightArrow && slider) {
        leftArrow.onclick = () => { slider.scrollLeft -= 200 }
        rightArrow.onclick = () => { slider.scrollLeft += 200 }
    }
} 

function shiftRight(row, startIndex) {

  const order = ["mostLikely", "likely", "possible"]

  for (let i = order.length - 1; i > startIndex; i--) {

    const current = row.querySelector("." + order[i])
    const previous = row.querySelector("." + order[i - 1])

    if (previous.children.length > 0) {

      const moving = previous.children[0]

      if (current.children.length > 0) {
        current.innerHTML = ""   
      }

      current.appendChild(moving)
    }
  }
}

function initSortable() {

  const group = {
    name: "accounts",
    pull: true,
    put: true
  }
  new Sortable(document.getElementById("destinationList"), {
    group,
    animation: 150
  })

  document.querySelectorAll(".drop-cell").forEach(cell => {

    new Sortable(cell, {
      group,
      animation: 150,

      onAdd: function(evt) {

        const row = evt.to.closest(".mapping-row")
        const cells = [
          row.querySelector(".mostLikely"),
          row.querySelector(".likely"),
          row.querySelector(".possible")
        ]

        const index = cells.indexOf(evt.to)

        shiftRight(cells, index)
      }
    })
  })
}

function shiftRight(cells, startIndex) {

  for (let i = cells.length - 1; i > startIndex; i--) {

    const prev = cells[i - 1].firstElementChild

    if (prev) {
      cells[i].innerHTML = ""
      cells[i].appendChild(prev)
    }
  }

  const current = cells[startIndex]

  if (current.children.length > 1) {
    current.removeChild(current.firstElementChild)
  }
}


document.addEventListener("DOMContentLoaded", () => {

  let pending = 0

  function finished() {
    pending--
    if (pending === 0) {
      populateLists()
      setupDestinationFilter()
    populateLists()
    setupDestinationFilter()
    initDragDrop()

    }
  }

  if (!localStorage.getItem("masterSheetData")) {
    pending++
    fetchExcel(
      "Resources/Master Chart of account.xlsx",
      "masterSheetData",
      convertToNested,
      finished
    )
  }

  if (!localStorage.getItem("destinationSheetData")) {
    pending++
    fetchExcel(
      "Resources/destination chart of account.xlsx",
      "destinationSheetData",
      convertDestinationToNested,
      finished
    )
  }

  if (pending === 0) {
    populateLists()
    setupDestinationFilter()    
    // initSortable()
    populateLists()
    setupDestinationFilter()
    initDragDrop()

  }
})