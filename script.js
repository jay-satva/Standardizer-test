const mappingState = {}
let currentCategory = 'assets'
const usedCards = new Set()
let allDestinationAccounts = []

if (!localStorage.getItem("isLoggedIn")) {
    window.location.href = "login.html"
}

window.addEventListener("storage", function () {
  if (!localStorage.getItem("isLoggedIn")) {
      alert('Session expired')
      window.location.reload()
  }
})

function categorizeDestinationType(accountTypeName, subAccountName) {
    const type = (accountTypeName || '').toLowerCase().trim()
    const sub = (subAccountName || '').toLowerCase().trim()
    if (sub.includes('professional') && sub.includes('revenue')) return 'revenue'
    if (sub.includes('product') && sub.includes('revenue')) return 'revenue'
    if (sub.includes('professional') && sub.includes('cost')) return 'cogs'
    if (sub.includes('product') && sub.includes('cost')) return 'cogs'
    if (sub.includes('labor')) return 'expense'
    if (sub.includes('other') && (sub.includes('employee') || sub.includes('revenue') || sub.includes('expense'))) {
        return 'other rev & exp'
    }
    
    if (type.includes('asset')) return 'assets'
    if (type.includes('liabilit')) return 'liabilities'  
    if (type.includes('equity') || type.includes('capital')) return 'equity'  
    if (type.includes('revenue') || type.includes('income')) return 'revenue'
    if (type.includes('cog') || type.includes('cost of goods')) return 'cogs'
    if (type.includes('expense')) return 'expense'  
    console.warn('Uncategorized:', type, sub)
    return null
}

function convertToNested(rows) {
    const result = {}
    rows.forEach(row => {
        const type = (row.Type || "").toLowerCase().trim()
        const group = (row.Group || "").toLowerCase().trim()
        if (!type || !group || !row.Number) return

        if (!result[type]) result[type] = { group: {} }
        if (!result[type].group[group]) result[type].group[group] = {}

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
        const type = categorizeDestinationType(row.AccountTypeName, row.SubAccountName)
        if (!type) return
        
        const group = (row.SubAccountName || "").toLowerCase().trim()

        if (!result[type]) result[type] = { group: {} }
        if (!result[type].group[group]) result[type].group[group] = {}

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

function fetchExcel(filePath, storageKey, converterFn, done) {
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
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false })
            const structured = converterFn(rows)
            finalResult[sheetName] = structured
        })

        localStorage.setItem(storageKey, JSON.stringify(finalResult))
        if (done) done()
    }
    
    xhr.onerror = () => {
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
    div.innerHTML = `<b>${number}</b> — ${text}`
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

function getSourceAccountsByCategory(category) {
    const master = JSON.parse(localStorage.getItem("masterSheetData"))
    if (!master) return []
    
    const masterSheet = Object.values(master)[0]
    if (!masterSheet[category]) return []
    
    return flattenData({ [category]: masterSheet[category] })
}

function getDestAccountsByCategory(category) {
    const destination = JSON.parse(localStorage.getItem("destinationSheetData"))
    if (!destination) return []
    
    const destSheet = Object.values(destination)[0]
    if (!destSheet[category]) return []
    
    return flattenData({ [category]: destSheet[category] })
}

function updateUsedCards(category) {
    usedCards.clear()
    
    if (!mappingState[category]) return
    
    Object.values(mappingState[category]).forEach(state => {
        if (state.mostLikely) usedCards.add(String(state.mostLikely))
        if (state.likely) usedCards.add(String(state.likely))
        if (state.possible) usedCards.add(String(state.possible))
    })
}

function updateDestinationVisibility() {
    const destList = document.getElementById("a-accountContainer")
    if (!destList) return
    
    destList.querySelectorAll(".account-card").forEach(card => {
        const cardNumber = String(card.dataset.number)
        if (usedCards.has(cardNumber)) {
            card.style.display = 'none'
        } else {
            card.style.display = ''
        }
    })
}

function renderDestinationOnly(type) {

  const destList = document.getElementById("a-accountContainer")
  if (!destList) return

  destList.innerHTML = ""

  const destination = JSON.parse(localStorage.getItem("destinationSheetData"))
  if (!destination) return

  const destSheet = Object.values(destination)[0]
  if (allDestinationAccounts.length === 0) {
    allDestinationAccounts = flattenData(destSheet)
}

  let data = []
  if (type === "all") {
    data = allDestinationAccounts
} else {
    data = allDestinationAccounts.filter(a => a.type === type)
}


  data.forEach(acc => {
      const card = createCard(acc.name, acc.number, { type: acc.type })
      destList.appendChild(card)
  })

  // updateUsedCards(currentCategory)
  // updateDestinationVisibility()
}

function renderPage(category) {
    currentCategory = category
    
    const table = document.getElementById("mappingTableBody")
    const destList = document.getElementById("a-accountContainer")
    
    if (!table || !destList) {
        return
    }
    
    table.innerHTML = ""
    destList.innerHTML = ""
    
    const sourceAccounts = getSourceAccountsByCategory(category)
    const destAccounts = getDestAccountsByCategory(category)
    
    console.log(`Rendering ${category}:`, {
        source: sourceAccounts.length,
        dest: destAccounts.length
    })
    
    if (sourceAccounts.length === 0) {
        table.innerHTML = `<div class="alert alert-warning m-3">No source accounts for <b>${category}</b></div>`
    }
    
    if (destAccounts.length === 0) {
        destList.innerHTML = `<div class="alert alert-warning m-2 text-center">No destination accounts</div>`
    }
    
    if (!mappingState[category]) {
        mappingState[category] = {}
    }
  
    sourceAccounts.forEach(acc => {
        if (!mappingState[category][acc.number]) {
            mappingState[category][acc.number] = {
                mostLikely: null,
                likely: null,
                possible: null
            }
        }
        
        const row = document.createElement("div")
        row.className = "row g-2 mb-2 mapping-row"
        row.dataset.source = acc.number
        row.dataset.category = category

        row.innerHTML = `
        <div class="col-3">
        <div class="source-cell p-2 me-2 bg-light border rounded">${acc.number} — ${acc.name}</div>
        </div>
        <div class="col-3">
            <div class="slot-cell p-2 ms-1 me-2 border rounded" data-slot="mostLikely"></div>
        </div>
        <div class="col-3">
            <div class="slot-cell p-2 ms-2 me-1 border rounded" data-slot="likely"></div>
        </div>
        <div class="col-3">
            <div class="slot-cell p-2 ms-3 border rounded" data-slot="possible"></div>
        </div>
        `
        table.appendChild(row)
        renderRowFromState(category, acc.number)
    })
    
    // destAccounts.forEach(acc => {
    //     const card = createCard(acc.name, acc.number, { type: acc.type })
    //     destList.appendChild(card)
    // })
    renderDestinationOnly("all")

    
    // updateUsedCards(category)
    // updateDestinationVisibility()
    
    setTimeout(() => initSortable(), 100)
    updateSliderActiveState(category)
}

function renderRowFromState(category, sourceId) {
    const row = document.querySelector(`.mapping-row[data-source="${sourceId}"][data-category="${category}"]`)
    if (!row) return

    const state = mappingState[category]?.[sourceId]
    if (!state) return
    
    // const destAccounts = getDestAccountsByCategory(category)
    const destAccounts = allDestinationAccounts

    row.querySelectorAll(".slot-cell").forEach(cell => {
        const slot = cell.dataset.slot
        cell.innerHTML = ""
        const destNumber = state[slot]
        if (!destNumber) return

        const destAcc = destAccounts.find(a => a.number == destNumber)
        if (!destAcc) return
        
        const dropped = document.createElement("div")
        dropped.className = "dropped-card bg-white p-2 border rounded text-truncate"
        dropped.dataset.number = destAcc.number
        dropped.innerHTML = `<b>${destAcc.number}</b> — ${destAcc.name}`
        dropped.title = `${destAcc.number} — ${destAcc.name} (Click to remove)`
        dropped.style.cursor = 'pointer'
        
        dropped.addEventListener('click', () => {
            state[slot] = null
            renderRowFromState(category, sourceId)
            // updateUsedCards(category)
            // updateDestinationVisibility()
        })
        cell.appendChild(dropped)
    })
}

function initSortable() {
    const destList = document.getElementById("a-accountContainer")
    if (!destList) return
    
    if (destList.sortableInstance) {
        destList.sortableInstance.destroy()
    }
    
    if (typeof Sortable !== 'undefined') {
        destList.sortableInstance = new Sortable(destList, {
            group: { name: 'accounts', pull: 'clone', put: true },
            animation: 150,
            sort: false,
            onAdd: function(evt) {
              evt.item.remove()
          }
            // onEnd: () => updateDestinationVisibility()
        })
    }
    
    document.querySelectorAll(".slot-cell").forEach(cell => {
        if (cell.sortableInstance) cell.sortableInstance.destroy()
        
        cell.sortableInstance = new Sortable(cell, {
            group: { name: 'accounts', pull: true, put: true },
            animation: 150,
            sort: true,
            onAdd: function(evt) {

              const item = evt.item
              const sourceRow = evt.to.closest(".mapping-row")
              const sourceId = sourceRow.dataset.source
              const category = sourceRow.dataset.category
              const slot = evt.to.dataset.slot
          
              let destNumber
              if (item.classList.contains("account-card")) {
                  destNumber = item.dataset.number
              }
              else if (item.classList.contains("dropped-card")) {
                  destNumber = item.dataset.number
                  Object.values(mappingState[category]).forEach(rowState => {
                    Object.keys(rowState).forEach(k => {
                        if (rowState[k] == destNumber) rowState[k] = null
                    })
                  })

              }
          
              item.remove()
          
              handleDropSortable(category, sourceId, slot, destNumber)
          }
          
        })
    })
}

function handleDropSortable(category, sourceId, slot, destNumber) {

  if (!mappingState[category]) mappingState[category] = {}

  if (!mappingState[category][sourceId]) {
      mappingState[category][sourceId] = {
          mostLikely: null,
          likely: null,
          possible: null
      }
  }

  const state = mappingState[category][sourceId]
  if (
    state.mostLikely == destNumber ||
    state.likely == destNumber ||
    state.possible == destNumber
) {
    alert("This destination account is already used in this row.")
    return
}

  const order = ["mostLikely", "likely", "possible"]
  const index = order.indexOf(slot)
  let rightHasSpace = false

  for (let i = index; i < order.length; i++) {
      if (state[order[i]] == null) {
          rightHasSpace = true
          break
      }
  }

  if (rightHasSpace) {

      for (let i = order.length - 1; i > index; i--) {
          if (state[order[i]] == null) {
              for (let j = i; j > index; j--) {
                  state[order[j]] = state[order[j - 1]]
              }
              break
          }
      }
      state[slot] = destNumber
  }
  else {
      for (let i = order.length - 1; i > index; i--) {
          state[order[i]] = state[order[i - 1]]
      }
      state[slot] = destNumber
  }

  renderRowFromState(category, sourceId)
  // updateUsedCards(category)
  // updateDestinationVisibility()
}


function setupDestinationFilter() {
    const search = document.getElementById("search")
    const list = document.getElementById("a-accountContainer")
    const buttons = document.querySelectorAll(".nav-item-link")

    let currentType = "all"

    function applyFilter() {
        const text = search.value.toLowerCase()
        const normalize = s => (s || "").replace(/[^a-z0-9]/g, "")

        list.querySelectorAll(".account-card").forEach(card => {
            const cardNumber = String(card.dataset.number)
            
            if (usedCards.has(cardNumber)) {
                card.style.display = 'none'
                return
            }
            
            const matchText = card.dataset.name.includes(text) || card.dataset.number.includes(text)
            const matchType = currentType === "all" || normalize(card.dataset.type) === normalize(currentType)

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
  
          renderDestinationOnly(currentType)
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

function updateSliderActiveState(category) {
    const buttons = document.querySelectorAll(".nav-item-link")
    buttons.forEach(btn => {
        const normalize = s => (s || "").replace(/[^a-z0-9]/g, "")
        if (normalize(btn.dataset.type) === normalize(category)) {
            buttons.forEach(b => b.classList.remove("active"))
            btn.classList.add("active")
        }
    })
}

function setupCategoryNavigation() {
    const navButtons = document.querySelectorAll(".nav-btn")
    
    navButtons.forEach(btn => {
        btn.addEventListener("click", () => {
            const category = btn.dataset.page
            navButtons.forEach(b => {
                b.classList.remove("btn-primary")
                b.classList.add("btn-secondary")
            })
            btn.classList.remove("btn-secondary")
            btn.classList.add("btn-primary")
            renderPage(category)
        })
    })
}

function saveAllMappings() {
    localStorage.setItem("savedMappings", JSON.stringify(mappingState))
    alert("Mappings saved successfully!")
}

function confirmLogout(){
    if(confirm('Are you sure to logout?')){
      localStorage.clear()
      window.location.reload()
    }
}

function loadSavedMappings() {
    const saved = localStorage.getItem("savedMappings")
    if (saved) {
        Object.assign(mappingState, JSON.parse(saved))
        console.log("Loaded saved mappings")
    }
}

const logout = document.getElementById('logout')
if(logout) logout.addEventListener('click', confirmLogout)

function setupSubmitButton() {
    const submitBtn = document.querySelector(".btn-success")
    if (submitBtn) submitBtn.addEventListener("click", saveAllMappings)
}

function setupDeleteButton() {
    const deleteBtn = document.querySelector(".btn-danger")
    if (deleteBtn) {
        deleteBtn.addEventListener("click", () => {
            if (confirm("Clear all data?")) {
                localStorage.removeItem("savedMappings")
                location.reload()
            }
        })
    }
}

function debugAllTypes() {
    const master = JSON.parse(localStorage.getItem("masterSheetData"))
    const destination = JSON.parse(localStorage.getItem("destinationSheetData"))
    
    if (master) {
        const masterSheet = Object.values(master)[0]
        console.log("MASTER TYPES:", Object.keys(masterSheet))
    }
    
    if (destination) {
        const destSheet = Object.values(destination)[0]
        console.log("DESTINATION TYPES:", Object.keys(destSheet))
    }
}

document.addEventListener("DOMContentLoaded", () => {
    let pending = 0
    function finished() {
        pending--
        if (pending === 0) {
            debugAllTypes()
            loadSavedMappings()
            setupCategoryNavigation()
            setupDestinationFilter()
            setupSubmitButton()
            setupDeleteButton()
            renderPage('assets')
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
        debugAllTypes()
        loadSavedMappings()
        setupCategoryNavigation()
        setupDestinationFilter()
        setupSubmitButton()
        setupDeleteButton()
        renderPage('assets')
    }
})