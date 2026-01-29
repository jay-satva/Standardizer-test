if (!localStorage.getItem("isLoggedIn")) {
    window.location.href = "login.html"
}

document.addEventListener("DOMContentLoaded", () => {
    if (!localStorage.getItem("masterSheetData")) {
        fetchExcel("Resources/Master Chart of account.xlsx", "masterSheetData", convertToNested)
    }
    if (!localStorage.getItem("destinationSheetData")) {
        fetchExcel("Resources/destination chart of account.xlsx", "destinationSheetData", convertDestinationToNested)
    }
    console.log("Master:", JSON.parse(localStorage.getItem("masterSheetData")))
    console.log("Destination:", JSON.parse(localStorage.getItem("destinationSheetData")))
})


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

        const type = (row.AccountTypeName || "").toLowerCase()
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

function fetchExcel(filePath, storageKey, converterFn) {

    const xhr = new XMLHttpRequest()
    xhr.open("GET", filePath, true)
    xhr.responseType = "arraybuffer"
    xhr.onload = function () {

        if (xhr.status !== 200) {
            console.error("Excel load failed:", filePath)
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
        console.log(storageKey, finalResult)
    }
    xhr.onerror = () => console.error("XHR error")
    xhr.send()
}