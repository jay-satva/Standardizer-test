if (!localStorage.getItem("isLoggedIn")) {
    window.location.href = "login.html"
}

document.addEventListener("DOMContentLoaded", () => {
    const cached = localStorage.getItem("sheetData")
    if (cached) {
        console.log(JSON.parse(cached))
    } else {
        fetchExcel()
    }
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

function fetchExcel() {

    const xhr = new XMLHttpRequest()
    xhr.open("GET", "Resources/Master Chart of account.xlsx", true)
    xhr.responseType = "arraybuffer"
    xhr.onload = function () {

        if (xhr.status !== 200) {
            console.error("Excel load failed")
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
            const structured = convertToNested(rows)

            finalResult[sheetName] = structured
        })

        localStorage.setItem("sheetData", JSON.stringify(finalResult))
        console.log("Structured JSON:")
        console.log(finalResult)
    }

    xhr.onerror = () => console.error("XHR error")
    xhr.send()
}