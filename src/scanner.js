// === Globale variabelen ===
let scannedCode = "Aanwezig";
let notScannedCode = "";

let examName = "";
let excelData = [];
        try { document.getElementById("undoSelectedBtn").disabled = true; } catch(e) {}
let scanCount = 0;
let outCount = 0;
let totalStudents = 0;

let scanMode = "in";
let focusEnabled = false;

// === UNDO SUPPORT ===
let __undoKeyboardWired = false;
const historyByStudent = new Map(); // studentnummer -> [{prevIn, prevOut, ts}]

function pushHistory(studentnummer, prevIn, prevOut) {
    const arr = historyByStudent.get(studentnummer) || [];
    arr.push({ prevIn, prevOut, ts: Date.now() });
    historyByStudent.set(studentnummer, arr);
}

function undoSelectedRow() {
const table = document.getElementById("dataTable");
    const scanStatusEl = document.getElementById("scanStatus");
    if (!table) return;
    const tr = table.querySelector("tr.row-selected");
    if (!tr) {
        if (scanStatusEl) {
            scanStatusEl.style.display = "block";
            scanStatusEl.innerHTML = "<span class='error'>Selecteer eerst een rij.</span>";
        
    try { if (tr) tr.classList.remove("row-selected"); } catch(e) {}
}
        return;
    }
    const studentnummer = (tr.cells[2]?.textContent || "").trim();
    if (!studentnummer) return;

    const dataIndex = excelData.findIndex(r => r[2] === studentnummer);
    if (dataIndex <= 0) {
        if (scanStatusEl) {
            scanStatusEl.style.display = "block";
            scanStatusEl.innerHTML = "<span class='error'>Student niet gevonden in data.</span>";
        }
        return;
    }

    // Maak leeg: IN, OUT, OPM
    excelData[dataIndex][4] = "";
    excelData[dataIndex][5] = "";
    excelData[dataIndex][6] = "";

    // Bewaar en reload UI
    sessionStorage.setItem("excelData", JSON.stringify(excelData));
    rebuildTable();

            try { document.getElementById("undoSelectedBtn").disabled = false; } catch(e) {}
            try { document.getElementById("undoSelectedBtn").disabled = false; } catch(e) {}
    try { enableRowSelection(); } catch(e) {}
    try { updateInOutCounters(); } catch(e) {}

    if (scanStatusEl) {
        const firstName = (tr.cells[0]?.textContent || "").trim();
        const lastName  = (tr.cells[1]?.textContent || "").trim();
        scanStatusEl.style.display = "block";
        scanStatusEl.innerHTML = `<span class='success'>In/Out/Opmerkingen leeggemaakt voor ${firstName} ${lastName} (${studentnummer}).</span>`;
    }

}

function enableRowSelection() {
    // no-op: we gebruiken nu alleen delegated row selection
}



// === Delegated row selection (robust across rebuilds) ===
let __rowSelectionDelegated = false;
function setupDelegatedRowSelection() {
    if (__rowSelectionDelegated) return;
    const table = document.getElementById("dataTable");
    if (!table) return;
    table.addEventListener("click", (e) => {
        const tr = e.target.closest("tr");
        if (!tr) return;
        // skip header
        const header = table.querySelector("tr");
        if (tr === header) return;

        const prev = table.querySelector("tr.row-selected");
        if (prev && prev !== tr) prev.classList.remove("row-selected");
        tr.classList.toggle("row-selected");
    });
    __rowSelectionDelegated = true;
}



function createUndoButton() {
    // Zoek bestaande knop
    let btn = document.getElementById("undoSelectedBtn");

    // Bestaat hij niet? Maak 'm aan en voeg toe
    if (!btn) {
        btn = document.createElement("button");
        btn.id = "undoSelectedBtn";
        btn.type = "button";
        btn.textContent = "Undo geselecteerde rij";
        btn.style.margin = "8px";
        const targets = [
            document.getElementById("controls"),
            document.getElementById("buttons"),
            document.querySelector(".buttonRow"),
            document.getElementById("scanStatus"),
            document.body
        ];
        for (const t of targets) {
            if (t) { t.appendChild(btn); break; }
        }
    }

    // Wire de click handler als dit nog niet gebeurd is
    if (!btn.dataset || !btn.dataset.wired) {
        btn.addEventListener("click", undoSelectedRow);
        if (!btn.dataset) btn.dataset = {};
        btn.dataset.wired = "1";
    }

    // Keyboard shortcut: Ctrl/⌘ + Z (eens registreren)
    if (!__undoKeyboardWired) {
        document.addEventListener("keydown", (e) => {
            const key = (e.key || "").toLowerCase();
            if ((e.ctrlKey || e.metaKey) && key === "z") {
                e.preventDefault();
                undoSelectedRow();
            }
        });
        __undoKeyboardWired = true;
    }
}
// === END UNDO SUPPORT ===

// === Functies ===

function getCurrentTimeString() {
    let now = new Date();
    return now.toLocaleTimeString("nl-BE", {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
    });
}

function handleManualInput(event) {
    if (event.key === "Enter") {
        let input = event.target.value.trim();
        if (input !== "") {
            let firstTryID = extractID(input, 2);
            let secondTryID = extractID(input, 1);

            if (firstTryID) {
                checkBarcode(firstTryID);

                if (document.getElementById("scanStatus").innerText.includes("staat niet in de lijst")) {
                    if (secondTryID) {
                        checkBarcode(secondTryID);
                    }
                }
            } else {
                document.getElementById("scanStatus").innerHTML = `<span class='error'>Barcode niet herkend: ${input}</span>`;
            }

            event.target.value = "";
            if (focusEnabled) {
                event.target.focus();
            }
        }
    }
}


function extractID(barcode) {
    const sixDigitGroups = barcode.match(/\d{6}/g) || [];

    for (const cand of sixDigitGroups) {
        if (excelData?.some?.(row => row[2] === cand)) return cand;
    }
    if (sixDigitGroups.length === 1) return sixDigitGroups[0];

    const onlyDigits = barcode.replace(/\D/g, "");
    if (onlyDigits.length === 6) return onlyDigits;

    if (onlyDigits.length > 6) {
        for (let i = 0; i <= onlyDigits.length - 6; i++) {
            const cand = onlyDigits.slice(i, i + 6);
            if (excelData?.some?.(row => row[2] === cand)) return cand;
        }
        return onlyDigits.slice(-6);
    }
    return null;
}





	

function checkBarcode(IDtoCheck) {
    let table = document.getElementById("dataTable");
    let rows = Array.from(table.rows).slice(1);
    let matchFound = false;

    for (let row of rows) {
        let studentnummer = row.cells[2].textContent.trim();

        if (studentnummer === IDtoCheck) {
            let firstName = row.cells[0].textContent.trim();
            let lastName = row.cells[1].textContent.trim();
            const dataIndex = excelData.findIndex(r => r[2] === IDtoCheck);
            if (dataIndex <= 0) { break; }

            let nonParticipant = document.getElementById("nonParticipantToggle").checked;
            let scanStatusEl = document.getElementById("scanStatus");
            scanStatusEl.style.display = "block";
            try {
                const prevInHist = (row.cells[4]?.textContent || "").trim();
                const prevOutHist = (row.cells[5]?.textContent || "").trim();
                pushHistory(IDtoCheck, prevInHist, prevOutHist);
            } catch(e) {}

            if (scanMode === "in") {
                if (row.cells[4].textContent.trim() === "" || row.cells[4].textContent === "NVT") {
                    if (nonParticipant) {
                        row.cells[4].textContent = "Afgetekend zonder deelname";
                        row.cells[5].textContent = "Afgetekend zonder deelname";

                        excelData[dataIndex][4] = "Afgetekend zonder deelname";
                        excelData[dataIndex][5] = "Afgetekend zonder deelname";

                        row.classList.remove("match", "no-match", "row-in", "row-out");
                        row.classList.add("row-in");

                        scanStatusEl.innerHTML = `<span class='success'>[IN] ${firstName} ${lastName} (${IDtoCheck}) → Afgetekend zonder deelname</span>`;
                    } else {
                        let tijd = getCurrentTimeString();

                        row.cells[4].textContent = tijd;
                        row.cells[5].textContent = "";

                        excelData[dataIndex][4] = tijd;
                        excelData[dataIndex][5] = "";

                        row.classList.remove("match", "no-match", "row-in", "row-out");
                        row.classList.add("row-in");

                        scanStatusEl.innerHTML = `<span class='success'>[IN] ${firstName} ${lastName} (${IDtoCheck}) is IN gescand om ${tijd}</span>`;
                    }
                } else {
                    scanStatusEl.innerHTML = `<span class='error'>${firstName} ${lastName} (${IDtoCheck}) is al IN gescand → dubbele IN niet toegestaan</span>`;
                    return;
                }
            } else {
                if (row.cells[5].textContent.trim() === "" || row.cells[5].textContent === "NVT") {
                    if (nonParticipant) {
                        row.cells[5].textContent = "Afgetekend zonder deelname";

                        excelData[dataIndex][5] = "Afgetekend zonder deelname";

                        row.classList.remove("match", "no-match", "row-in", "row-out");
                        row.classList.add("row-out");

                        scanStatusEl.innerHTML = `<span class='success'>[OUT] ${firstName} ${lastName} (${IDtoCheck}) → Afgetekend zonder deelname</span>`;
                    } else {
                        let tijd = getCurrentTimeString();

                        row.cells[5].textContent = tijd;

                        excelData[dataIndex][5] = tijd;

                        row.classList.remove("match", "no-match", "row-in", "row-out");
                        row.classList.add("row-out");

                        scanStatusEl.innerHTML = `<span class='success'>[OUT] ${firstName} ${lastName} (${IDtoCheck}) is OUT gescand om ${tijd}</span>`;
                    }
                } else {
                    scanStatusEl.innerHTML = `<span class='error'>${firstName} ${lastName} (${IDtoCheck}) is al OUT gescand → dubbele OUT niet toegestaan</span>`;
                    return;
                }
            }

            table.insertBefore(row, table.rows[1]);

            updateInOutCounters();
            sessionStorage.setItem("excelData", JSON.stringify(excelData));

            matchFound = true;
            break;
        }
    }

    if (!matchFound) {
        const scanStatusEl = document.getElementById("scanStatus");
        scanStatusEl.style.display = "block";
        scanStatusEl.innerHTML = `<span class='error'>Studentnummer ${IDtoCheck} niet gevonden in lijst!</span>`;
    }
}



function updateInOutCounters() {
    let table = document.getElementById("dataTable");
    let rows = Array.from(table.rows).slice(1);

    let inCount = 0;
    let outCountTemp = 0;

    for (let row of rows) {
        if (row.cells[4].textContent.trim() !== "" && row.cells[4].textContent.trim() !== "NVT") {
            inCount++;
        }
        if (row.cells[5].textContent.trim() !== "" && row.cells[5].textContent.trim() !== "NVT") {
            outCountTemp++;
        }
    }

    const scanCounterEl = document.getElementById("scanCounter");
    if (scanCounterEl) {
        scanCounterEl.textContent = inCount;
    }

    const outCounterEl = document.getElementById("outCounter");
    if (outCounterEl) {
        outCounterEl.textContent = outCountTemp;
    }
}


function toggleScanMode() {
    scanMode = document.getElementById("scanModeSwitch").checked ? "out" : "in";
    document.getElementById("scanModeLabel").textContent = scanMode.toUpperCase();
}

function toggleFocus() {
    focusEnabled = document.getElementById("focusToggle").checked;
    if (focusEnabled) {
        document.getElementById("manualInput").focus();
    }
}

function rebuildTable() {
    let table = document.getElementById("dataTable");
    table.innerHTML = "";

    const sortedData = getSortedExcelDataCopy();

    sortedData.forEach((row, index) => {
        // Zet NVT indien leeg
        if (index > 0) {
            if (row[4].trim() === "") {
                row[4] = "NVT";
            }
            if (row[5].trim() === "") {
                row[5] = "NVT";
            }
            if (!row[6]) {
                row[6] = "";
            }
        }

        // Bouw de rij
        let tr = document.createElement("tr");

        row.forEach((cell, cellIndex) => {
            let cellElement = document.createElement(index === 0 ? "th" : "td");

            // Voor opmerkingen-kolom → toon als textContent
            if (cellIndex === 6 && index > 0) {
                // Opmerkingen tonen als lijst
                cellElement.innerHTML = cell
                    .split("\n")
                    .map(r => `<div>${r}</div>`)
                    .join("");

                // Voeg knop toe in aparte div
                let buttonDiv = document.createElement("div");
                buttonDiv.style.marginTop = "5px";

                let button = document.createElement("button");
                button.textContent = "📝 Opmerking";
                button.onclick = function() {
                    addRemark(row[2]); // studentnummer
                };
                buttonDiv.appendChild(button);

                cellElement.appendChild(buttonDiv);
            } else {
                // Voor andere kolommen → normale tekst
                cellElement.textContent = cell;
            }

            tr.appendChild(cellElement);
        });

        // Kleur toepassen
        if (index > 0) {
            tr.classList.remove("match", "no-match", "row-in", "row-out");

            const scannedIn = row[4];
            const scannedOut = row[5];

            if (scannedIn !== "NVT") {
                tr.classList.add("row-in");
            }

            if (scannedOut !== "NVT") {
                tr.classList.remove("row-in");
                tr.classList.add("row-out");
            }
            if (row[7] === "Afgetekend zonder deelname") {
                tr.classList.add("row-nonparticipant");
            }
        }

        table.appendChild(tr);
    });

    let totalStudents = excelData.length - 1;
    document.getElementById("totalStudents").textContent = totalStudents;

    updateInOutCounters();
    sessionStorage.setItem("excelData", JSON.stringify(excelData));
    try { applyRowClasses(); } catch(e) {}    try { enableRowSelection(); } catch(e) {}
}




function insertInfoRow(table, text, colspan = 2) {
    let row = table.insertRow(0);
    let cell = row.insertCell(0);
    cell.textContent = text;
    cell.colSpan = colspan;
}


function getSortedExcelDataCopy() {
    const header = excelData[0];
    const dataRows = excelData.slice(1);

    dataRows.sort((a, b) => {
        let scoreA = 2;
        if (a[5] && a[5].trim() !== "" && a[5].trim() !== "NVT") {
            scoreA = 0;
        } else if (a[4] && a[4].trim() !== "" && a[4].trim() !== "NVT") {
            scoreA = 1;
        }

        let scoreB = 2;
        if (b[5] && b[5].trim() !== "" && b[5].trim() !== "NVT") {
            scoreB = 0;
        } else if (b[4] && b[4].trim() !== "" && b[4].trim() !== "NVT") {
            scoreB = 1;
        }

        return scoreA - scoreB;
    });

    return [header, ...dataRows];
}

function finalizeExcel() {
    const sortedData = getSortedExcelDataCopy();
    const date = document.getElementById("examDate").value;
    const time = document.getElementById("examTime").value;
    const location = document.getElementById("examLocationInput").value;

    const headerRow = [`Examen: ${examName}`, "", "", "", "", "", ""];
    const infoRow = ["Datum:", date, "Uur:", time, "Locatie:", location, "", ""];

    const emptyRow = ["", "", "", "", "", "", ""];

    return [headerRow, infoRow, emptyRow, sortedData[0], ...sortedData.slice(1)];
}


function exportToExcel() {
    const exportData = finalizeExcel();

    const ws = XLSX.utils.aoa_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Scan resultaten");

    XLSX.writeFile(wb, `${examName}.xlsx`);
}




function exportToPDF() {
    const sortedData = getSortedExcelDataCopy();
    const doc = new jspdf.jsPDF(); // standaard portrait A4

    const columns = sortedData[0];
    const rows = sortedData.slice(1);

    const head = columns;
    const body = rows;

    // Datum als DD/MM/YYYY
    const dateValue = document.getElementById("examDate").value;
    const dateObj = new Date(dateValue);
    const day = String(dateObj.getDate()).padStart(2, '0');
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const year = dateObj.getFullYear();
    const examDateFormatted = `${day}/${month}/${year}`;

    // Header info
    let examNameText = `Examen: ${examName}`;
    let examDateText = `Datum: ${examDateFormatted}`;
    let examTimeText = `Uur: ${document.getElementById("examTime").value}`;
    let examLocationText = `Lokaal: ${document.getElementById("examLocationInput").value}`;
    let scanInText = `Aantal IN: ${document.getElementById("scanCounter").textContent}`;
    let scanOutText = `Aantal OUT: ${document.getElementById("outCounter").textContent}`;

    let headerLeftX = 15;
    let headerTopY = 20;
    let lineHeight = 7;

    doc.setFontSize(10);
    doc.text(examNameText, headerLeftX, headerTopY);
    doc.text(examDateText, headerLeftX, headerTopY + lineHeight);
    doc.text(examTimeText, headerLeftX, headerTopY + lineHeight * 2);
    doc.text(examLocationText, headerLeftX, headerTopY + lineHeight * 3);
    doc.text(scanInText, headerLeftX, headerTopY + lineHeight * 4);
    doc.text(scanOutText, headerLeftX, headerTopY + lineHeight * 5);

    let headerHeight = lineHeight * 6 + 5;
    doc.rect(headerLeftX - 5, headerTopY - 5, 190, headerHeight, "S");

    let yPos = headerTopY + headerHeight + 5;

    // AutoTable
    doc.autoTable({
        startY: yPos,
        head: [head],
        body: body,
        styles: {
            fontSize: 9
        },
        headStyles: {
            fillColor: [227, 6, 19]
        },
        theme: 'grid',
        tableWidth: 'auto', // AUTOMATISCH aanpassen
        columnStyles: {
            0: {
                cellWidth: 'auto'
            }, // Voornaam
            1: {
                cellWidth: 'auto'
            }, // Achternaam
            2: {
                cellWidth: 'auto'
            }, // Pointer
            3: {
                cellWidth: 'auto'
            }, // Subgroep
            4: {
                halign: "center",
                cellWidth: 'auto'
            }, // IN
            5: {
                halign: "center",
                cellWidth: 'auto'
            }, // OUT
            6: {
                cellWidth: 'auto'
            }, // Opmerkingen
        },

        didDrawPage: function(data) {
            let pageCount = doc.internal.getNumberOfPages();
            let pageCurrent = doc.internal.getCurrentPageInfo().pageNumber;

            let pageText = `Pagina ${pageCurrent} van ${pageCount}`;
            let pageWidth = doc.internal.pageSize.getWidth();
            let textWidth = doc.getTextWidth(pageText);

            doc.setFontSize(9);
            doc.text(pageText, pageWidth - textWidth - 10, doc.internal.pageSize.height - 10);
        },
    });

    doc.save(`${examName}.pdf`);
}



function addRemark(studentNumber) {
    let studentIndex = excelData.findIndex(row => row[2] === studentNumber);

    if (studentIndex > 0) {
        let newRemark = prompt("Voer opmerking in voor student " + studentNumber + ":");

        if (newRemark !== null && newRemark.trim() !== "") {
            let now = new Date();
            let time = now.toLocaleTimeString("nl-BE", {
                hour: "2-digit",
                minute: "2-digit",
                second: "2-digit",
            });

            let existingRemarks = excelData[studentIndex][6] ? excelData[studentIndex][6] : "";

            if (existingRemarks !== "") {
                existingRemarks += "\n";
            }
            existingRemarks += `${time} - ${newRemark.trim()}`;

            excelData[studentIndex][6] = existingRemarks;

            rebuildTable();
            try { document.getElementById("undoSelectedBtn").disabled = false; } catch(e) {}
            sessionStorage.setItem("excelData", JSON.stringify(excelData));
        }
    } else {
        console.log("Student niet gevonden!");
    }
}

function filterTable() {
    const input = document.getElementById("searchInput").value.toLowerCase();
    const rows = document.getElementById("dataTable").getElementsByTagName("tr");

    // Skip header (index 0)
    for (let i = 1; i < rows.length; i++) {
        const cells = rows[i].getElementsByTagName("td");

        // Match op voornaam (0), achternaam (1), pointer (2)
        const firstName = cells[0] ? cells[0].textContent.toLowerCase() : "";
        const lastName = cells[1] ? cells[1].textContent.toLowerCase() : "";
        const pointer = cells[2] ? cells[2].textContent.toLowerCase() : "";

        if (firstName.includes(input) || lastName.includes(input) || pointer.includes(input)) {
            rows[i].style.display = "";
        } else {
            rows[i].style.display = "none";
        }
    }
}



// === DOMContentLoaded ===
window.addEventListener("DOMContentLoaded", function() {
    
    // Undo disabled by default
    try { document.getElementById("undoSelectedBtn").disabled = true; } catch(e) {}
try { setupDelegatedRowSelection(); } catch(e) {}

    document.getElementById("manualInput").addEventListener("keydown", handleManualInput);
    document.getElementById("scanModeSwitch").addEventListener("change", toggleScanMode);
    document.getElementById("focusToggle").addEventListener("change", toggleFocus);

    document.getElementById("exportExcelBtn").addEventListener("click", exportToExcel);
    document.getElementById("exportPdfBtn").addEventListener("click", exportToPDF);

    document.getElementById("scanStatus").innerHTML = "Nog niets gescand";

    try { createUndoButton(); } catch(e) {}
    try { enableRowSelection(); } catch(e) {}

    // === Restore examName, date, time, location ===
    if (sessionStorage.getItem("examName")) {
        document.getElementById("examInput").value = sessionStorage.getItem("examName");
        examName = sessionStorage.getItem("examName");
    

    try { createUndoButton(); } catch(e) {}
    try { enableRowSelection(); } catch(e) {}
}
    if (sessionStorage.getItem("examDate")) {
        document.getElementById("examDate").value = sessionStorage.getItem("examDate");
    } else {
        // default datum als leeg → vandaag zetten:
        const today = new Date();
        const day = String(today.getDate()).padStart(2, '0');
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const year = today.getFullYear();
        document.getElementById("examDate").value = `${year}-${month}-${day}`;
    }
    if (sessionStorage.getItem("examTime")) {
        document.getElementById("examTime").value = sessionStorage.getItem("examTime");
    }
    if (sessionStorage.getItem("examLocation")) {
        document.getElementById("examLocationInput").value = sessionStorage.getItem("examLocation");
    }

    // === Listeners voor bijhouden in sessionStorage ===
    document.getElementById("examInput").addEventListener("input", function() {
        sessionStorage.setItem("examName", this.value);
        examName = this.value;
    });
    document.getElementById("examDate").addEventListener("input", function() {
        sessionStorage.setItem("examDate", this.value);
    });
    document.getElementById("examTime").addEventListener("input", function() {
        sessionStorage.setItem("examTime", this.value);
    });
    document.getElementById("examLocationInput").addEventListener("input", function() {
        sessionStorage.setItem("examLocation", this.value);
    });

    // === Restore excelData als het bestaat ===
    if (sessionStorage.getItem("excelData")) {
        try { document.getElementById("undoSelectedBtn").disabled = false; } catch(e) {}
        excelData = JSON.parse(sessionStorage.getItem("excelData"));
        rebuildTable();
            try { document.getElementById("undoSelectedBtn").disabled = false; } catch(e) {}
        document.getElementById("exportExcelBtn").disabled = false;
        document.getElementById("exportPdfBtn").disabled = false;
    }

    // === File input voor Excel ===
    document.getElementById("fileInput").addEventListener("change", function(event) {

        scanCount = 0;
        outCount = 0;

        let file = event.target.files[0];
        let reader = new FileReader();

        reader.onload = function(e) {
            let data = new Uint8Array(e.target.result);
            let workbook = XLSX.read(data, {
                type: "array"
            });
            let sheet = workbook.Sheets[workbook.SheetNames[0]];

            let rawData = XLSX.utils.sheet_to_json(sheet, {
                header: 1
            });

            const headerRow = rawData[0];
            const pointerCol = headerRow.findIndex(col => col.toLowerCase().includes("pointer"));
            const studentCol = headerRow.findIndex(col => col.toLowerCase().includes("student"));
            const subgroepCol = headerRow.indexOf("Subgroep");
            const vrijstellingCol = headerRow.indexOf("Vrijstelling");

            excelData = [
                ["Voornaam", "Achternaam", "Pointer", "Subgroep", "Scanned In", "Scanned Out", "Opmerkingen"]
            ];

            for (let i = 1; i < rawData.length; i++) {
                const row = rawData[i];
                const pointer = row[pointerCol] ? String(row[pointerCol]).trim() : "";
                const studentNaam = row[studentCol] ? String(row[studentCol]).trim() : "";
                const subgroep = row[subgroepCol] ? String(row[subgroepCol]).trim() : "";
                const vrijstelling = row[vrijstellingCol] ? String(row[vrijstellingCol]).trim() : "";

                if (vrijstelling.toUpperCase() === "EVK") {
                    continue;
                }

                const nameParts = studentNaam.split(" ");
                const voornaam = nameParts.pop(); // laatste woord = voornaam
                const achternaam = nameParts.join(" "); // rest = achternaam

                excelData.push([voornaam, achternaam, pointer, subgroep, "", "", ""]);
            }

            rebuildTable();
            try { document.getElementById("undoSelectedBtn").disabled = false; } catch(e) {}

            totalStudents = excelData.length - 1;
            document.getElementById("exportExcelBtn").disabled = false;
            document.getElementById("exportPdfBtn").disabled = false;

            if (focusEnabled) document.getElementById("manualInput").focus();

            sessionStorage.setItem("excelData", JSON.stringify(excelData));

            // Bestandsnaam automatisch zetten
            let lastDashIndex = file.name.lastIndexOf("-");
            if (lastDashIndex !== -1) {
                examName = file.name.substring(0, lastDashIndex).trim();
            } else {
                examName = file.name;
            }

            examName = examName.replace(/\.[^/.]+$/, "");

            const input = document.getElementById("examInput");
            if (input) {
                input.value = examName;
                sessionStorage.setItem("examName", examName);
            }
        };

        reader.readAsArrayBuffer(file);
    });
});

// Helper om rijen consequent te kleuren (IN/OUT/NonParticipant)
function applyRowClasses() {
    const table = document.getElementById("dataTable");
    if (!table) return;
    const rows = Array.from(table.rows).slice(1);
    for (const tr of rows) {
        tr.classList.remove("match","no-match","row-in","row-out","row-nonparticipant");
        const scannedIn  = (tr.cells[4]?.textContent || "").trim();
        const scannedOut = (tr.cells[5]?.textContent || "").trim();
        const isNonParticipant = scannedIn === "Afgetekend zonder deelname" || scannedOut === "Afgetekend zonder deelname";
        if (scannedIn && scannedIn !== "NVT") tr.classList.add("row-in");
        if (scannedOut && scannedOut !== "NVT") { tr.classList.remove("row-in"); tr.classList.add("row-out"); }
        if (isNonParticipant) tr.classList.add("row-nonparticipant");
    }
}
