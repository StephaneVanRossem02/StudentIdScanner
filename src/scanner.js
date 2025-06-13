let scannedCode = "Aanwezig";
            let notScannedCode = "";

            let examName = "";
            let excelData = [];
            let scanCount = 0;
            let outCount = 0;
            let totalStudents = 0;

            let scanMode = "in"; // standaard
            let focusEnabled = false; // standaard focus uit

            document.getElementById("scanStatus").innerHTML = "Nog niets gescand";
            document.getElementById("fileInput").addEventListener("change", function (event) {
                document.getElementById("loadingSpinner").style.display = "block";

                document.getElementById("chooseFileBtn").classList.add("chosen");
                document.getElementById("fileName").classList.add("chosen");

                scanCount = 0;
                outCount = 0;
                document.getElementById("scanCounter").textContent = scanCount;
                document.getElementById("outCounter").textContent = outCount;

                let file = event.target.files[0];
                let reader = new FileReader();

                reader.onload = function (e) {
                    let data = new Uint8Array(e.target.result);
                    let workbook = XLSX.read(data, { type: "array" });
                    let sheet = workbook.Sheets[workbook.SheetNames[0]];

                    excelData = XLSX.utils.sheet_to_json(sheet, { header: 1 }).map((row) => row.slice(0, 3));

                    let table = document.getElementById("dataTable");
                    table.innerHTML = "";

                    // examName veilig extraheren
                    let lastDashIndex = file.name.lastIndexOf("-");
                    if (lastDashIndex !== -1) {
                        examName = file.name.substring(0, lastDashIndex).trim();
                    } else {
                        examName = file.name;
                    }

                    const input = document.getElementById("examInput");
                    if (input) {
                        input.value = examName;
                        sessionStorage.setItem("examName", examName);
                    }
                    // table vullen
                    excelData.forEach((row, index) => {
                        if (index === 0) {
                            row.push("Scanned In");
                            row.push("Scanned Out");
                        } else {
                            row.push("");
                            row.push("");
                        }

                        let tr = document.createElement("tr");
                        row.forEach((cell) => {
                            let cellElement = document.createElement(index === 0 ? "th" : "td");
                            cellElement.textContent = cell;
                            tr.appendChild(cellElement);
                        });
                        table.appendChild(tr);
                    });

                    totalStudents = excelData.length - 1;
                    document.getElementById("totalStudents").textContent = `Aantal studenten in lijst: ${totalStudents}`;
                    document.getElementById("scanTotal").textContent = totalStudents;
                    document.getElementById("scanTotalOut").textContent = totalStudents;

                    if (excelData.length > 1) {
                        // er moet minstens 1 rij + header zijn
                        document.getElementById("exportExcelBtn").disabled = false;
                        document.getElementById("exportPdfBtn").disabled = false;
                    } else {
                        document.getElementById("exportExcelBtn").disabled = true;
                        document.getElementById("exportPdfBtn").disabled = true;
                    }

                    document.getElementById("loadingSpinner").style.display = "none";

                    if (focusEnabled) document.getElementById("manualInput").focus();

                    sessionStorage.setItem("excelData", JSON.stringify(excelData));
                };

                reader.readAsArrayBuffer(file);
            });

            function getCurrentTimeString() {
                let now = new Date();
                return now.toLocaleTimeString("nl-BE", {
                    hour: "2-digit",
                    minute: "2-digit",
                    second: "2-digit",
                });
            }

            function checkBarcode(IDtoCheck) {
                let matchFound = false;
                let table = document.getElementById("dataTable");
                let rows = Array.from(table.rows).slice(1);
                let firstName, lastName;

                for (let row of rows) {
                    let idCell = row.cells[2];
                    if (idCell && idCell.textContent === IDtoCheck) {
                        // === Reset alle mogelijke klassen eerst ===
                        let rowIndex = excelData.findIndex((dataRow) => dataRow[2] === IDtoCheck);

                        firstName = row.cells[0].textContent;
                        lastName = row.cells[1].textContent;

                        if (scanMode === "in") {
                            if (row.cells[3].textContent.trim() === "" || row.cells[3].textContent === "NVT") {
                                row.cells[3].textContent = getCurrentTimeString();
                                row.cells[4].textContent = ""; // reset Out

                                row.classList.remove("match", "no-match", "row-in", "row-out");

                                // === Geef GEEL bij IN ===
                                row.classList.add("row-in");
                                excelData[rowIndex][3] = row.cells[3].textContent;
                                excelData[rowIndex][4] = row.cells[4].textContent;
                            } else {
                                document.getElementById("scanStatus").innerHTML = `<span class='error'>${firstName} ${lastName} (${IDtoCheck}) is al IN gescand → dubbele IN niet toegestaan</span>`;
                                return;
                            }
                        } else if (scanMode === "out") {
                            if (row.cells[3].textContent.trim() !== "" && row.cells[3].textContent !== "NVT") {
                                if (row.cells[4].textContent.trim() === "" || row.cells[4].textContent === "NVT") {
                                    row.cells[4].textContent = getCurrentTimeString();

                                    row.classList.remove("match", "no-match", "row-in", "row-out");

                                    // === Geef GROEN bij OUT ===
                                    row.classList.add("row-out");
                                    excelData[rowIndex][4] = row.cells[4].textContent;
                                } else {
                                    document.getElementById("scanStatus").innerHTML = `<span class='error'>${firstName} ${lastName} (${IDtoCheck}) is al OUT gescand → dubbele OUT niet toegestaan</span>`;
                                    return;
                                }
                            } else {
                                document.getElementById("scanStatus").innerHTML = `<span class='error'>${firstName} ${lastName} (${IDtoCheck}) heeft nog geen IN-scan → OUT niet toegestaan</span>`;
                                return;
                            }
                        }

                        matchFound = true;
                        table.insertBefore(row, table.rows[1]); // verplaats naar boven
                        break;
                    }
                }

                if (matchFound) {
                    document.getElementById("scanStatus").innerHTML = `<span class='success'>[${scanMode.toUpperCase()}] Gescanned: ${firstName} ${lastName} (${IDtoCheck})</span>`;
                } else {
                    document.getElementById("scanStatus").innerHTML = `<span class='error'>${IDtoCheck} staat niet in de lijst</span>`;
                }

                if (matchFound) {
                    document.getElementById("scanStatus").innerHTML = `<span class='success'>[${scanMode.toUpperCase()}] Gescanned: ${firstName} ${lastName} (${IDtoCheck})</span>`;

                    updateInOutCounters();
                    sortExcelDataForDisplay();
                    rebuildTable();
                    sessionStorage.setItem("excelData", JSON.stringify(excelData));
                }
            }

            function updateInOutCounters() {
                let table = document.getElementById("dataTable");
                let rows = Array.from(table.rows).slice(1);

                let inCount = 0;
                let outCountTemp = 0;

                for (let row of rows) {
                    if (row.cells[3].textContent.trim() !== "" && row.cells[3].textContent.trim() !== "NVT") {
                        inCount++;
                    }
                    if (row.cells[4].textContent.trim() !== "" && row.cells[4].textContent.trim() !== "NVT") {
                        outCountTemp++;
                    }
                }

                document.getElementById("scanCounter").textContent = inCount;
                document.getElementById("outCounter").textContent = outCountTemp;

                // Correcte elementen gebruiken → dit bestaat WEL in jouw HTML:
                document.getElementById("scanTotal").textContent = totalStudents;
                document.getElementById("scanTotalOut").textContent = totalStudents;
            }

            function handleManualInput(event) {
                if (event.key === "Enter") {
                    let input = event.target.value.trim();
                    if (input !== "") {
                        let cleaned = cleanBarcode(input);
                        if (cleaned !== "") {
                            checkBarcode(cleaned);
                        }
                        event.target.value = ""; // *** dit volstaat! ***
                        if (focusEnabled) {
                            event.target.focus(); // optioneel: focus behouden
                        }
                    }
                }
            }

            function cleanBarcode(barcode) {
                if (barcode.length >= 8) {
                    barcode = barcode.substring(barcode.length - 2 - 6, barcode.length - 2);
                    return "s" + barcode + "@ap.be";
                } else {
                    return barcode;
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

            function finalizeExcel(dataToExport) {
                let table = document.createElement("table");

                dataToExport.forEach((rowData, rowIndex) => {
                    const tr = document.createElement("tr");

                    rowData.forEach((cellText) => {
                        const cellTag = rowIndex === 0 ? "th" : "td";
                        const cell = document.createElement(cellTag);
                        cell.textContent = cellText;
                        tr.appendChild(cell);
                    });

                    if (rowIndex > 0) {
                        if (tr.cells[3].textContent.trim() === "") {
                            tr.cells[3].textContent = "NVT";
                        }
                        if (tr.cells[4].textContent.trim() === "" && tr.cells[3].textContent !== "NVT") {
                            tr.cells[4].textContent = "NVT";
                        }
                    }

                    table.appendChild(tr);
                });

                return table;
            }

            function exportToExcel() {
                const sortedExcelData = getSortedExcelDataCopy();
                const exportTable = finalizeExcel(sortedExcelData);

                insertInfoRow(exportTable, `Examen: ${examName}`, exportTable.rows[0].cells.length);
                insertInfoRow(exportTable, `Aantal gescande studenten IN: ${document.getElementById("scanCounter").textContent} / ${totalStudents}`, exportTable.rows[0].cells.length);
                insertInfoRow(exportTable, `Aantal gescande studenten OUT: ${document.getElementById("outCounter").textContent} / ${totalStudents}`, exportTable.rows[0].cells.length);
                insertInfoRow(exportTable, `Lokaal: ${document.getElementById("examLocationInput").value}`, exportTable.rows[0].cells.length);

                let ws = XLSX.utils.table_to_sheet(exportTable);
                let wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "Scans");
                XLSX.writeFile(wb, `${examName}.xlsx`);
            }

            function exportToPDF() {
                const sortedExcelData = getSortedExcelDataCopy();
                const exportTable = finalizeExcel(sortedExcelData);
                const { jsPDF } = window.jspdf;
                let doc = new jsPDF();

                doc.setFontSize(14);
                doc.setTextColor(0, 0, 0);

                let headerLeftX = 10;
                let headerTopY = 10;
                let lineHeight = 8;

                let examNameText = `Examen: ${examName}`;
                let examDateText = `Datum: ${formatDateForBE(document.getElementById("examDate").value)}`;
                let examTimeText = `Uur: ${document.getElementById("examTime").value}`;
                let examLocationText = `Lokaal: ${document.getElementById("examLocationInput").value}`;
                let scanInText = `Aantal gescande studenten IN: ${document.getElementById("scanCounter").textContent} / ${totalStudents}`;
                let scanOutText = `Aantal gescande studenten OUT: ${document.getElementById("outCounter").textContent} / ${totalStudents}`;

                doc.text(examNameText, headerLeftX, headerTopY);
                doc.text(examDateText, headerLeftX, headerTopY + lineHeight);
                doc.text(examTimeText, headerLeftX, headerTopY + lineHeight * 2);
                doc.text(examLocationText, headerLeftX, headerTopY + lineHeight * 3);
                doc.text(scanInText, headerLeftX, headerTopY + lineHeight * 4);
                doc.text(scanOutText, headerLeftX, headerTopY + lineHeight * 5);

                let headerHeight = lineHeight * 6 + 5;
                doc.rect(headerLeftX - 5, headerTopY - 5, 190, headerHeight, "S");

                let yPos = headerTopY + headerHeight + 5;

                let rows = Array.from(exportTable.rows);
                let header = rows.shift();
                let head = Array.from(header.cells).map((cell) => cell.textContent);
                let body = rows.map((row) => Array.from(row.cells).map((cell) => cell.textContent));

                doc.autoTable({
                    startY: yPos,
                    head: [head],
                    body: body,
                    styles: { fontSize: 10 },
                    headStyles: { fillColor: [227, 6, 19] },
                    columnStyles: {
                        3: { halign: "center", cellWidth: 25 },
                        4: { halign: "center", cellWidth: 25 },
                    },
                    margin: { top: 20 },
                    didDrawPage: function (data) {
                        let pageCount = doc.internal.getNumberOfPages();
                        let pageCurrent = doc.internal.getCurrentPageInfo().pageNumber;

                        let pageText = `Pagina ${pageCurrent} van ${pageCount}`;
                        let pageWidth = doc.internal.pageSize.getWidth();
                        let textWidth = doc.getTextWidth(pageText);

                        doc.setFontSize(10);
                        doc.text(pageText, pageWidth - textWidth - 10, doc.internal.pageSize.height - 10);
                    },
                });

                doc.save(`${examName}.pdf`);
            }

            function insertInfoRow(table, text, colspan) {
                let row = document.createElement("tr");
                let cell = document.createElement("td");
                cell.colSpan = colspan;
                cell.textContent = text;
                cell.style.fontWeight = "bold";
                row.appendChild(cell);
                table.insertBefore(row, table.firstChild);
            }

            document.getElementById("examInput").addEventListener("input", function () {
                sessionStorage.setItem("examName", this.value);
            });

            document.getElementById("examDate").addEventListener("input", function () {
                sessionStorage.setItem("examDate", this.value);
            });

            document.getElementById("examTime").addEventListener("input", function () {
                sessionStorage.setItem("examTime", this.value);
            });

            document.getElementById("examLocationInput").addEventListener("input", function () {
                sessionStorage.setItem("examLocation", this.value);
            });

            document.addEventListener("keydown", function (e) {
                if (e.key === "Enter") {
                    document.getElementById("manualInput").focus();
                }
            });

            function formatDateForBE(isoDateStr) {
                if (!isoDateStr) return "";
                let parts = isoDateStr.split("-");
                return `${parts[2]}/${parts[1]}/${parts[0]}`;
            }

            function updateExamName() {
                examName = document.getElementById("examInput").value;
                sessionStorage.setItem("examName", examName);
            }

            function getSortedExcelDataCopy() {
                const header = excelData[0];
                const dataRows = excelData.slice(1);

                dataRows.sort((a, b) => {
                    let scoreA = 2;
                    if (a[4] && a[4].trim() !== "" && a[4].trim() !== "NVT") {
                        scoreA = 0;
                    } else if (a[3] && a[3].trim() !== "" && a[3].trim() !== "NVT") {
                        scoreA = 1;
                    }

                    let scoreB = 2;
                    if (b[4] && b[4].trim() !== "" && b[4].trim() !== "NVT") {
                        scoreB = 0;
                    } else if (b[3] && b[3].trim() !== "" && b[3].trim() !== "NVT") {
                        scoreB = 1;
                    }

                    return scoreA - scoreB;
                });

                return [header, ...dataRows];
            }

            function sortExcelDataForDisplay() {
                // Sorteer niet de header
                const header = excelData[0];
                const dataRows = excelData.slice(1);

                dataRows.sort((a, b) => {
                    let scoreA = 2;
                    if (a[4] && a[4].trim() !== "" && a[4].trim() !== "NVT") {
                        scoreA = 0;
                    } else if (a[3] && a[3].trim() !== "" && a[3].trim() !== "NVT") {
                        scoreA = 1;
                    }

                    let scoreB = 2;
                    if (b[4] && b[4].trim() !== "" && b[4].trim() !== "NVT") {
                        scoreB = 0;
                    } else if (b[3] && b[3].trim() !== "" && b[3].trim() !== "NVT") {
                        scoreB = 1;
                    }

                    return scoreA - scoreB;
                });

                // Zet de header + gesorteerde rijen terug samen
                excelData = [header, ...dataRows];
            }

            function rebuildTable() {
                let table = document.getElementById("dataTable");
                table.innerHTML = "";

                excelData.forEach((rowData, rowIndex) => {
                    const tr = document.createElement("tr");

                    rowData.forEach((cellText, colIndex) => {
                        const cellTag = rowIndex === 0 ? "th" : "td";
                        const cell = document.createElement(cellTag);
                        cell.textContent = cellText;
                        tr.appendChild(cell);
                    });

                    if (rowIndex > 0) {
                        tr.classList.remove("match", "no-match", "row-in", "row-out");

                        const scannedIn = rowData[3];
                        const scannedOut = rowData[4];

                        if (scannedIn && scannedIn.trim() !== "" && scannedIn.trim() !== "NVT") {
                            tr.classList.add("row-in");
                        }

                        if (scannedOut && scannedOut.trim() !== "" && scannedOut.trim() !== "NVT") {
                            tr.classList.remove("row-in");
                            tr.classList.add("row-out");
                        }
                    }

                    table.appendChild(tr);
                });

                updateInOutCounters();
                sessionStorage.setItem("excelData", JSON.stringify(excelData));
            }

            window.addEventListener("DOMContentLoaded", function () {
                // === Veldwaarden terugzetten ===
                if (sessionStorage.getItem("examName")) {
                    document.getElementById("examInput").value = sessionStorage.getItem("examName");
                    examName = document.getElementById("examInput").value;
                }
                if (sessionStorage.getItem("examDate")) {
                    document.getElementById("examDate").value = sessionStorage.getItem("examDate");
                }
                if (sessionStorage.getItem("examTime")) {
                    document.getElementById("examTime").value = sessionStorage.getItem("examTime");
                }
                if (sessionStorage.getItem("examLocation")) {
                    document.getElementById("examLocationInput").value = sessionStorage.getItem("examLocation");
                }

                // === Tabel terugzetten uit sessionStorage ===
                if (sessionStorage.getItem("excelData")) {
                    excelData = JSON.parse(sessionStorage.getItem("excelData"));

                    let table = document.getElementById("dataTable");
                    table.innerHTML = "";

                    excelData.forEach((rowData, rowIndex) => {
                        const tr = document.createElement("tr");

                        rowData.forEach((cellText, colIndex) => {
                            const cellTag = rowIndex === 0 ? "th" : "td";
                            const cell = document.createElement(cellTag);
                            cell.textContent = cellText;
                            tr.appendChild(cell);
                        });

                        // === Kleuren ENKEL voor datarijen (niet header) ===
                        if (rowIndex > 0) {
                            tr.classList.remove("match", "no-match", "row-in", "row-out");

                            const scannedIn = rowData[3];
                            const scannedOut = rowData[4];

                            if (scannedIn && scannedIn.trim() !== "" && scannedIn.trim() !== "NVT") {
                                tr.classList.add("row-in");
                            }

                            if (scannedOut && scannedOut.trim() !== "" && scannedOut.trim() !== "NVT") {
                                tr.classList.remove("row-in");
                                tr.classList.add("row-out");
                            }
                        }

                        table.appendChild(tr);
                    });

                    // Update counters & enable export buttons
                    totalStudents = excelData.length - 1;
                    document.getElementById("totalStudents").textContent = `Aantal studenten in lijst: ${totalStudents}`;
                    document.getElementById("scanTotal").textContent = totalStudents;
                    document.getElementById("scanTotalOut").textContent = totalStudents;

                    updateInOutCounters();
                    document.getElementById("exportExcelBtn").disabled = false;
                    document.getElementById("exportPdfBtn").disabled = false;
                }
            });