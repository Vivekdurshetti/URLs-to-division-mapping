let urlDivisionMap = {};
let filteredResults = [];

let currentPage = 1;

const pageSize = 25;

// ========================================
// FILE UPLOAD
// ========================================

document
.getElementById('fileInput')
.addEventListener('change', handleFile);

const dropZone =
document.getElementById('dropZone');

dropZone.addEventListener('dragover', e => {

    e.preventDefault();

    dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {

    dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', e => {

    e.preventDefault();

    dropZone.classList.remove('dragover');

    const file = e.dataTransfer.files[0];

    if(file){

        processFile(file);
    }
});

function handleFile(e){

    const file = e.target.files[0];

    if(file){

        processFile(file);
    }
}

function processFile(file){

    const reader = new FileReader();

    const isExcel =
        file.name.endsWith(".xlsx") ||
        file.name.endsWith(".xls");

    reader.onload = function(e){

        if(isExcel){

            parseExcel(e.target.result);

        }else{

            parseCSV(e.target.result);
        }
    };

    if(isExcel){

        reader.readAsArrayBuffer(file);

    }else{

        reader.readAsText(file);
    }
}

// ========================================
// PARSE EXCEL
// ========================================

function parseExcel(data){

    const workbook =
        XLSX.read(new Uint8Array(data), {
            type:'array'
        });

    const firstSheet =
        workbook.Sheets[workbook.SheetNames[0]];

    const excelData =
        XLSX.utils.sheet_to_json(firstSheet, {
            header:1
        });

    processRows(excelData);
}

// ========================================
// PARSE CSV
// ========================================

function parseCSV(text){

    const rows =
        text
        .split(/\r?\n/)
        .map(r => r.split(","));

    processRows(rows);
}

// ========================================
// PROCESS ROWS
// ========================================

function processRows(rows){

    if(rows.length < 2){

        alert("Invalid file");

        return;
    }

    const headers = rows[0].map(h =>
        h.toString().trim().toLowerCase()
    );

    const getIndex = name =>
        headers.indexOf(name.toLowerCase());

    const pageUrlIndex =
        getIndex("page urls");

    if(pageUrlIndex === -1){

        alert("Missing 'Page URLs' column");

        return;
    }

    urlDivisionMap = {};

    for(let i=1; i<rows.length; i++){

        const row = rows[i];

        if(!row || row.length === 0) continue;

        const url =
            row[pageUrlIndex]
            ? normalizeUrl(
                row[pageUrlIndex].toString()
            )
            : "";

        if(!url) continue;

        urlDivisionMap[url] = {

            id: row[getIndex("id")] || "",

            language:
                row[getIndex("language")] || "",

            path:
                row[getIndex("path")] || "",

            division:
                row[getIndex("divisions")] || "",

            subDivision:
                row[getIndex("sub divisions")] || "",

            pageType:
                row[getIndex("page types")] || "",

            templateName:
                row[getIndex("template name")] || "",

            createdDate:
                row[getIndex("page cretaed date")] || "",

            pageName:
                row[getIndex("page name")] || "",

            branchTemplate:
                row[getIndex("branch template")] || ""
        };
    }

    enableControls();

    alert(
        "Loaded " +
        Object.keys(urlDivisionMap).length +
        " URLs successfully."
    );
}

// ========================================
// ENABLE CONTROLS
// ========================================

function enableControls(){

    document.getElementById('lookupBtn').disabled = false;
    document.getElementById('clearBtn').disabled = false;
    document.getElementById('downloadExcelBtn').disabled = false;
    document.getElementById('downloadCsvBtn').disabled = false;
    document.getElementById('filterInput').disabled = false;
    document.getElementById('searchInput').disabled = false;
}

// ========================================
// NORMALIZE URL
// ========================================

function normalizeUrl(url){

    return url
        .trim()
        .replace(/\/$/, "")
        .toLowerCase();
}

// ========================================
// LOOKUP
// ========================================

document
.getElementById('lookupBtn')
.addEventListener('click', () => {

    const urls =
        document
        .getElementById('urlInput')
        .value
        .split(/\n/)
        .map(u => normalizeUrl(u))
        .filter(Boolean);

    if(urls.length === 0){

        alert("Enter URLs");

        return;
    }

    const seen = new Set();

    filteredResults = urls.map(url => {

        const data =
            urlDivisionMap[url];

        const duplicate =
            seen.has(url);

        seen.add(url);

        return {

            url,

            duplicate,

            ...(data || {
                division:"Not Found"
            })
        };
    });

    currentPage = 1;

    renderResults();

    renderStats();
});

// ========================================
// RENDER STATS
// ========================================

function renderStats(){

    const total =
        filteredResults.length;

    const found =
        filteredResults.filter(r =>
            r.division !== "Not Found"
        ).length;

    const notFound =
        total - found;

    const duplicates =
        filteredResults.filter(r =>
            r.duplicate
        ).length;

    const divisions = {};

    filteredResults.forEach(r => {

        if(
            r.division &&
            r.division !== "Not Found"
        ){

            divisions[r.division] =
                (divisions[r.division] || 0) + 1;
        }
    });

    let divisionHTML = `
        <table class="summary-table">
            <tr>
                <th>Division</th>
                <th>Count</th>
            </tr>
    `;

    Object.entries(divisions)
    .sort((a,b)=>b[1]-a[1])
    .forEach(([division,count]) => {

        divisionHTML += `
            <tr>
                <td>${division}</td>
                <td>${count}</td>
            </tr>
        `;
    });

    divisionHTML += `</table>`;

    document.getElementById('stats').innerHTML = `

        <div class="stat-box">
            <h3>Total URLs</h3>
            <p>${total}</p>
        </div>

        <div class="stat-box">
            <h3>Found</h3>
            <p>${found}</p>
        </div>

        <div class="stat-box">
            <h3>Not Found</h3>
            <p>${notFound}</p>
        </div>

        <div class="stat-box">
            <h3>Duplicates</h3>
            <p>${duplicates}</p>
        </div>

        <div class="stat-box" style="grid-column:1/-1;">
            <h3>Division Summary</h3>
            ${divisionHTML}
        </div>
    `;
}

// ========================================
// FILTERS
// ========================================

document
.getElementById('filterInput')
.addEventListener('input', renderResults);

document
.getElementById('searchInput')
.addEventListener('input', renderResults);

// ========================================
// RENDER RESULTS
// ========================================

function renderResults(){

    const filterValue =
        document
        .getElementById('filterInput')
        .value
        .toLowerCase();

    const searchValue =
        document
        .getElementById('searchInput')
        .value
        .toLowerCase();

    let visibleResults =
        filteredResults.filter(r => {

            const matchesDivision =
                (r.division || "")
                .toLowerCase()
                .includes(filterValue);

            const matchesSearch =
                (r.url || "")
                .toLowerCase()
                .includes(searchValue)
                ||
                (r.pageName || "")
                .toLowerCase()
                .includes(searchValue);

            return matchesDivision && matchesSearch;
        });

    const start =
        (currentPage - 1) * pageSize;

    const paged =
        visibleResults.slice(
            start,
            start + pageSize
        );

    let html = `
        <table>

            <tr>
                <th>URL</th>
                <th>Division</th>
                <th>Sub Division</th>
                <th>Language</th>
                <th>Page Type</th>
                <th>Template</th>
                <th>Page Name</th>
                <th>Duplicate</th>
            </tr>
    `;

    paged.forEach(r => {

        html += `
            <tr>

                <td>${r.url}</td>

                <td class="${
                    r.division === "Not Found"
                    ? "not-found"
                    : ""
                }">
                    ${r.division || ""}
                </td>

                <td>${r.subDivision || ""}</td>

                <td>${r.language || ""}</td>

                <td>${r.pageType || ""}</td>

                <td>${r.templateName || ""}</td>

                <td>${r.pageName || ""}</td>

                <td class="${
                    r.duplicate
                    ? "duplicate"
                    : ""
                }">
                    ${r.duplicate ? "Yes" : "No"}
                </td>

            </tr>
        `;
    });

    html += `</table>`;

    document.getElementById('results').innerHTML = html;

    renderPagination(
        visibleResults.length
    );
}

// ========================================
// PAGINATION
// ========================================

function renderPagination(total){

    const totalPages =
        Math.ceil(total / pageSize);

    const pagination =
        document.getElementById('pagination');

    pagination.innerHTML = "";

    if(totalPages <= 1) return;

    for(let i=1; i<=totalPages; i++){

        const btn =
            document.createElement('button');

        btn.innerText = i;

        if(i === currentPage){

            btn.disabled = true;
        }

        btn.addEventListener('click', () => {

            currentPage = i;

            renderResults();
        });

        pagination.appendChild(btn);
    }
}

// ========================================
// CLEAR
// ========================================

document
.getElementById('clearBtn')
.addEventListener('click', () => {

    document.getElementById('urlInput').value = "";

    document.getElementById('results').innerHTML = "";

    document.getElementById('pagination').innerHTML = "";

    document.getElementById('filterInput').value = "";

    document.getElementById('searchInput').value = "";

    document.getElementById('stats').innerHTML = "";

    filteredResults = [];
});

// ========================================
// DOWNLOAD EXCEL
// ========================================

document
.getElementById('downloadExcelBtn')
.addEventListener('click', () => {

    if(filteredResults.length === 0){

        alert("No results");

        return;
    }

    const exportData =
        filteredResults.map(r => ({

            URL: r.url,

            Division: r.division,

            "Sub Division":
                r.subDivision,

            Language:
                r.language,

            "Page Type":
                r.pageType,

            Template:
                r.templateName,

            "Page Name":
                r.pageName,

            Duplicate:
                r.duplicate
                ? "Yes"
                : "No"
        }));

    const worksheet =
        XLSX.utils.json_to_sheet(exportData);

    const workbook =
        XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        "Results"
    );

    XLSX.writeFile(
        workbook,
        "URL_Division_Mapping.xlsx"
    );
});

// ========================================
// DOWNLOAD CSV
// ========================================

document
.getElementById('downloadCsvBtn')
.addEventListener('click', () => {

    if(filteredResults.length === 0){

        alert("No results");

        return;
    }

    let csv =
`URL,Division,Sub Division,Language,Page Type,Template,Page Name,Duplicate\n`;

    filteredResults.forEach(r => {

        csv += `"${r.url}","${r.division}","${r.subDivision}","${r.language}","${r.pageType}","${r.templateName}","${r.pageName}","${r.duplicate ? "Yes" : "No"}"\n`;
    });

    const blob =
        new Blob([csv], {
            type:'text/csv;charset=utf-8;'
        });

    const link =
        document.createElement('a');

    link.href =
        URL.createObjectURL(blob);

    link.download =
        "URL_Division_Mapping.csv";

    link.click();
});
