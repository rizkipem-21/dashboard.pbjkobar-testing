let data = [];
let filteredData = [];
let tempSelected = [];
let activeFilter = [];
let sortState = { key: null, asc: true };

// FORMAT
function formatRupiah(angka) {
    return new Intl.NumberFormat("id-ID").format(angka || 0);
}

// LOAD
async function loadData() {
    const res = await fetch("data/rekap.json");
    data = await res.json();
    filteredData = [...data];

    renderTable();
    createDropdown();
}

// TABLE
function renderTable() {
    const tbody = document.getElementById("tableBody");

    tbody.innerHTML = filteredData.map(row => `
        <tr class="border-b table-row-hover">
            <td class="p-2">${row["Satuan Kerja"]}</td>
            <td class="p-2 text-right">${formatRupiah(row["Pagu Program"])}</td>
            <td class="p-2 text-right">${formatRupiah(row["RUP Penyedia"])}</td>
            <td class="p-2 text-right">${formatRupiah(row["RUP Swakelola"])}</td>
            <td class="p-2 text-right font-bold">${formatRupiah(row["Total RUP Terumumkan"])}</td>
            <td class="p-2 text-right">${row["Persentase"].toFixed(2)}%</td>
        </tr>
    `).join("");
}

// DROPDOWN
function createDropdown() {
    const unique = [...new Set(data.map(d => d["Satuan Kerja"]))].sort();
    tempSelected = [...unique];

    renderDropdown(unique);
}

// RENDER DROPDOWN
function renderDropdown(list) {
    const el = document.getElementById("dropdown-nama");

    el.innerHTML = `
        <input id="searchDropdown" placeholder="Cari..." 
            class="w-full p-1 mb-2 border rounded text-sm">

        <label class="block border-b pb-1 mb-1">
            <input type="checkbox" id="selectAll" checked> <b>Select All</b>
        </label>

        <div id="checkboxList">
        ${list.map(val => `
            <label class="block">
                <input type="checkbox" class="cb" value="${val}" checked> ${val}
            </label>
        `).join("")}
        </div>

        <button onclick="applyFilter()" 
            class="mt-2 w-full bg-blue-600 text-white p-1 rounded">
            Terapkan Filter
        </button>
    `;

    // SEARCH
    document.getElementById("searchDropdown").addEventListener("input", e => {
        const keyword = e.target.value.toLowerCase();
        const filtered = list.filter(v => v.toLowerCase().includes(keyword));
        renderDropdown(filtered);
    });

    // SELECT ALL
    document.getElementById("selectAll").addEventListener("change", function() {
        document.querySelectorAll(".cb").forEach(cb => cb.checked = this.checked);
    });
}

// APPLY FILTER (TIDAK REALTIME)
function applyFilter() {
    const selected = [...document.querySelectorAll(".cb:checked")]
        .map(cb => cb.value);

    activeFilter = selected;

    filteredData = data.filter(row =>
        activeFilter.includes(row["Satuan Kerja"])
    );

    renderTable();
}

// SORT
function sortData(key) {
    if (sortState.key === key) {
        sortState.asc = !sortState.asc;
    } else {
        sortState.key = key;
        sortState.asc = true;
    }

    filteredData.sort((a, b) => {
        return sortState.asc
            ? a[key] - b[key]
            : b[key] - a[key];
    });

    renderTable();
}

// TOGGLE
function toggleDropdown() {
    document.getElementById("dropdown-nama").classList.toggle("show");
}

// INIT
loadData();