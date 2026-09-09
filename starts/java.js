/* =========================================================
   ZYALO COLLECTION DASHBOARD
   Excel + Chart.js
========================================================= */


/* =========================================================
   GLOBAL
========================================================= */

let excelData = [];

let executiveData = [];

let executiveChart = null;
let sanctionChart = null;
let percentageChart = null;



/* =========================================================
   HELPERS
========================================================= */

function cleanName(value) {

    if (value === null || value === undefined) {
        return "";
    }

    return String(value).trim();

}


function numberValue(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {
        return 0;
    }

    if (typeof value === "number") {
        return isNaN(value) ? 0 : value;
    }

    let text = String(value)
        .replace(/₹/g, "")
        .replace(/,/g, "")
        .replace(/%/g, "")
        .trim();

    let number = parseFloat(text);

    return isNaN(number) ? 0 : number;

}


function money(value) {

    value = numberValue(value);

    if (Math.abs(value) >= 10000000) {

        return "₹" +
            (value / 10000000).toFixed(2) +
            " Cr";

    }

    if (Math.abs(value) >= 100000) {

        return "₹" +
            (value / 100000).toFixed(2) +
            " L";

    }

    return "₹" +
        value.toLocaleString("en-IN", {
            maximumFractionDigits: 0
        });

}


function moneyFull(value) {

    return "₹" +
        numberValue(value).toLocaleString(
            "en-IN",
            {
                maximumFractionDigits: 0
            }
        );

}


/* =========================================================
   FIND COLUMN
========================================================= */

function getColumn(row, possibleNames) {

    for (const name of possibleNames) {

        if (
            Object.prototype.hasOwnProperty.call(
                row,
                name
            )
        ) {

            return row[name];

        }

    }

    return 0;

}


/* =========================================================
   LOAD EXCEL
========================================================= */

async function loadExcelData() {

    const status =
        document.getElementById("dataStatus");


    try {

        status.textContent =
            "Loading Excel data...";


        /*
         IMPORTANT:
         Excel file must be in same folder
         as collection.html
        */

        const response =
            await fetch(
                "ZYALO CVS MASTER.xlsx"
            );


        if (!response.ok) {

            throw new Error(
                "Excel file not found. HTTP " +
                response.status
            );

        }


        const buffer =
            await response.arrayBuffer();


        const workbook =
            XLSX.read(
                buffer,
                {
                    type: "array"
                }
            );


        const firstSheet =
            workbook.Sheets[
                workbook.SheetNames[0]
            ];


        excelData =
            XLSX.utils.sheet_to_json(
                firstSheet,
                {
                    defval: ""
                }
            );


        /*
         Clean column names
        */

        excelData =
            excelData.map(row => {

                const cleanRow = {};

                Object.keys(row).forEach(key => {

                    cleanRow[
                        key.trim()
                    ] = row[key];

                });

                return cleanRow;

            });


        if (excelData.length === 0) {

            throw new Error(
                "Excel file contains no data."
            );

        }


        status.textContent =
            excelData.length.toLocaleString(
                "en-IN"
            ) +
            " records loaded successfully";


        status.classList.add(
            "success-status"
        );


        calculateDashboard();


    } catch (error) {

        console.error(
            "Excel Error:",
            error
        );


        status.textContent =
            "❌ " + error.message;


        status.classList.add(
            "error-status"
        );

    }

}



/* =========================================================
   CALCULATE DASHBOARD
========================================================= */

function calculateDashboard() {


    let totalCases =
        excelData.length;


    let totalRepay = 0;

    let totalReceived = 0;

    let totalPending = 0;


    const executives = {};


    excelData.forEach(row => {


        /*
         REPAY AMOUNT
        */

        const repay =
            numberValue(
                getColumn(
                    row,
                    [
                        "Repay amt",
                        "Repay Amount",
                        "EXACT REPAY (PRIN+INT)",
                        "Total Repay amt."
                    ]
                )
            );


        /*
         RECEIVED
        */

        const received =
            numberValue(
                getColumn(
                    row,
                    [
                        "Total Received",
                        "Total Received ",
                        "Received Amount",
                        "Actual Repay amount"
                    ]
                )
            );


        /*
         PENDING
        */

        let pending =
            numberValue(
                getColumn(
                    row,
                    [
                        "Total O/S",
                        "Pending Amount",
                        "Repayment O/s (Pri+Int)"
                    ]
                )
            );


        /*
         If pending is not available,
         calculate it.
        */

        if (
            pending === 0 &&
            repay > received
        ) {

            pending =
                repay - received;

        }


        /*
         SANCTION
        */

        const sanction =
            numberValue(
                getColumn(
                    row,
                    [
                        "Loan Amount",
                        "Sanction",
                        "Sanction Amount"
                    ]
                )
            );


        totalRepay += repay;

        totalReceived += received;

        totalPending += pending;


        /*
         EXECUTIVE
        */

        let executive =
            cleanName(
                getColumn(
                    row,
                    [
                        "Sales Empl.",
                        "Sales Empl",
                        "Executive",
                        "Sales Employee"
                    ]
                )
            );


        if (!executive) {

            executive =
                "Unassigned";

        }


        if (!executives[executive]) {

            executives[executive] = {

                cases: 0,

                sanction: 0,

                repay: 0,

                received: 0,

                pending: 0

            };

        }


        executives[executive].cases++;

        executives[executive].sanction += sanction;

        executives[executive].repay += repay;

        executives[executive].received += received;

        executives[executive].pending += pending;


    });


    /*
     Update KPIs
    */

    updateElement(
        "totalCases",
        totalCases.toLocaleString("en-IN")
    );


    updateElement(
        "repayAmount",
        money(totalRepay)
    );


    updateElement(
        "receivedAmount",
        money(totalReceived)
    );


    updateElement(
        "pendingAmount",
        money(totalPending)
    );


    /*
     Collection %
    */

    let collectionPercentage = 0;

    if (totalRepay > 0) {

        collectionPercentage =
            (
                totalReceived /
                totalRepay
            ) * 100;

    }


    /*
     Summary
    */

    updateElement(
        "overviewCases",
        totalCases.toLocaleString("en-IN")
    );


    updateElement(
        "overviewReceived",
        money(totalReceived)
    );


    updateElement(
        "overviewCollection",
        collectionPercentage.toFixed(1) + "%"
    );


    updateElement(
        "overviewPending",
        money(totalPending)
    );


    /*
     Convert executive object to array
    */

    executiveData =
        Object.keys(executives)
            .map(name => {

                const item =
                    executives[name];


                const percentage =
                    item.repay > 0
                        ? (
                            item.received /
                            item.repay
                        ) * 100
                        : 0;


                return {

                    name: name,

                    cases: item.cases,

                    sanction: item.sanction,

                    repay: item.repay,

                    received: item.received,

                    pending: item.pending,

                    percentage:
                        percentage

                };

            });


    /*
     Sort by received amount
    */

    executiveData.sort(
        (a, b) =>
            b.received -
            a.received
    );


    createCharts();

    createTable();

}



/* =========================================================
   UPDATE ELEMENT
========================================================= */

function updateElement(
    id,
    value
) {

    const element =
        document.getElementById(id);


    if (element) {

        element.textContent =
            value;

    }

}



/* =========================================================
   CHART DEFAULTS
========================================================= */

function chartDefaults() {

    return {

        responsive: true,

        maintainAspectRatio: false,

        animation: {

            duration: 900

        },

        plugins: {

            legend: {

                position: "bottom",

                labels: {

                    usePointStyle: true,

                    padding: 18,

                    font: {

                        size: 12

                    }

                }

            },

            tooltip: {

                backgroundColor:
                    "#172f68",

                titleFont: {

                    size: 13

                },

                bodyFont: {

                    size: 12

                },

                padding: 12,

                cornerRadius: 8

            }

        }

    };

}



/* =========================================================
   CREATE CHARTS
========================================================= */

function createCharts() {


    /*
     Destroy old charts
    */

    if (executiveChart) {

        executiveChart.destroy();

    }

    if (sanctionChart) {

        sanctionChart.destroy();

    }

    if (percentageChart) {

        percentageChart.destroy();

    }


    /*
     Show maximum 12 executives
     */

    const chartData =
        executiveData.slice(
            0,
            12
        );


    const names =
        chartData.map(
            item => item.name
        );


    /* =====================================================
       1. EXECUTIVE COLLECTION
    ===================================================== */

    const executiveCanvas =
        document.getElementById(
            "executiveCollectionChart"
        );


    if (executiveCanvas) {

        executiveChart =
            new Chart(
                executiveCanvas,
                {

                    type: "bar",

                    data: {

                        labels: names,

                        datasets: [

                            {

                                label:
                                    "Received",

                                data:
                                    chartData.map(
                                        item =>
                                            item.received
                                    ),

                                backgroundColor:
                                    "#16a36a",

                                borderRadius: 7,

                                borderSkipped: false

                            },

                            {

                                label:
                                    "Pending",

                                data:
                                    chartData.map(
                                        item =>
                                            item.pending
                                    ),

                                backgroundColor:
                                    "#ef5350",

                                borderRadius: 7,

                                borderSkipped: false

                            }

                        ]

                    },

                    options: {

                        ...chartDefaults(),

                        indexAxis: "y",

                        scales: {

                            x: {

                                beginAtZero: true,

                                grid: {

                                    color:
                                        "#edf0f5"

                                },

                                ticks: {

                                    callback:
                                        function(value) {

                                            return money(value);

                                        }

                                }

                            },

                            y: {

                                grid: {

                                    display: false

                                }

                            }

                        }

                    }

                }

            );

    }



    /* =====================================================
       2. SANCTION VS REPAY
    ===================================================== */

    const sanctionCanvas =
        document.getElementById(
            "sanctionRepayChart"
        );


    if (sanctionCanvas) {

        sanctionChart =
            new Chart(
                sanctionCanvas,
                {

                    type: "bar",

                    data: {

                        labels: names,

                        datasets: [

                            {

                                label:
                                    "Sanction",

                                data:
                                    chartData.map(
                                        item =>
                                            item.sanction
                                    ),

                                backgroundColor:
                                    "#6475f5",

                                borderRadius: 7

                            },

                            {

                                label:
                                    "Repay",

                                data:
                                    chartData.map(
                                        item =>
                                            item.repay
                                    ),

                                backgroundColor:
                                    "#f2a93b",

                                borderRadius: 7

                            }

                        ]

                    },

                    options: {

                        ...chartDefaults(),

                        scales: {

                            x: {

                                grid: {

                                    display: false

                                }

                            },

                            y: {

                                beginAtZero: true,

                                grid: {

                                    color:
                                        "#edf0f5"

                                },

                                ticks: {

                                    callback:
                                        function(value) {

                                            return money(value);

                                        }

                                }

                            }

                        }

                    }

                }

            );

    }



    /* =====================================================
       3. COLLECTION PERCENTAGE
    ===================================================== */

    const percentageCanvas =
        document.getElementById(
            "collectionPercentageChart"
        );


    if (percentageCanvas) {

        percentageChart =
            new Chart(
                percentageCanvas,
                {

                    type: "bar",

                    data: {

                        labels: names,

                        datasets: [

                            {

                                label:
                                    "Collection %",

                                data:
                                    chartData.map(
                                        item =>
                                            Number(
                                                item.percentage.toFixed(1)
                                            )
                                    ),

                                backgroundColor:
                                    "#536dff",

                                borderRadius: 8,

                                barThickness: 25

                            }

                        ]

                    },

                    options: {

                        ...chartDefaults(),

                        indexAxis: "y",

                        plugins: {

                            ...chartDefaults().plugins,

                            legend: {

                                display: false

                            },

                            tooltip: {

                                callbacks: {

                                    label:
                                        function(context) {

                                            return (
                                                " Collection: " +
                                                context.raw +
                                                "%"
                                            );

                                        }

                                }

                            }

                        },

                        scales: {

                            x: {

                                beginAtZero: true,

                                max: 100,

                                grid: {

                                    color:
                                        "#edf0f5"

                                },

                                ticks: {

                                    callback:
                                        function(value) {

                                            return value + "%";

                                        }

                                }

                            },

                            y: {

                                grid: {

                                    display: false

                                }

                            }

                        }

                    }

                }

            );

    }

}



/* =========================================================
   CREATE TABLE
========================================================= */

function createTable() {


    const body =
        document.getElementById(
            "collectionBody"
        );


    if (!body) {

        return;

    }


    body.innerHTML = "";


    executiveData.forEach(
        (item, index) => {


            const row =
                document.createElement(
                    "tr"
                );


            let percentageClass =
                "percentage-low";


            if (item.percentage >= 80) {

                percentageClass =
                    "percentage-high";

            }
            else if (
                item.percentage >= 50
            ) {

                percentageClass =
                    "percentage-medium";

            }


            row.innerHTML = `

                <td>
                    <span class="rank">
                        ${index + 1}
                    </span>
                </td>

                <td>
                    <strong class="executive-name">
                        ${escapeHTML(item.name)}
                    </strong>
                </td>

                <td>
                    ${item.cases.toLocaleString("en-IN")}
                </td>

                <td>
                    ${money(item.sanction)}
                </td>

                <td>
                    ${money(item.repay)}
                </td>

                <td class="received-text">
                    ${money(item.received)}
                </td>

                <td class="pending-text">
                    ${money(item.pending)}
                </td>

                <td>

                    <span class="percentage ${percentageClass}">
                        ${item.percentage.toFixed(1)}%
                    </span>

                </td>

            `;


            body.appendChild(row);

        }

    );

}



/* =========================================================
   HTML ESCAPE
========================================================= */

function escapeHTML(value) {

    return String(value)

        .replace(/&/g, "&amp;")

        .replace(/</g, "&lt;")

        .replace(/>/g, "&gt;")

        .replace(/"/g, "&quot;")

        .replace(/'/g, "&#039;");

}



/* =========================================================
   START
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    function() {

        loadExcelData();

    }
);