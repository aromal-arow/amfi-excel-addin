const AMFI_DATA_URL =
    "https://aromal-arow.github.io/amfi-excel-addin/data/NAVAll.txt";

let amfiCache = null;

async function loadAMFIData() {

    if (amfiCache !== null) {
        return amfiCache;
    }

    const response = await fetch(AMFI_DATA_URL);

    if (!response.ok) {
        throw new Error("Unable to load AMFI data");
    }

    const text = await response.text();

    const funds = new Map();

    const lines = text.split(/\r?\n/);

    for (const line of lines) {

        const parts = line.split(";");

        // Actual scheme rows begin with a numeric scheme code
        if (parts.length >= 5 && /^\d+$/.test(parts[0].trim())) {

            const schemeCode = parts[0].trim();

            funds.set(schemeCode, {
                code: schemeCode,
                isin1: (parts[1] || "").trim(),
                isin2: (parts[2] || "").trim(),
                name: (parts[3] || "").trim(),
                nav: (parts[4] || "").trim(),
                date: (parts[5] || "").trim()
            });
        }
    }

    amfiCache = funds;

    return funds;
}


/**
 * Returns the latest NAV for an AMFI scheme code.
 *
 * @customfunction NAV
 * @param {number} schemeCode AMFI scheme code.
 * @returns {Promise<number>} Latest NAV.
 */
async function nav(schemeCode) {

    const funds = await loadAMFIData();

    const code = String(Math.trunc(schemeCode));

    const fund = funds.get(code);

    if (!fund) {
        throw new Error("Scheme code not found: " + code);
    }

    const value = parseFloat(fund.nav);

    if (Number.isNaN(value)) {
        throw new Error("NAV unavailable for scheme: " + code);
    }

    return value;
}

CustomFunctions.associate("NAV", nav);
