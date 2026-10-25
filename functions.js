const AMFI_DATA_URL =
    "https://aromal-arow.github.io/amfi-excel-addin/data/NAVAll.txt";

let amfiCache = null;


/**
 * Load the latest AMFI data from GitHub Pages.
 */
async function loadAMFIData() {

    if (amfiCache !== null) {
        return amfiCache;
    }

    const response = await fetch(AMFI_DATA_URL);

    if (!response.ok) {
        throw new Error(
            "Unable to load AMFI data. HTTP " + response.status
        );
    }

    const text = await response.text();

    const funds = {};

    const lines = text.split(/\r?\n/);

    for (const line of lines) {

        const parts = line.split(";");

        /*
         * AMFI format:
         *
         * 0 = Scheme Code
         * 1 = ISIN 1
         * 2 = ISIN 2
         * 3 = Scheme Name
         * 4 = Plan
         * 5 = Option
         * 6 = NAV
         * 7 = Date
         */

        if (parts.length >= 8) {

            const code = parts[0].trim();

            if (/^\d+$/.test(code)) {

                funds[code] = {
                    code: code,
                    isin1: (parts[1] || "").trim(),
                    isin2: (parts[2] || "").trim(),
                    name: (parts[3] || "").trim(),
                    plan: (parts[4] || "").trim(),
                    option: (parts[5] || "").trim(),
                    nav: (parts[6] || "").trim(),
                    date: (parts[7] || "").trim()
                };

            }
        }
    }

    amfiCache = funds;

    return funds;
}


/**
 * Returns the latest NAV for an AMFI scheme code.
 */
async function nav(schemeCode) {

    try {

        const funds = await loadAMFIData();

        const code = String(
            Math.trunc(Number(schemeCode))
        );

        const fund = funds[code];

        if (!fund) {
            return -404;
        }

        const value = Number(fund.nav);

        if (!Number.isFinite(value)) {
            return -405;
        }

        return value;

    } catch (error) {

        return -500;

    }
}


CustomFunctions.associate("NAV", nav);
