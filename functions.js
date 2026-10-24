const AMFI_DATA_URL =
    "https://aromal-arow.github.io/amfi-excel-addin/data/NAVAll.txt";

let amfiCache = null;


/*
 * Load and parse AMFI data.
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
         * Current AMFI format:
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


/*
 * Find a fund using its AMFI scheme code.
 */
async function getFund(schemeCode) {

    const funds = await loadAMFIData();

    const code = String(
        Math.trunc(Number(schemeCode))
    );

    return funds[code] || null;
}


/*
 * AMFI.NAV
 */
async function nav(schemeCode) {

    const fund = await getFund(schemeCode);

    if (!fund) {
        throw new Error("Scheme code not found");
    }

    const value = Number(fund.nav);

    if (!Number.isFinite(value)) {
        throw new Error("NAV unavailable");
    }

    return value;
}


/*
 * AMFI.NAME
 */
async function name(schemeCode) {

    const fund = await getFund(schemeCode);

    if (!fund) {
        throw new Error("Scheme code not found");
    }

    return fund.name;
}


/*
 * AMFI.ISIN
 */
async function isin(schemeCode) {

    const fund = await getFund(schemeCode);

    if (!fund) {
        throw new Error("Scheme code not found");
    }

    if (
        fund.isin1 &&
        fund.isin1 !== "-"
    ) {
        return fund.isin1;
    }

    if (
        fund.isin2 &&
        fund.isin2 !== "-"
    ) {
        return fund.isin2;
    }

    return "N/A";
}


/*
 * AMFI.PLAN
 */
async function plan(schemeCode) {

    const fund = await getFund(schemeCode);

    if (!fund) {
        throw new Error("Scheme code not found");
    }

    return fund.plan;
}


/*
 * AMFI.OPTION
 */
async function option(schemeCode) {

    const fund = await getFund(schemeCode);

    if (!fund) {
        throw new Error("Scheme code not found");
    }

    return fund.option;
}


/*
 * AMFI.DATE
 */
async function navDate(schemeCode) {

    const fund = await getFund(schemeCode);

    if (!fund) {
        throw new Error("Scheme code not found");
    }

    return fund.date;
}


/*
 * Register functions with Excel.
 */
CustomFunctions.associate("NAV", nav);
CustomFunctions.associate("NAME", name);
CustomFunctions.associate("ISIN", isin);
CustomFunctions.associate("PLAN", plan);
CustomFunctions.associate("OPTION", option);
CustomFunctions.associate("DATE", navDate);
