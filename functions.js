const AMFI_DATA_URL =
    "https://aromal-arow.github.io/amfi-excel-addin/data/NAVAll.txt";

let amfiCache = null;


/*
 * Download and parse AMFI NAV data.
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
         * AMFI format
         *
         * 0 = Scheme Code
         * 1 = ISIN 1
         * 2 = ISIN 2
         * 3 = Scheme Name
         * 4 = Plan
         * 5 = Option
         * 6 = NAV
         * 7 = NAV Date
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
 * Get fund by AMFI scheme code.
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

    try {

        const fund = await getFund(schemeCode);

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


/*
 * AMFI.SCHEMENAME
 */
async function schemeName(schemeCode) {

    try {

        const fund = await getFund(schemeCode);

        if (!fund) {
            return "Scheme not found";
        }

        return fund.name;

    } catch (error) {

        return "Unable to load AMFI data";
    }
}


/*
 * AMFI.ISIN
 */
async function isin(schemeCode) {

    try {

        const fund = await getFund(schemeCode);

        if (!fund) {
            return "Scheme not found";
        }

        if (fund.isin1 && fund.isin1 !== "-") {
            return fund.isin1;
        }

        if (fund.isin2 && fund.isin2 !== "-") {
            return fund.isin2;
        }

        return "N/A";

    } catch (error) {

        return "Unable to load AMFI data";
    }
}


/*
 * AMFI.PLANTYPE
 */
async function planType(schemeCode) {

    try {

        const fund = await getFund(schemeCode);

        if (!fund) {
            return "Scheme not found";
        }

        return fund.plan;

    } catch (error) {

        return "Unable to load AMFI data";
    }
}


/*
 * AMFI.OPTIONTYPE
 */
async function optionType(schemeCode) {

    try {

        const fund = await getFund(schemeCode);

        if (!fund) {
            return "Scheme not found";
        }

        return fund.option;

    } catch (error) {

        return "Unable to load AMFI data";
    }
}


/*
 * AMFI.NAVDATE
 */
async function navDate(schemeCode) {

    try {

        const fund = await getFund(schemeCode);

        if (!fund) {
            return "Scheme not found";
        }

        return fund.date;

    } catch (error) {

        return "Unable to load AMFI data";
    }
}


/*
 * Register every function individually.
 */
CustomFunctions.associate("NAV", nav);
CustomFunctions.associate("SCHEMENAME", schemeName);
CustomFunctions.associate("ISIN", isin);
CustomFunctions.associate("PLANTYPE", planType);
CustomFunctions.associate("OPTIONTYPE", optionType);
CustomFunctions.associate("NAVDATE", navDate);
