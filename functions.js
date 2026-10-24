const AMFI_DATA_URL =
    "https://aromal-arow.github.io/amfi-excel-addin/data/NAVAll.txt";

let amfiCache = null;

async function loadAMFIData() {

    if (amfiCache) {
        return amfiCache;
    }

    const response = await fetch(AMFI_DATA_URL);

    if (!response.ok) {
        throw new Error("HTTP " + response.status);
    }

    const text = await response.text();
    const funds = {};

    const lines = text.split(/\r?\n/);

    for (const line of lines) {

        const parts = line.split(";");

        if (parts.length >= 6) {

            const code = parts[0].trim();

            if (/^\d+$/.test(code)) {

                funds[code] = {
                    code: code,
                    isin1: (parts[1] || "").trim(),
                    isin2: (parts[2] || "").trim(),
                    name: (parts[3] || "").trim(),
                    nav: (parts[4] || "").trim(),
                    date: (parts[5] || "").trim()
                };

            }
        }
    }

    amfiCache = funds;
    return funds;
}


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
